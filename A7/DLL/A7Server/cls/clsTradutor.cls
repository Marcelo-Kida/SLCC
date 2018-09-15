VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTradutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF76ADE00BC"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pela tradução de mensagens no sistema A7.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private xmlRegraSaidaA7                     As MSXML2.DOMDocument40

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private ObjectContext                       As COMSVCSLib.ObjectContext

'Variavel auxiliar para o tradutor de mensagens
Private blnUtilizarRepeticao                As Boolean

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
    Set xmlRegraSaidaA7 = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
    
End Sub

'Rotina publica para tradução de mensagens. Recebe e devolve a mensagem dentro de um XML.
Public Function TraduzirMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                        Optional ByVal pstrNomeFila As String) As Boolean

Dim objRegraTransporte                      As A7Server.clsRegraTransporte
Dim objFilaErro                             As A7Server.clsFilaErro
Dim objXml                                  As MSXML2.DOMDocument40
Dim objMensagemEntrada                      As MSXML2.DOMDocument40
Dim objMensagemTraduzida                    As MSXML2.DOMDocument40
Dim objRegra                                As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strMensagemTraduzida                    As String
Dim strErroValidacao                        As String
Dim strDetalheOcorrenciaErro                As String
Dim strProtocolo                            As String
Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

'Objetos para validação com XSD
Dim objDOMSchema                            As MSXML2.DOMDocument40
Dim objSchemaCache                          As MSXML2.XMLSchemaCache40

Dim strHeaderNZ                             As String
Dim xmlMensagemTraduzida                    As MSXML2.DOMDocument40

Dim strParametroGeral                       As String
Dim xmlParametroGeral                       As MSXML2.DOMDocument40
    
    On Error GoTo ErrorHandler
    
    strProtocolo = Mid(pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR").Text, 1, Len(udtProtocolo))
    
    udtProtocoloAux.String = strProtocolo
    LSet udtProtocolo = udtProtocoloAux
    
    Set objRegraTransporte = CreateObject("A7Server.clsRegraTransporte")
    
    Set objRegra = CreateObject("MSXML2.DOMDocument.4.0")
    
    objRegra.validateOnParse = False
    objRegra.resolveExternals = False
       
    objRegra.loadXML objRegraTransporte.ObterInformacoesRegra(pxmlMensagem)
    
    Set objRegraTransporte = Nothing
    
    If objRegra.parseError.errorCode <> 0 Then
        'XML da regra de tradução inválido.
        lngCodigoErroNegocio = 1008
        GoTo ErrorHandler
    End If
    
    'Se existe regra de tradução
    If objRegra.documentElement.selectSingleNode("IN_EXIS_REGR_TRNF").Text = "S" Then
        
        If Trim(objRegra.documentElement.selectSingleNode("TX_REGR_TRNF_MESG").Text) = vbNullString Then
        
            strDetalheOcorrenciaErro = "Regra não cadastrada ou corrompida para a mensagem." & _
                                       "Tipo Mensagem   : " & pxmlMensagem.documentElement.selectSingleNode("//TP_MESG").Text & vbCrLf & _
                                       "Código Empresa  : " & pxmlMensagem.documentElement.selectSingleNode("//CO_EMPR_ORIG").Text & vbCrLf & _
                                       "Sistema Origem  : " & pxmlMensagem.documentElement.selectSingleNode("//SG_SIST_ORIG").Text & vbCrLf & _
                                       "Sistema Destino : " & pxmlMensagem.documentElement.selectSingleNode("//SG_SIST_DEST").Text
                                       
            pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetalheOcorrenciaErro)
            
            Set objRegra = Nothing
            Set objXml = Nothing
            Set objMensagemEntrada = Nothing
            
            Exit Function
            
        End If
        
        Set objXml = CreateObject("MSXML2.DOMDocument.4.0")
            
        fgAppendNode objXml, vbNullString, "Documento", vbNullString
        fgAppendNode objXml, "Documento", "Mensagem", vbNullString
        fgAppendAttribute objXml, "Mensagem", "Tipo", flTipoEntradaToSTR(CLng(objRegra.documentElement.selectSingleNode("TP_FORM_ENTR_MESG").Text))
        
        If CLng(objRegra.documentElement.selectSingleNode("TP_FORM_ENTR_MESG").Text) = enumTipoEntradaMensagem.EntradaCSV Then
            fgAppendAttribute objXml, "Mensagem", "Delimitador", objRegra.documentElement.selectSingleNode("TP_CTER_DELI_ENTR").Text
        Else
            fgAppendAttribute objXml, "Mensagem", "Delimitador", objRegra.documentElement.selectSingleNode("TP_CTER_DELI_SAID").Text
        End If
        
        'Se a mensagem de entrada for do tipo XML
        If CLng(objRegra.documentElement.selectSingleNode("TP_FORM_ENTR_MESG").Text) = enumTipoEntradaMensagem.EntradaXML Then
            'Entrada tipo XML. Fazer load.
            Set objMensagemEntrada = CreateObject("MSXML2.DOMDocument.4.0")
            
            objMensagemEntrada.validateOnParse = False
            objMensagemEntrada.resolveExternals = False
                       
            'Carrega a mensagem de entrada
            objMensagemEntrada.preserveWhiteSpace = False
            
            objMensagemEntrada.loadXML Mid(pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR").Text, Len(strProtocolo) + 1)
            
            If objMensagemEntrada.parseError.errorCode <> 0 Then
                'XML de entrada inválido
                TraduzirMensagem = False
                
                strDetalheOcorrenciaErro = "XML de entrada inválido." & vbCrLf & _
                                           "Parser Error Reason: " & objMensagemEntrada.parseError.reason
                
                pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetalheOcorrenciaErro)
                
                Set objRegra = Nothing
                Set objXml = Nothing
                Set objMensagemEntrada = Nothing
                
                Exit Function
            End If
                                        
            'Para mensagens MDR
            If UCase(objMensagemEntrada.childNodes(0).nodeTypeString) = UCase("element") Then
                objXml.selectSingleNode("//Documento/Mensagem").appendChild objMensagemEntrada.childNodes(0)
            Else
                Dim lngCont As Long
                
                For lngCont = 0 To objMensagemEntrada.childNodes.length
                    If UCase(objMensagemEntrada.childNodes(lngCont).nodeTypeString) = UCase("element") Then
                        objXml.selectSingleNode("//Documento/Mensagem").appendChild objMensagemEntrada.childNodes(lngCont)
                        Exit For
                    End If
                Next
            End If
            
            Set objMensagemEntrada = Nothing
            
        Else
            'Se a Mensagem de entrada for do Tipo CSV ou String
            objXml.selectSingleNode("//Documento/Mensagem").Text = Mid(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, Len(strProtocolo) + 1)
        End If
        
        fgAppendXML objXml, "Documento", objRegra.documentElement.selectSingleNode("TX_REGR_TRNF_MESG").Text
        
        If Not objXml.documentElement.selectSingleNode("Formato/@FilaDestino") Is Nothing Then
            fgAppendNode pxmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", objXml.documentElement.selectSingleNode("Formato/@FilaDestino").Text
        Else
            fgAppendNode pxmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", ""
        End If
        
        'Carrega Informações para validação XSD
        Set objDOMSchema = CreateObject("MSXML2.DOMDocument.4.0")
        Set objSchemaCache = CreateObject("MSXML2.XMLSchemaCache.4.0")
        Set objMensagemTraduzida = CreateObject("MSXML2.DOMDocument.4.0")
    
        objDOMSchema.async = False
    
        objDOMSchema.loadXML objRegra.selectSingleNode("//TX_VALID_SAID_MESG").Text
    
        objSchemaCache.Add "", objDOMSchema
        
        'Atribui o Validador XSD para o DOMDocument da mensagem Traduzida
        With objMensagemTraduzida
            
            .preserveWhiteSpace = True
            .async = True
            .validateOnParse = True
            .resolveExternals = False
            
            Set .schemas = objSchemaCache
            
            objXml.preserveWhiteSpace = True
            
            Set xmlRegraSaidaA7 = CreateObject("MSXML2.DOMDocument.4.0")
            If Dir(App.Path & "\RegraFormatacaoSaidaA7.xml") <> vbNullString Then
                If Not xmlRegraSaidaA7.Load(App.Path & "\RegraFormatacaoSaidaA7.xml") Then
                    Call fgAppendNode(xmlRegraSaidaA7, "", "RegraFormatacaoSaidaA7", "")
                End If
            End If
            
            'Se aqui funcionar tamo feito / senão deu problema
            strMensagemTraduzida = flTraduzir(objXml)
            
            Set xmlRegraSaidaA7 = Nothing
            
            'Carregar a mensagem traduzida no DOM para ser validada com XSD
            .loadXML strMensagemTraduzida
           
           'Verifica se ocorreu erro de validação
           If .parseError <> 0 Then
                '*************  ERRO TRADUÇÂO ***************************
                TraduzirMensagem = False
                
                strErroValidacao = .parseError.reason
                
                If InStr(1, strErroValidacao, "'") > 0 Then
                    strErroValidacao = Mid$(strErroValidacao, InStr(1, strErroValidacao, "'") + 1, Len(strErroValidacao))
                    strErroValidacao = StrReverse(strErroValidacao)
                    strErroValidacao = Mid$(strErroValidacao, InStr(1, strErroValidacao, "'") + 1, Len(strErroValidacao))
                    strErroValidacao = StrReverse(strErroValidacao)
                    
                    strDetalheOcorrenciaErro = "Campo : [ " & strErroValidacao & "] obrigatório ausente ou conteúdo do campo inválido."
                Else
                    If InStr(1, strErroValidacao, ".") > 0 Then
                        
                        If InStr(1, strErroValidacao, "Expecting:") > 0 Then
                            strErroValidacao = Mid$(strErroValidacao, InStr(1, strErroValidacao, ".") + 1, Len(strErroValidacao))
                            strErroValidacao = Mid$(strErroValidacao, InStr(1, strErroValidacao, ":") + 1, Len(strErroValidacao))
                            strErroValidacao = Replace(strErroValidacao, ".", vbNullString)
                            strErroValidacao = Replace(strErroValidacao, vbNullChar, vbNullString)
                            strErroValidacao = Replace(strErroValidacao, vbCr, vbNullString)
                            strErroValidacao = Replace(strErroValidacao, vbCrLf, vbNullString)
                            strErroValidacao = Replace(strErroValidacao, vbLf, vbNullString)
                            
                            strDetalheOcorrenciaErro = "Campo(s) obrigatório(s) ausente(s) : " & vbCrLf & strErroValidacao
                            
                        Else
                            strErroValidacao = Mid$(strErroValidacao, InStr(1, strErroValidacao, ".") + 1, Len(strErroValidacao))
                            strDetalheOcorrenciaErro = "Campo obrigatório ausente ou conteúdo do campo inválido. Error reason: " & vbCrLf & strErroValidacao
                        End If
                    Else
                        strErroValidacao = Mid$(strErroValidacao, 1, 130)
                        strDetalheOcorrenciaErro = "Campo obrigatório ausente ou conteúdo do campo inválido. Error reason: " & vbCrLf & strErroValidacao
                    End If
                    
                End If
                
                pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetalheOcorrenciaErro)
                                
                'Retorna a Mensagem Traduzida mas não valida
                pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_SAID").appendChild fgCreateCDATASection(strMensagemTraduzida)
                
                '*******************************************************************
                ' Adrian - Demanda de Reprocessamento de mensagens Canceladas no A7
                Set objFilaErro = CreateObject("A7Server.clsFilaErro")
                
                Call objFilaErro.PutFilaErro("A7Q.E.ENTRADA", _
                                             "A7Q.E.ERRO", _
                                             pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, _
                                             pxmlMensagem.selectSingleNode("//TX_DTLH_OCOR_ERRO").Text, _
                                             vbNullString, _
                                             True)
                
                Set objFilaErro = Nothing
                
           End If
           
        End With
        
        Set objDOMSchema = Nothing
        Set objSchemaCache = Nothing
        
        If strErroValidacao = "" Then
            'Verificar tipo de Saida do Evento
            'Caso tenha string ou CSV, concatenar apenas o Text do
            'Node pai do xml de string
            Select Case CLng(objRegra.documentElement.selectSingleNode("TP_FORM_SAID_MESG").Text)
                Case enumTipoSaidaMensagem.SaidaCSV
                    
                    strMensagemTraduzida = objMensagemTraduzida.selectSingleNode("//SaidaCSV").Text
                
                Case enumTipoSaidaMensagem.SaidaString
                                    
                    If UCase(Trim(udtProtocolo.SiglaSistemaDestino)) = "DV" Or _
                       UCase(Trim(udtProtocolo.SiglaSistemaDestino)) = "BG" Or _
                       UCase(Trim(udtProtocolo.SiglaSistemaDestino)) = "PZ" Then
                        
                        strMensagemTraduzida = objMensagemTraduzida.selectSingleNode("//SaidaSTR").Text
                    Else
                        strMensagemTraduzida = strProtocolo & objMensagemTraduzida.selectSingleNode("//SaidaSTR").Text
                    End If
                
                Case enumTipoSaidaMensagem.SaidaCSVXML
                   
                    strMensagemTraduzida = objMensagemTraduzida.selectSingleNode("//SaidaCSV").Text
                    
                    strMensagemTraduzida = strMensagemTraduzida & flMontaSaidaXML(strProtocolo, _
                                                                                  objRegra.documentElement.selectSingleNode("NO_TITU_MESG").Text, _
                                                                                  objMensagemTraduzida)
                    
                Case enumTipoSaidaMensagem.SaidaStringXML
                    
                    strMensagemTraduzida = objMensagemTraduzida.selectSingleNode("//SaidaSTR").Text
                    
                    strMensagemTraduzida = strMensagemTraduzida & flMontaSaidaXML(strProtocolo, _
                                                                                  objRegra.documentElement.selectSingleNode("NO_TITU_MESG").Text, _
                                                                                  objMensagemTraduzida)
                
                Case enumTipoSaidaMensagem.SaidaXML
                    
                    strMensagemTraduzida = flMontaSaidaXML(strProtocolo, _
                                                           objRegra.documentElement.selectSingleNode("NO_TITU_MESG").Text, _
                                                           objMensagemTraduzida)
            
                    Set xmlMensagemTraduzida = CreateObject("MSXML2.DOMDocument.4.0")
                    Call xmlMensagemTraduzida.loadXML(strMensagemTraduzida)
                
                    If Not xmlMensagemTraduzida.selectSingleNode("//TP_MESG") Is Nothing Then
                        If Val(xmlMensagemTraduzida.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.RetornoCompromissada Then
                            If Not xmlMensagemTraduzida.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                                If xmlMensagemTraduzida.selectSingleNode("//SG_SIST_ORIG").Text = "YS" Then
                                    If Not xmlMensagemTraduzida.selectSingleNode("//CO_EMPR") Is Nothing Then
                    
                                        strParametroGeral = fgSelectVarchar4000(0, False)
                                    
                                        Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
                                        xmlParametroGeral.preserveWhiteSpace = True
                                        xmlParametroGeral.loadXML strParametroGeral
                
                                        If Not xmlParametroGeral.selectSingleNode("//ALTERAR_EMPRESA_VOLTA_SAC") Is Nothing Then
                                            If xmlParametroGeral.selectSingleNode("//ALTERAR_EMPRESA_VOLTA_SAC/EMPRESA/@OBRIG").Text = "S" Then
                                                xmlMensagemTraduzida.selectSingleNode("//CO_EMPR").Text = xmlParametroGeral.selectSingleNode("//ALTERAR_EMPRESA_VOLTA_SAC/EMPRESA").Text
                                            End If
                                        End If
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    strMensagemTraduzida = xmlMensagemTraduzida.xml
                    Set xmlMensagemTraduzida = Nothing
            
            End Select
            
            'Retorna o Codigo de composicao atributo identificador emensagem
            pxmlMensagem.documentElement.selectSingleNode("CO_CMPO_ATRB_IDEF_MESG").Text = objMensagemTraduzida.selectSingleNode("//SaidaID").Text
                 
            If CLng(objRegra.documentElement.selectSingleNode("TP_NATZ_MESG").Text) = enumNaturezaMensagem.MensagemEnvio Then
                
                'Retorna a Mensagem Traduzida
                pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_SAID").appendChild fgCreateCDATASection(strMensagemTraduzida)
            
            ElseIf CLng(objRegra.documentElement.selectSingleNode("TP_NATZ_MESG").Text) = enumNaturezaMensagem.MensagemConsulta Then
                    
                'Retorna a Mensagem Consulta
                pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_SAID").appendChild fgCreateCDATASection(flRespostaMensagemConsulta(objMensagemTraduzida.selectSingleNode("//SaidaID").Text))
                            
            End If
            
            TraduzirMensagem = True
        
        End If
    
    ElseIf objRegra.documentElement.selectSingleNode("IN_EXIS_REGR_TRNF").Text = "N" Then
            
        If CLng(objRegra.documentElement.selectSingleNode("TP_NATZ_MESG").Text) = enumNaturezaMensagem.MensagemECO Then
            
            pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_SAID").appendChild fgCreateCDATASection(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text)
        
        ElseIf CLng(objRegra.documentElement.selectSingleNode("TP_NATZ_MESG").Text) = enumNaturezaMensagem.MensagemEnvio Then
               
            strMensagemTraduzida = flMontaMensagemSaidaSemTraducao(pxmlMensagem, _
                                                                   pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text)
                                                  
            If Val(pxmlMensagem.selectSingleNode("//TP_MESG").Text) = 1002 Then
                
                strHeaderNZ = Left$(strMensagemTraduzida, InStr(1, strMensagemTraduzida, "<") - 1)
                
                Set xmlMensagemTraduzida = CreateObject("MSXML2.DOMDocument.4.0")
                Call xmlMensagemTraduzida.loadXML(Replace$(strMensagemTraduzida, strHeaderNZ, vbNullString))
                Call TraduzBook2(xmlMensagemTraduzida, enumTipoTraducaoBookSPB.BookAntigo_Book2)
                
                strMensagemTraduzida = strHeaderNZ & xmlMensagemTraduzida.xml
                Set xmlMensagemTraduzida = Nothing
                
            End If
            
            pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_SAID").appendChild fgCreateCDATASection(strMensagemTraduzida)
                        
            pxmlMensagem.documentElement.selectSingleNode("CO_CMPO_ATRB_IDEF_MESG").Text = flObterCodigoMensagemNZ(pxmlMensagem.selectSingleNode("//TP_MESG").Text, _
                                                                                                                   pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text)
                        
            Set objXml = CreateObject("MSXML2.DOMDocument.4.0")
            objXml.loadXML objRegra.documentElement.selectSingleNode("TX_REGR_TRNF_MESG").Text
            
            If Not objXml.selectSingleNode("Formato/@FilaDestino") Is Nothing Then
                fgAppendNode pxmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", objXml.selectSingleNode("Formato/@FilaDestino").Text
            Else
                fgAppendNode pxmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", ""
            End If
            
            
        End If
        
        TraduzirMensagem = True
        
    End If
    
    Set objMensagemEntrada = Nothing
    Set objMensagemTraduzida = Nothing
    Set objRegra = Nothing
    Set objXml = Nothing
    
Exit Function
ErrorHandler:
    
    Set objDOMSchema = Nothing
    Set objSchemaCache = Nothing
    Set objRegraTransporte = Nothing
    Set objMensagemEntrada = Nothing
    Set objMensagemTraduzida = Nothing
    Set objRegra = Nothing
    Set objXml = Nothing
    Set objFilaErro = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "TraduzirMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Converte valores long para a string do domínio de "Tipo de Entrada" (Str / CSV / XML)
Private Function flTipoEntradaToSTR(ByVal plngTipoEntrada As Long) As String

    Select Case plngTipoEntrada
        Case enumTipoEntradaMensagem.EntradaCSV
            flTipoEntradaToSTR = "CSV"
        Case enumTipoEntradaMensagem.EntradaString
            flTipoEntradaToSTR = "String"
        Case enumTipoEntradaMensagem.EntradaXML
            flTipoEntradaToSTR = "XML"
    End Select

End Function

'Formata a saida de mensagens tipo string após a tradução.
'Esta rotina recebe a mensagem no formato do tradutor e formata de acordo com o esperado pelo sistema.
Private Function flMontaSaidaXML(ByVal pstrProtocolo As String, _
                                 ByVal pstrNomeTagPrincipal As String, _
                                 ByRef pxmlMensagemTraduzida As MSXML2.DOMDocument40) As String

Dim xmlMensagemSaida                        As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMensagem                             As A7Server.clsMensagem
Dim strMensagemSaida                        As String

On Error GoTo ErrorHandler
    
    Set xmlMensagemSaida = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagemSaida.preserveWhiteSpace = False
    xmlMensagemSaida.resolveExternals = False
        
    If Trim(pstrNomeTagPrincipal) = vbNullString Then
       pstrNomeTagPrincipal = "MESG"
    Else
       pstrNomeTagPrincipal = Trim(pstrNomeTagPrincipal)
    End If
        
    udtProtocoloAux.String = pstrProtocolo
    LSet udtProtocolo = udtProtocoloAux
        
    If IsNumeric(udtProtocolo.TipoMensagem) Then
        Call fgAppendNode(xmlMensagemSaida, "", pstrNomeTagPrincipal, "")
        Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, "TP_MESG", udtProtocolo.TipoMensagem)
        Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, "SG_SIST_ORIG", UCase(Trim(udtProtocolo.SiglaSistemaOrigem)))
        Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, "SG_SIST_DEST", UCase(Trim(udtProtocolo.SiglaSistemaDestino)))
        Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, "CO_EMPR", CLng(udtProtocolo.CodigoEmpresa))
           
        For Each xmlNode In pxmlMensagemTraduzida.selectNodes("//SaidaXML/*")
            Call fgAppendXML(xmlMensagemSaida, pstrNomeTagPrincipal, xmlNode.xml)
        Next
        
        If UCase(Trim(udtProtocolo.SiglaSistemaOrigem)) = "NZ" Then
            Call TraduzBook2(xmlMensagemSaida, enumTipoTraducaoBookSPB.Book2_BookAntigo)
        End If
        
        strMensagemSaida = xmlMensagemSaida.xml
    Else
        
        Call fgAppendNode(xmlMensagemSaida, "", "SISMSG", "")
        Call fgAppendNode(xmlMensagemSaida, "SISMSG", pstrNomeTagPrincipal, "")
    
        For Each xmlNode In pxmlMensagemTraduzida.selectNodes("//SaidaXML/*")
            
            If Mid(UCase(xmlNode.nodeName), 1, 2) = "DT" Or Mid(UCase(xmlNode.nodeName), 1, 2) = "DH" Then
                
                If xmlNode.Text <> "0" Then
                    'Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, xmlNode.nodeName, xmlNode.Text)
                    Call fgAppendXML(xmlMensagemSaida, pstrNomeTagPrincipal, xmlNode.xml)
                Else
                    Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, xmlNode.nodeName, "")
                End If
                
            Else
                'Call fgAppendNode(xmlMensagemSaida, pstrNomeTagPrincipal, xmlNode.nodeName, xmlNode.Text)
                Call fgAppendXML(xmlMensagemSaida, pstrNomeTagPrincipal, xmlNode.xml)
            End If
            
        Next
            
        If UCase(Trim(udtProtocolo.SiglaSistemaDestino)) = "NZ" Then
            Call TraduzBook2(xmlMensagemSaida, enumTipoTraducaoBookSPB.BookAntigo_Book2)
        End If
        
        
        
        strMensagemSaida = xmlMensagemSaida.xml
        
    End If
    
    flMontaSaidaXML = strMensagemSaida

    Set xmlMensagemSaida = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    Set xmlMensagemSaida = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaSaidaXML Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Retorna valores à mensagens do tipo consulta
Private Function flRespostaMensagemConsulta(ByVal pstrIdentificadorMensagem As String)

Dim xmlResposta                             As MSXML2.DOMDocument40
Dim Rs                                      As ADODB.Recordset
Dim strSQL                                  As String
Dim strProtocolo                            As String
Dim udtProtocolo                            As udtProtocolo
Dim strMensagemEntrada                      As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT A.CO_MESG,               " & _
                    " A.CO_CMPO_ATRB_IDEF_MESG," & _
                    " A.CO_TEXT_XML_ENTR,      " & _
                    " B.DE_ABRV_OCOR_MESG,     " & _
                    " B.DE_OCOR_MESG,          " & _
                    " C.NU_SEQU_SITU_MESG      " & _
            "   FROM  A7.TB_MESG      A ,      " & _
                    " A7.TB_OCOR_MESG B ,      " & _
                    " A7.TB_SITU_MESG C        " & _
            "  WHERE  C.CO_MESG = A.CO_MESG    " & _
            "    AND  C.CO_OCOR_MESG = B.CO_OCOR_MESG  " & _
            "    AND  A.CO_CMPO_ATRB_IDEF_MESG = '" & Trim(pstrIdentificadorMensagem) & "' " & _
            "    AND  C.NU_SEQU_SITU_MESG  =   " & _
            " (SELECT MAX(E.NU_SEQU_SITU_MESG) " & _
            "  FROM A7.TB_MESG      D ,        " & _
            "       A7.TB_SITU_MESG E          " & _
            " WHERE D.CO_MESG = E.CO_MESG      " & _
            "   AND D.CO_CMPO_ATRB_IDEF_MESG = '" & Trim(pstrIdentificadorMensagem) & "')"

    Set Rs = fgQuerySQL(strSQL)
        
    Set xmlResposta = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Rs.EOF Then
        Call fgAppendNode(xmlResposta, "", "MESG", "")
        Call fgAppendNode(xmlResposta, "MESG", "DE_OCOR_MESG", "Mensagem não Encontrada")
    Else
        
        strMensagemEntrada = fgSelectVarchar4000(Rs!CO_TEXT_XML_ENTR)
        strProtocolo = Mid(strMensagemEntrada, 1, Len(udtProtocolo))
        
        Call fgAppendNode(xmlResposta, "", "MESG", "")
        Call fgAppendNode(xmlResposta, "MESG", "TX_PRTC", strProtocolo)
        Call fgAppendNode(xmlResposta, "MESG", "CO_MESG", Rs!CO_MESG)
        Call fgAppendNode(xmlResposta, "MESG", "DE_OCOR_MESG", Rs!DE_OCOR_MESG)
        Call fgAppendNode(xmlResposta, "MESG", "DE_ABRV_OCOR_MESG", Trim(Rs!DE_ABRV_OCOR_MESG))
    
    End If
    
    flRespostaMensagemConsulta = xmlResposta.xml
    
    Set xmlResposta = Nothing
    
    Exit Function
ErrorHandler:
        
    Set xmlResposta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRespostaMensagemConsulta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Formata a mensagem de entrada de acordo com o formato de saida esperado pelo sistema.
'Não aplica tradução na mensagem recebida.
Private Function flMontaMensagemSaidaSemTraducao(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                                 ByVal pstrMensagemEntrada As String) As String
                                                 
Dim xmlMensagemSaidaSemTraducao                     As MSXML2.DOMDocument40
Dim udtProtocolo                                    As udtProtocolo
                                                 
On Error GoTo ErrorHandler
                                                 
    Set xmlMensagemSaidaSemTraducao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Trim(pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text) = "A8" Then
    
        With pxmlMensagem.documentElement
            Call fgAppendNode(xmlMensagemSaidaSemTraducao, "", "MESG", "")
            Call fgAppendNode(xmlMensagemSaidaSemTraducao, "MESG", "TP_MESG", .selectSingleNode("TP_MESG").Text)
            Call fgAppendNode(xmlMensagemSaidaSemTraducao, "MESG", "SG_SIST_ORIG", UCase(Trim(.selectSingleNode("SG_SIST_ORIG").Text)))
            Call fgAppendNode(xmlMensagemSaidaSemTraducao, "MESG", "SG_SIST_DEST", UCase(Trim(.selectSingleNode("SG_SIST_DEST").Text)))
            Call fgAppendNode(xmlMensagemSaidaSemTraducao, "MESG", "CO_EMPR", CLng(.selectSingleNode("CO_EMPR_ORIG").Text))
            Call fgAppendXML(xmlMensagemSaidaSemTraducao, "MESG", "<TX_MESG>" & pstrMensagemEntrada & "</TX_MESG>")
        End With
        
        If Trim(pxmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text) = "NZ" And _
           Trim(pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text) = "A8" Then
            Call TraduzBook2(xmlMensagemSaidaSemTraducao, Book2_BookAntigo)
        End If
        
        flMontaMensagemSaidaSemTraducao = xmlMensagemSaidaSemTraducao.xml
    Else
        
        flMontaMensagemSaidaSemTraducao = Mid(pstrMensagemEntrada, Len(udtProtocolo) + 1)
        
    End If
    
    Set xmlMensagemSaidaSemTraducao = Nothing
                                                 
    Exit Function
ErrorHandler:
                                                        
    Set xmlMensagemSaidaSemTraducao = Nothing
                                                        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaMensagemSaidaSemTraducao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
                                                 
End Function

'**************************************************************************************
'*** Nova implementação para o tradutor (com repetição)
'**************************************************************************************
'Prepara a mensagem recebida para a tradução
'(Obtem informações de regra e tipo do formato de saída)
Private Function flTraduzir(ByRef pxmlMensagem As DOMDocument40) As String
On Error GoTo ErrorHandler
Dim xmlNode                                 As IXMLDOMNode
Dim strTipoSaida                            As String
Dim lngTipoSaida                            As Long
Dim strTagFormato                           As String
Dim xmlDOMSaida                             As DOMDocument40

Dim lngX                                    As Long
Dim blnTraduzTipo                           As Boolean

    Set xmlDOMSaida = New DOMDocument40
    
    xmlDOMSaida.loadXML "<Saida></Saida>"
    
    If pxmlMensagem.selectSingleNode("//Documento/Mensagem/@Tipo").Text = "XML" Then
        flPrepararXML pxmlMensagem
    End If
    
    'Inicializa a variavel auxiliar para utilização de repetições
    blnUtilizarRepeticao = False
    
    'For Each xmlNode In pxmlMensagem.selectNodes("//Documento/Formato/*")
    For lngX = 1 To 4
        
        blnTraduzTipo = False
        
        Select Case lngX
            Case 1
                If pxmlMensagem.selectNodes("//Documento/Formato/IDOutPut").length > 0 Then
                    strTipoSaida = "SaidaID"
                    lngTipoSaida = enumTipoParteSaida.ParteId
                    strTagFormato = "IDOutPut"
                    blnTraduzTipo = True
                End If
            Case 2
                If pxmlMensagem.selectNodes("//Documento/Formato/STROutPut").length > 0 Then
                    strTipoSaida = "SaidaSTR"
                    lngTipoSaida = enumTipoParteSaida.ParteSTR
                    strTagFormato = "STROutPut"
                    blnTraduzTipo = True
                End If
            Case 3
                If pxmlMensagem.selectNodes("//Documento/Formato/CSVOutPut").length > 0 Then
                    strTipoSaida = "SaidaCSV"
                    lngTipoSaida = enumTipoParteSaida.ParteCSV
                    strTagFormato = "CSVOutPut"
                    blnTraduzTipo = True
                End If
            Case 4
                If pxmlMensagem.selectNodes("//Documento/Formato/XMLOutPut").length > 0 Then
                    strTipoSaida = "SaidaXML"
                    lngTipoSaida = enumTipoParteSaida.ParteXML
                    strTagFormato = "XMLOutPut"
                    blnTraduzTipo = True
                End If
        End Select
        
        If blnTraduzTipo Then
            fgAppendNode xmlDOMSaida, _
                         "Saida", strTipoSaida, vbNullString
                                 
            flSubConverter pxmlMensagem.selectSingleNode("//Documento/Mensagem"), _
                           pxmlMensagem.selectSingleNode("//" & strTagFormato), _
                           xmlDOMSaida.selectSingleNode("//" & strTipoSaida), _
                           lngTipoSaida
        End If
    Next

    flTraduzir = xmlDOMSaida.xml
    
    Set xmlDOMSaida = Nothing
    
    Exit Function
ErrorHandler:
    Set xmlDOMSaida = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTraduzir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Converte efetivamente a mensagem recebida.
'Recebe como parametros iniciais: Mensagem, Formato da Regra, Parte da SAida (ID, XML, STR...) e o destino da tradução (DOM)
Private Sub flSubConverter(ByRef pxmlMensagem As IXMLDOMNode, _
                           ByRef pxmlNodeFormato As IXMLDOMNode, _
                           ByRef pxmlNodeSaida As IXMLDOMNode, _
                           ByRef pTipoParteSaida As enumTipoParteSaida, _
                  Optional ByVal plngNumeroIteracao As Long)

Dim xmlNode                                 As IXMLDOMNode
Dim strValor                                As String
Dim strTagPai                               As String
Dim lngRepeticoes                           As Long
Dim lngX                                    As Long
Dim strXPath                                As String
Dim lngTamanhoPai                           As Long
Dim lngTamanhoAvo                           As Long
Dim lngIteracaoAvo                          As Long
Dim blnUsarNumeroInteracaoPai               As Boolean
Dim xmlNodeContexto                         As IXMLDOMNode
Dim blnLimparIndicesRepeticao               As Boolean

Dim xmlAttribute                            As MSXML2.IXMLDOMAttribute
Dim AUXMLnode                               As IXMLDOMNode

On Error GoTo ErrorHandler

    If pxmlMensagem Is Nothing Then Exit Sub
        
    For Each xmlNode In pxmlNodeFormato.selectNodes("./*")
        
        'Obter o valor do campo de saida
        
        Select Case pxmlMensagem.selectSingleNode("//Documento/Mensagem/@Tipo").Text
        
            Case "XML"
                If xmlNode.selectSingleNode("@Tipo").Text <> "Grupo" And _
                   xmlNode.selectSingleNode("@TargetTag").Text <> vbNullString Then
                    If InStr(1, xmlNode.parentNode.nodeName, "OutPut") = 0 Then
                        
                        strXPath = ".//" & xmlNode.selectSingleNode("@TargetTag").Text
                        
                        'Verificação para não impactar nas mensagens existentes
                        If Not xmlNode.selectSingleNode("@RepetTag") Is Nothing Then
                            If CBool(Val(xmlNode.selectSingleNode("@RepetTag").Text)) Then
                                'Se a tradução possui repetição utilziar posição
                                
                                'If Not xmlNode.parentNode.selectSingleNode("@UltimaPosicao") Is Nothing Then
                                '    'Utilizar a posicao da tag de repeticao (repet dentro de repet) para não sair do contexto
                                '    If CLng(Val(xmlNode.parentNode.selectSingleNode("@UltimaPosicao").Text)) > 0 Then
                                '
                                '        strXPath = ".//*[@Posicao='" & CLng(Val(xmlNode.parentNode.selectSingleNode("@UltimaPosicao").Text)) & "']//" & _
                                '                    Replace$(strXPath, ".//", "")
                                '    End If
                                'End If
                                
                                strXPath = strXPath & "[@Posicao>'" & xmlNode.selectSingleNode("@UltimaPosicao").Text & "']"
                                
                            End If
                        End If
                        
                        'Esta verificação é usada para itens que tem pai (possiveis repeticoes)
                        'Verifica se existe algum nó após a ultima prosicao pesquisada
                        If Not pxmlMensagem.selectSingleNode(strXPath) Is Nothing Then
                                
                            strValor = pxmlMensagem.selectSingleNode(strXPath).Text
                            
                            'Verificação para não impactar nas mensagens existentes
                            If Not xmlNode.selectSingleNode("@UltimaPosicao") Is Nothing Then
                                'Guarda a ultima psicao
                                xmlNode.selectSingleNode("@UltimaPosicao").Text = pxmlMensagem.selectSingleNode(strXPath & "/@Posicao").Text
                            End If
                            
                        Else
                            'Não existem mais nós.
                            strValor = vbNullString
                        End If
                    Else

                        'Primeiro Nível do XML
                        If Not pxmlMensagem.selectSingleNode(".//" & xmlNode.selectSingleNode("@TargetTag").Text) Is Nothing Then
                            strValor = pxmlMensagem.selectSingleNode(".//" & xmlNode.selectSingleNode("@TargetTag").Text).Text
                        Else
                            strValor = vbNullString
                        End If
                    End If
                Else
                    strValor = vbNullString
                End If
                
                If InStr(1, xmlNode.baseName, "CO_FORM_LIQU") = 0 And _
                   InStr(1, xmlNode.baseName, "CO_BANC") = 0 And _
                   InStr(1, xmlNode.baseName, "CO_AGEN") = 0 And _
                   InStr(1, xmlNode.baseName, "NU_CC") = 0 And _
                   InStr(1, xmlNode.baseName, "VA_") = 0 And _
                   InStr(1, xmlNode.baseName, "PU_") = 0 And _
                   InStr(1, xmlNode.baseName, "PE_") = 0 And _
                   InStr(1, xmlNode.baseName, "QT_") = 0 Then
                
                    If Not xmlNode.selectSingleNode("@Default") Is Nothing Then
                        If xmlNode.selectSingleNode("@DefaultObrigatorio").Text = "1" Then
                            strValor = xmlNode.selectSingleNode("@Default").Text
                        Else
                            If Trim$(xmlNode.selectSingleNode("@Default").Text) <> vbNullString Then
                                If Trim$(strValor) = vbNullString Then
                                    strValor = xmlNode.selectSingleNode("@Default").Text
                                Else
                                    If IsNumeric(strValor) Then
                                        If Val(strValor) = 0 Then
                                            strValor = xmlNode.selectSingleNode("@Default").Text
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            
            Case "String"
                
                If xmlNode.selectSingleNode("./@Tipo").Text = "Grupo" Then
                    'Tratamento para Item de grupo
                    strValor = vbNullString
                Else
                    
                    If xmlNode.parentNode.selectSingleNode("./@TamanhoOriginal") Is Nothing Then
                        lngTamanhoPai = 0
                    Else
                        lngTamanhoPai = CLng(Val(xmlNode.parentNode.selectSingleNode("./@TamanhoOriginal").Text))
                    End If
                                        
                    If xmlNode.parentNode.parentNode.parentNode.selectSingleNode("./@TamanhoOriginal") Is Nothing Then
                        lngTamanhoAvo = 0
                    Else
                        lngTamanhoAvo = CLng(Val(xmlNode.parentNode.parentNode.parentNode.selectSingleNode("./@TamanhoOriginal").Text))
                        If Not xmlNode.parentNode.parentNode.parentNode.parentNode.selectSingleNode("./@Iteracao") Is Nothing Then
                            lngIteracaoAvo = CLng(xmlNode.parentNode.parentNode.parentNode.parentNode.selectSingleNode("./@Iteracao").Text) - 1
                        Else
                            lngIteracaoAvo = 0
                        End If
                    End If
                    
                    If Trim(xmlNode.selectSingleNode("@Inicio").Text) <> vbNullString And _
                       Trim(xmlNode.selectSingleNode("@TamanhoOriginal").Text) <> vbNullString Then
                        
                        strValor = Mid$(pxmlMensagem.Text, _
                                   xmlNode.selectSingleNode("@Inicio").Text + (plngNumeroIteracao * lngTamanhoPai) + (lngIteracaoAvo * lngTamanhoAvo), _
                                   xmlNode.selectSingleNode("@TamanhoOriginal").Text)
                    Else
                        strValor = vbNullString
                    End If
                    
                End If
    
                If InStr(1, xmlNode.baseName, "CO_FORM_LIQU") = 0 And _
                   InStr(1, xmlNode.baseName, "CO_BANC") = 0 And _
                   InStr(1, xmlNode.baseName, "CO_AGEN") = 0 And _
                   InStr(1, xmlNode.baseName, "NU_CC") = 0 And _
                   InStr(1, xmlNode.baseName, "VA_") = 0 And _
                   InStr(1, xmlNode.baseName, "PU_") = 0 And _
                   InStr(1, xmlNode.baseName, "PE_") = 0 And _
                   InStr(1, xmlNode.baseName, "QT_") = 0 Then
                
                    If Not xmlNode.selectSingleNode("@Default") Is Nothing Then
                        If xmlNode.selectSingleNode("@DefaultObrigatorio").Text = "1" Then
                            strValor = xmlNode.selectSingleNode("@Default").Text
                        Else
                            If Trim$(xmlNode.selectSingleNode("@Default").Text) <> vbNullString Then
                                If Trim$(strValor) = vbNullString Then
                                    strValor = xmlNode.selectSingleNode("@Default").Text
                                Else
                                    If IsNumeric(strValor) Then
                                        If Val(strValor) = 0 Then
                                            strValor = xmlNode.selectSingleNode("@Default").Text
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            
            Case "CSV"
                strValor = flValueFromCSV(pxmlMensagem.Text, _
                                          pxmlMensagem.selectSingleNode("//Documento/Mensagem/@Delimitador").Text, _
                                          CLng(xmlNode.selectSingleNode("@Indice").Text), _
                                          xmlNode.selectSingleNode("@Tipo").Text, _
                                          CBool(xmlNode.selectSingleNode("@Obrigatorio").Text))
        End Select
        
        If pTipoParteSaida = enumTipoParteSaida.ParteId Or pTipoParteSaida = enumTipoParteSaida.ParteSTR Then
            
            If xmlNode.selectSingleNode("@Tipo").Text <> "Grupo" Then
                'Tratamento para item de grupo
                'formatar como CopyFixo
              
                    strValor = flValorToSTR(strValor, _
                                            xmlNode.selectSingleNode("@Tipo").Text, _
                                            CLng(xmlNode.selectSingleNode("@Tamanho").Text), _
                                            CLng(xmlNode.selectSingleNode("@Decimais").Text), _
                                            CBool(xmlNode.selectSingleNode("@Obrigatorio").Text), _
                                            xmlNode.nodeName)
            End If
        
        ElseIf pTipoParteSaida = enumTipoParteSaida.ParteXML Then
            
            If pxmlMensagem.selectSingleNode("//Documento/Mensagem/@Tipo").Text = "String" Then
                strValor = flValorToXML(strValor, _
                                        xmlNode.selectSingleNode("@Tipo").Text, _
                                        CLng(xmlNode.selectSingleNode("@Tamanho").Text), _
                                        CLng(xmlNode.selectSingleNode("@Decimais").Text), _
                                        CBool(xmlNode.selectSingleNode("@Obrigatorio").Text))
            End If
        
            If xmlNode.selectSingleNode("@Tipo").Text = "number" And CLng(xmlNode.selectSingleNode("@Tamanho").Text) <= 15 Then
                
                If Not xmlNode.selectSingleNode("@TargetTag") Is Nothing Then
                    Select Case Trim$(xmlNode.selectSingleNode("@TargetTag").Text)
                    
                        Case "CO_PARP_CAMR", "CO_CNPT_CAMR", "CO_CLIE_CAMR", _
                             "CO_BANC_LIQU_CAMR", "CO_ANUE_CAMR", _
                             "CO_CEDE_CAMR", "CO_ADQU_CAMR"
    
                        Case Else
                            'Retira os zeros a esqueda
                            If CLng(xmlNode.selectSingleNode("@Decimais").Text) = 0 Then
                                If IsNumeric(strValor) Then
                                    strValor = CDbl(strValor)
                                End If
                            End If
                            
                            'Pikachu - 03/11/2004
                            'Formatar conteudo tag de acordo com xml de apoio - RegraFormatacaoSaidaA7.xml

                            strValor = flFormataSaidaA7(pxmlMensagem, _
                                                         xmlNode.nodeName, _
                                                         strValor, _
                                                         CLng(xmlNode.selectSingleNode("@Tamanho").Text))
                    End Select
                Else
                    'Retira os zeros a esqueda
                    If CLng(xmlNode.selectSingleNode("@Decimais").Text) = 0 Then
                        If IsNumeric(strValor) Then
                            strValor = CDbl(strValor)
                        End If
                    End If
                    
                    'Pikachu - 03/11/2004
                    'Formatar conteudo tag de acordo com xml de apoio - RegraFormatacaoSaidaA7.xml
                    
                    strValor = flFormataSaidaA7(pxmlMensagem, _
                                                 xmlNode.nodeName, _
                                                 strValor, _
                                                 CLng(xmlNode.selectSingleNode("@Tamanho").Text))
                    
                End If
            End If
        End If
        
        'Montar mensagem de saida
        If xmlNode.selectSingleNode("child::*") Is Nothing Then
            
            If Not (pxmlMensagem.selectSingleNode("//Documento/Mensagem/@Tipo").Text = "String" And _
               UCase(xmlNode.selectSingleNode("@Tipo").Text) = "NUMBER" And _
               Not CBool(xmlNode.selectSingleNode("@Obrigatorio").Text) And _
               Trim$(strValor) = vbNullString) Then
                'Quando a tag for numérica e não obrigatória não colocar a tag na mensagem para entradas string(brancos)
            
                Set AUXMLnode = pxmlNodeSaida.ownerDocument.createElement(xmlNode.nodeName)
                AUXMLnode.Text = strValor
                pxmlNodeSaida.appendChild AUXMLnode
                Set AUXMLnode = Nothing
            
            End If
            
            If pTipoParteSaida = ParteCSV Then
                'Incluir Delimitador
                'xsl:variable name="Delimitador" select="//Documento/Mensagem/@Delimitador" />
                If xmlNode.nodeName <> pxmlNodeFormato.selectSingleNode("//*[position()=last()]").nodeName Then
                    fgAppendNode pxmlNodeSaida.ownerDocument, _
                                 pxmlNodeSaida.nodeName & "[position()=last()]", _
                                 "Separador", _
                                 pxmlNodeFormato.selectSingleNode("//Documento/Mensagem/@Delimitador").Text
                End If
            End If
            
        Else
            'Tag com Filho
            
            Set AUXMLnode = pxmlNodeSaida.ownerDocument.createElement(xmlNode.nodeName)
            pxmlNodeSaida.appendChild AUXMLnode
            
            'Inclui atributo para controle de iteracao - somente utilizado para string
            Set xmlAttribute = xmlNode.ownerDocument.createAttribute("Iteracao")
            xmlNode.attributes.setNamedItem xmlAttribute
            Set xmlAttribute = Nothing
            
            blnUsarNumeroInteracaoPai = False
            If CLng(xmlNode.selectSingleNode("@Repeticoes").Text) = 0 Then
                lngRepeticoes = 1
                blnUsarNumeroInteracaoPai = True
                
                'Deve-se limpar os indices pois o a não repetição é garantida pelo contexto do pai
                blnLimparIndicesRepeticao = True
            Else
                lngRepeticoes = CLng(xmlNode.selectSingleNode("@Repeticoes").Text)
                blnUtilizarRepeticao = True
                flZerarIndicesDaRepeticao xmlNode
                blnLimparIndicesRepeticao = False
            End If

            'Trecho inserido pois o sistema não está identificando a quantidade correta de repetições.
            'Ocorreu em produção com a CTP9015 enviada pelo A8 (Lançamento de PU).
            If lngRepeticoes = 1 Then
                If Not xmlNode.selectSingleNode("@TargetTag") Is Nothing Then
                    If xmlNode.selectSingleNode("@TargetTag").Text <> vbNullString Then
                        If pxmlMensagem.selectNodes("//" & xmlNode.selectSingleNode("@TargetTag").Text & "/*").length > lngRepeticoes And _
                           InStr(1, UCase(xmlNode.baseName), "REPET_") <> 0 And _
                           InStr(1, UCase(xmlNode.baseName), "GRUPO_") = 0 Then
        
                            lngRepeticoes = CLng(pxmlMensagem.selectNodes("//" & xmlNode.selectSingleNode("@TargetTag").Text & "/*").length)
                            blnUtilizarRepeticao = True
                            flZerarIndicesDaRepeticao xmlNode
                            blnLimparIndicesRepeticao = False
                        End If
                    End If
                End If
            End If
                
            'Verificar se o pai já tem contexto marcado
            strXPath = ".//"
'            If Not xmlNode.parentNode.selectSingleNode("@UltimaPosicao") Is Nothing Then
'                If CLng(Val(xmlNode.parentNode.selectSingleNode("@UltimaPosicao").Text)) > 0 Then
'                    strXPath = ".//*[@Posicao='" & _
'                               CLng(Val(xmlNode.parentNode.selectSingleNode("@UltimaPosicao").Text)) & _
'                               "']//"
'                End If
'            End If
            
            'Armazena o contexto a ser utilizado para a busca das informações
            Set xmlNodeContexto = pxmlMensagem
            
            'Armazenar o indice da repeticao para uso em caso de repet dentro de repet
            'Importante: caso não haja tag pai de repetição configurada o tradutor ignora o contexto e procura a partir da raiz
            ' caso haja tag pai configurada mas ela não seja encontrada o tradutor ignora o contexto e procura a aprtir da raiz
            ' e finalmente, se a tag pai estiver configurada e presente na mensagem de entrada o contexto é utilizado
            If Not xmlNode.selectSingleNode("@UltimaPosicao") Is Nothing Then
                If CLng(Val(xmlNode.selectSingleNode("@UltimaPosicao").Text)) = 0 Then
                    If Trim$(xmlNode.selectSingleNode("@TargetTag").Text) <> vbNullString Then
                        If Not pxmlMensagem.selectSingleNode(strXPath & xmlNode.selectSingleNode("@TargetTag").Text) Is Nothing Then
                            xmlNode.selectSingleNode("@UltimaPosicao").Text = pxmlMensagem.selectSingleNode(strXPath & xmlNode.selectSingleNode("@TargetTag").Text & "/@Posicao").Text
                            Set xmlNodeContexto = pxmlMensagem.selectSingleNode(strXPath & xmlNode.selectSingleNode("@TargetTag").Text)
                        Else
                            'caso não exista configuração de targettag para tags de repetição ignorar o fechamento do contexto
                            'assim a procura não fica restrita a uma tag pai
                            xmlNode.selectSingleNode("@UltimaPosicao").Text = "0"
                            'xmlNode.selectSingleNode("@UltimaPosicao").Text = "9999999"
                            
                            'Não deve-se limpar os indices pois a tag pai não existe e deve ser controlado pelo indice na busca
                            blnLimparIndicesRepeticao = False
                            
                        End If
                    Else
                        'caso não exista configuração de targettag para tags de repetição ignorar o fechamento do contexto
                        'assim a procura não fica restrita a uma tag pai
                        xmlNode.selectSingleNode("@UltimaPosicao").Text = "0"
                        
                        'Não deve-se limpar os indices pois a tag pai não existe e deve ser controlado pelo indice na busca
                        blnLimparIndicesRepeticao = False

                    End If
                Else
                    If Not pxmlMensagem.selectSingleNode(strXPath & xmlNode.selectSingleNode("@TargetTag").Text & "[@Posicao>'" & xmlNode.selectSingleNode("@UltimaPosicao").Text & "']") Is Nothing Then
                        xmlNode.selectSingleNode("@UltimaPosicao").Text = pxmlMensagem.selectSingleNode(strXPath & xmlNode.selectSingleNode("@TargetTag").Text & "[@Posicao>'" & xmlNode.selectSingleNode("@UltimaPosicao").Text & "']" & "/@Posicao").Text
                        Set xmlNodeContexto = pxmlMensagem.selectSingleNode(strXPath & xmlNode.selectSingleNode("@TargetTag").Text & "[@Posicao='" & xmlNode.selectSingleNode("@UltimaPosicao").Text & "']")
                    Else
                        'Ultrapassa o limite porque chegou ao fim das repetições
                        xmlNode.selectSingleNode("@UltimaPosicao").Text = "99999999"
                        Set xmlNodeContexto = Nothing
                    End If
                End If
            End If
            
            strTagPai = xmlNode.nodeName
            
            For lngX = 1 To lngRepeticoes  'To 1 Step -1
                
                'Armazena o numero da iteracao para uso no deslocamento string (repeticao string)
                xmlNode.selectSingleNode("@Iteracao").Text = lngX
                
                'Usar lnx-1 para pegar a primeira posicao.
                'flSubConverter pxmlMensagem
                flSubConverter xmlNodeContexto, _
                               xmlNode, _
                               pxmlNodeSaida.selectNodes("//" & xmlNode.nodeName).Item(pxmlNodeSaida.selectNodes("//" & xmlNode.nodeName).length - 1), _
                               pTipoParteSaida, _
                               IIf(blnUsarNumeroInteracaoPai, plngNumeroIteracao, lngX - 1)
                
            Next
            
            'If blnUsarNumeroInteracaoPai Then
            If blnLimparIndicesRepeticao Then
                'Limpar os indices armazenados para o contexto pois na proxima iteracao o indice do pai ira mudar
                flZerarIndicesDaRepeticao xmlNode
            End If
            
            
            If pTipoParteSaida = ParteXML Then
                'Remover o Item de Agruamento caso as tags estejam vazias.
                'Apenas do XML
                'String e ID são posicionais
                If pxmlNodeSaida.selectNodes("//" & strTagPai).Item(pxmlNodeSaida.selectNodes("//" & strTagPai).length - 1).Text = vbNullString Then
                    'Remover Tag
                    
                    pxmlNodeSaida.selectNodes("//" & strTagPai).Item(pxmlNodeSaida.selectNodes("//" & strTagPai).length - 1).parentNode.removeChild _
                            pxmlNodeSaida.selectNodes("//" & strTagPai).Item(pxmlNodeSaida.selectNodes("//" & strTagPai).length - 1)
                    
                End If
            End If
        End If
    Next

    Set AUXMLnode = Nothing
    
    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSubCOnverter Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

'Prepara o XML da mensagem de entrada para a tradução.
'Esta rotina inclui um índice em todas as tags do XML.
Private Sub flPrepararXML(ByRef pxmlMensagem As DOMDocument40)
On Error GoTo ErrorHandler
Dim lngX                                    As Long
Dim xmlElement                              As IXMLDOMElement

    'Verifica se a formatação de saida contém tags a partir do 1o. nivel.
    'Caso sim, indica que a mensagem pode ter repetições
    If Not pxmlMensagem.selectSingleNode("//Documento/Formato/*/*") Is Nothing Then
        
        lngX = 0
        'Adiciona indice nas tags
        For Each xmlElement In pxmlMensagem.selectNodes("//Documento/Mensagem//*")
            
            lngX = lngX + 1
            xmlElement.setAttribute "Posicao", lngX
                              
        Next
        
        For Each xmlElement In pxmlMensagem.selectNodes("//Documento/Formato/*//*")

            'Isso evita erro nas mensagem que já funcionam
            'Adiciona as tags de controle de repetição caso elas não existam
            If xmlElement.selectSingleNode("@RepetTag") Is Nothing Then
                xmlElement.setAttribute "RepetTag", 1
            End If
            If xmlElement.selectSingleNode("@UltimaPosicao") Is Nothing Then
                xmlElement.setAttribute "UltimaPosicao", 0
            End If
        
        Next
    
    End If

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPrepararXML Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

'Rotina auxiliar a SubConverter(Tradução)
'Limpa a marcação de indice utilizada nas traduções com repetições pois para cada contexto pai os indices precisam ser zerados.
Private Sub flZerarIndicesDaRepeticao(ByRef pxmlNodeRepet As IXMLDOMNode)

On Error Resume Next

Dim xmlNode                                 As IXMLDOMNode

    For Each xmlNode In pxmlNodeRepet.selectNodes(".//*")
        xmlNode.selectSingleNode("@UltimaPosicao").Text = "0"
    Next

End Sub

'Formata valores para tipos de saida String.
'Esta formatação é posicional.
Private Function flValorToSTR(psValor As String, _
                              psTipo As String, _
                              plTamanho As Long, _
                              plDecimais As Long, _
                              pbObrigatorio As Boolean, _
                              pstrNomeTag As String) As String

Dim lsNumero                                As String
Dim lsDecimal                               As String
                        
    Select Case UCase(psTipo)
      Case "STRING"
      
          flValorToSTR = Left(psValor & String(plTamanho, " "), plTamanho)
          
          If pbObrigatorio = "1" Then
            If Trim(flValorToSTR) = "" Then
                flValorToSTR = ""
            End If
          Else
            flValorToSTR = Left(psValor & String(plTamanho, " "), plTamanho)
          End If
            
      Case "NUMBER"
      
          If Not IsNumeric(plDecimais) Then
             plDecimais = CLng(0)
          End If
          
          
          If InStr(1, psValor, ",") > 0 Then
              
              lsNumero = Split(psValor, ",")(0)
              lsDecimal = Split(psValor, ",")(1)
               
              If Abs(plDecimais) > 0 Then
              
                  flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho - plDecimais) & _
                   Left(Trim(lsDecimal) & String(plDecimais, "0"), plDecimais)
               
              Else
              
                  flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho)
              
              End If
          Else
              
              flValorToSTR = Right(String(plTamanho, "0") & Trim(psValor), plTamanho - plDecimais) & String(plDecimais, "0")
           
          End If
          
          If pbObrigatorio = "1" Then
              If Trim(flValorToSTR) = "" Or Replace(Replace(flValorToSTR, "0", ""), ",", "") = "" Then
                  
                  If Left$(pstrNomeTag, 3) = "HO_" Then
                    flValorToSTR = "0000"
                  Else
                    flValorToSTR = ""
                  End If
              
              End If
          Else
              If InStr(1, psValor, ",") > 0 Then
                  
                  lsNumero = Split(psValor, ",")(0)
                  lsDecimal = Split(psValor, ",")(1)
                   
                  If Abs(plDecimais) > 0 Then
                  
                      flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho - plDecimais) & _
                       Left(Trim(lsDecimal) & String(plDecimais, "0"), plDecimais)
                   
                  Else
                  
                      flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho)
                  
                  End If
              Else
                  
                  flValorToSTR = Right(String(plTamanho, "0") & Trim(psValor), plTamanho - plDecimais) & String(plDecimais, "0")
               
              End If


          End If

    End Select
    
End Function

'Formata valores pata tipo de saida XML quando entrada  String.
'Motivo: Entradas string são posicionais. Dados numéricos não possuem separação de decimais e dados string vêm com espaços a direita.
Private Function flValorToXML(psValor As String, _
                              psTipo As String, _
                              plTamanho As Long, _
                              plDecimais As Long, _
                              pbObrigatorio As Boolean) As String

Dim lsNumero                                As String
Dim lsDecimal                               As String
Dim lsSinal                                 As String

    Select Case UCase(psTipo)
        Case "STRING"
        
            flValorToXML = Trim(psValor)
        
        Case "NUMBER"
            
            If Not IsNumeric(psValor) Then
                flValorToXML = psValor
                Exit Function
            End If
            
            If Not IsNumeric(plDecimais) Then
               plDecimais = CLng(0)
            End If
                                  
            If CDbl(psValor) < 0 Then
                lsSinal = "-"
            Else
                lsSinal = ""
            End If
            
            If pbObrigatorio Then
                If Replace(psValor, "0", "") = "" Then
                    flValorToXML = "0"
                Else
                  
                    If plDecimais > 0 Then
                        
                        If IsNumeric(Left(psValor, plTamanho - plDecimais)) Then
                            flValorToXML = FormatNumber(Left(psValor, plTamanho - plDecimais), 0, , , 0) & "," & Right(psValor, plDecimais)
                        Else
                            flValorToXML = vbNullString
                        End If
                        
                        If InStr(1, flValorToXML, "-") = 0 Then
                            flValorToXML = lsSinal & flValorToXML
                        End If
                        
                    Else
                        flValorToXML = (Left(psValor, plTamanho))
                        If plTamanho <= 15 Then
                            'Retira os zeros a esquerda
                            flValorToXML = CStr(Val(flValorToXML))
                            
                            If InStr(1, flValorToXML, "-") = 0 Then
                                flValorToXML = lsSinal & flValorToXML
                            End If

                        End If
                    End If
                    
                End If
            Else
                If plDecimais > 0 Then
                    
                    If IsNumeric(Left(psValor, plTamanho - plDecimais)) Then
                        flValorToXML = FormatNumber(Left(psValor, plTamanho - plDecimais), 0, , , 0) & "," & Right(psValor, plDecimais)
                    Else
                        flValorToXML = vbNullString
                    End If

                    
                    If InStr(1, flValorToXML, "-") = 0 Then
                        flValorToXML = lsSinal & flValorToXML
                    End If
                    
                Else
                    flValorToXML = (Left(psValor, plTamanho))
                    If plTamanho <= 15 Then
                        'Retira os zeros a esquerda
                        flValorToXML = CStr(Val(flValorToXML))
                        
                        If InStr(1, flValorToXML, "-") = 0 Then
                            flValorToXML = lsSinal & flValorToXML
                        End If
                        
                    End If
                End If
            End If
        
    End Select

End Function

'Retorna o valor de um determinado índice do CVS.
'Recebe como parametro o CSV e a posição a ser recuperada.
Private Function flValueFromCSV(psCSV As String, _
                                psSeparador As String, _
                                plIndex As Long, _
                                psTipo As String, _
                                pbObrigatorio As Boolean) As String
Dim lvCSV()                                 As String
Dim lsValor                                 As String
  
  lvCSV = Split(psCSV, psSeparador)

  If UBound(lvCSV) < (plIndex - 1) Then
      'Posição desejada não existe
      flValueFromCSV = ""
  Else

      lsValor = lvCSV(plIndex - 1)
      
      Select Case UCase(psTipo)
            Case "STRING"
                
                lsValor = Trim(lsValor)
                
            Case "NUMBER"
            
                If pbObrigatorio Then
                    If Replace(lsValor, "0", "") = "" Then
                        lsValor = ""
                    End If
                Else
                    If Replace(lsValor, "0", "") = "" Then
                        lsValor = "0"
                    End If
                End If
                
      End Select
      
      flValueFromCSV = lsValor

  End If

End Function

'Obtem o codigo da mensagem No. Controle IF para mensagens e tipo de mensagem para operações (ex. Previsto, Realizado, etc)
Private Function flObterCodigoMensagemNZ(ByVal pstrTipoMensagem As String, pstrHeaderNZ As String) As String

Dim udtProtocoloNZ                          As udtProtocoloNZ
Dim udtProtocoloNZAux                       As udtProtocoloNZAux

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim udtRemessaMovimento                     As udtRemessaMovimento
Dim udtRemessaMovimentoAux                  As udtRemessaMovimentoAux

Dim udtMaioresValores                       As udtMaioresValores
Dim udtMaioresValoresAux                    As udtMaioresValoresAux

Dim udtPZW0001                              As udtPZW0001
Dim udtPZW0001Aux                           As udtPZW0001Aux

Dim udtPZW0916                              As udtPZW0916
Dim udtPZW0916Aux                           As udtPZW0916Aux

Dim udtPZO00140                             As udtPZO00140
Dim udtPZO00140Aux                          As udtPZO00140Aux

Dim udtConsultaPZ                           As udtConsultaPZ
Dim udtConsultaPZAux                        As udtConsultaPZAux

Dim udtSTR0008R2                            As udtSTR0008R2
Dim udtSTR0008R2Aux                         As udtSTR0008R2Aux

Dim udtMovi_PJ_MoedaEstrangeira             As udtMovi_PJ_MoedaEstrangeira
Dim udtMovi_PJ_MoedaEstrangeiraAux          As udtMovi_PJ_MoedaEstrangeiraAux

On Error GoTo ErrorHandler

    Select Case CLng("0" & pstrTipoMensagem)
        
        Case enumTipoMensagemEntrada.MensagemNZA8, _
             enumTipoMensagemEntrada.MensagemA8NZ
        
            udtProtocoloNZAux.String = pstrHeaderNZ
            LSet udtProtocoloNZ = udtProtocoloNZAux
                        
            flObterCodigoMensagemNZ = udtProtocoloNZ.CodigoMensagem
        
        Case enumTipoMensagemEntrada.MensagemErroNZA8
    
            udtProtocoloErroNZAux.String = pstrHeaderNZ
            LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
                        
            flObterCodigoMensagemNZ = udtProtocoloErroNZ.CodigoMensagem
        
        Case enumTipoMensagemEntrada.MensagemA8PJPrevisto
            
            udtRemessaMovimentoAux.String = Mid(pstrHeaderNZ, 21)
            LSet udtRemessaMovimento = udtRemessaMovimentoAux
            
            If Trim(udtRemessaMovimento.TipoMovimento) = enumTipoMovimentoPJ.Previsto Then
                flObterCodigoMensagemNZ = "Previsto"
            ElseIf Trim(udtRemessaMovimento.TipoMovimento) = enumTipoMovimentoPJ.Realizado Then
                flObterCodigoMensagemNZ = "Realizado"
            ElseIf Trim(udtRemessaMovimento.TipoMovimento) = enumTipoMovimentoPJ.EstornoPrevisto Then
                flObterCodigoMensagemNZ = "Estorno Previsto"
            ElseIf Trim(udtRemessaMovimento.TipoMovimento) = enumTipoMovimentoPJ.EstornoRealizado Then
                flObterCodigoMensagemNZ = "Estorno Realizado"
            End If
            
            
        Case enumTipoMensagemEntrada.MensagemA8PJRealizado
        
            udtMaioresValoresAux.String = Mid(pstrHeaderNZ, 21)
            LSet udtMaioresValores = udtMaioresValoresAux
            
            If Trim(udtMaioresValores.TipoMovimento) = enumTipoMovimentoPJ.Previsto Then
                flObterCodigoMensagemNZ = "Previsto"
            ElseIf Trim(udtMaioresValores.TipoMovimento) = enumTipoMovimentoPJ.Realizado Then
                flObterCodigoMensagemNZ = "Realizado"
            ElseIf Trim(udtMaioresValores.TipoMovimento) = enumTipoMovimentoPJ.EstornoPrevisto Then
                flObterCodigoMensagemNZ = "Estorno Previsto"
            ElseIf Trim(udtMaioresValores.TipoMovimento) = enumTipoMovimentoPJ.EstornoRealizado Then
                flObterCodigoMensagemNZ = "Estorno Realizado"
            End If
            
        Case enumTipoMensagemEntrada.MensagemA8PJMoedaEstrangeira
        
            udtMovi_PJ_MoedaEstrangeiraAux.String = Mid(pstrHeaderNZ, 21)
            LSet udtMovi_PJ_MoedaEstrangeira = udtMovi_PJ_MoedaEstrangeiraAux
            
            If Trim(udtMovi_PJ_MoedaEstrangeira.TipoMovimento) = enumTipoMovimentoPJ.Previsto Then
                flObterCodigoMensagemNZ = "ME - Previsto"
            ElseIf Trim(udtMovi_PJ_MoedaEstrangeira.TipoMovimento) = enumTipoMovimentoPJ.Realizado Then
                flObterCodigoMensagemNZ = "ME - Realizado"
            ElseIf Trim(udtMovi_PJ_MoedaEstrangeira.TipoMovimento) = enumTipoMovimentoPJ.EstornoPrevisto Then
                flObterCodigoMensagemNZ = "ME - Estorno Previsto"
            ElseIf Trim(udtMovi_PJ_MoedaEstrangeira.TipoMovimento) = enumTipoMovimentoPJ.EstornoRealizado Then
                flObterCodigoMensagemNZ = "ME - Estorno Realizado"
            End If
            
        Case enumTipoMensagemEntrada.MensagemPZErro
            
            udtPZW0001Aux.String = Mid(pstrHeaderNZ, 21)
            LSet udtPZW0001 = udtPZW0001Aux
                        
            flObterCodigoMensagemNZ = udtPZW0001.ControleLegado
        
        Case enumTipoMensagemEntrada.MensagemPZR1
            
            udtPZO00140Aux.String = pstrHeaderNZ
            LSet udtPZO00140 = udtPZO00140Aux
            
            udtPZW0916Aux.String = udtPZO00140.HeaderNZ_PZ
            LSet udtPZW0916 = udtPZW0916Aux
            
            flObterCodigoMensagemNZ = udtPZW0916.NumeroControleLegado
        
        Case enumTipoMensagemEntrada.MensagemPZR2
            
            'RATS 400 - oCor 01
            If Mid(Mid(pstrHeaderNZ, 21), 38, 9) = "STR0008R2" Then
                udtConsultaPZAux.String = Mid(pstrHeaderNZ, 21)
                LSet udtConsultaPZ = udtConsultaPZAux
                flObterCodigoMensagemNZ = udtConsultaPZ.CO_MENSAGEM & " - " & udtConsultaPZ.RC_ROTINA
            Else
                udtSTR0008R2Aux.String = Mid(pstrHeaderNZ, 21)
                LSet udtSTR0008R2 = udtSTR0008R2Aux
                flObterCodigoMensagemNZ = udtSTR0008R2.NU_CTRL_EXTE_RECB
            End If
    End Select


    Exit Function
ErrorHandler:

    Err.Clear
    flObterCodigoMensagemNZ = ""

End Function


Private Function flFormataSaidaA7(ByRef pxmlMensagem As IXMLDOMNode, _
                                  ByVal pstrNomeCampoSaida As String, _
                                  ByVal vntValorEntrada As Variant, _
                                  ByVal plngTamanhoTag As Long) As Variant

Dim strTipoMensagem                         As String
Dim strXpathQuery                           As String
Dim xmlNodeRegra                            As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    flFormataSaidaA7 = vntValorEntrada
    
    If pxmlMensagem.selectSingleNode("//CO_MESG") Is Nothing Then Exit Function
    
    
    If IsNumeric(pxmlMensagem.selectSingleNode("//CO_MESG").Text) Then
        strTipoMensagem = Val(pxmlMensagem.selectSingleNode("//CO_MESG").Text)
    Else
        strTipoMensagem = Trim$(pxmlMensagem.selectSingleNode("//CO_MESG").Text)
    End If
    
    strXpathQuery = "//RegraFormatacaoSaidaA7/TIPO_MESG[@TP_MESG='" & strTipoMensagem & "']"
    
    If xmlRegraSaidaA7.selectSingleNode(strXpathQuery) Is Nothing Then Exit Function
        
    Set xmlNodeRegra = xmlRegraSaidaA7.selectSingleNode(strXpathQuery)

    If Not xmlNodeRegra.selectSingleNode(pstrNomeCampoSaida) Is Nothing Then
        
        strXpathQuery = pstrNomeCampoSaida & "/@FormataZerosEsquerda"
        
        If Not xmlNodeRegra.selectSingleNode(strXpathQuery) Is Nothing Then
            If xmlNodeRegra.selectSingleNode(strXpathQuery).Text = "S" Then
                flFormataSaidaA7 = fgCompletaString(vntValorEntrada, "0", CInt(plngTamanhoTag), True)
            End If
        End If
    End If
    
    
    Exit Function
ErrorHandler:
    
    Err.Clear
    flFormataSaidaA7 = vntValorEntrada
End Function

'Realiza a tradução dos tipos de dado aderentes ao Book 2.0 para o formato antigo
Public Function TraduzBook2(ByRef pxmlMensagem As DOMDocument40, _
                   Optional ByVal plngTipoTraducao As enumTipoTraducaoBookSPB = Book2_BookAntigo)

    '-------------------------------------------------------------------------------------
    'Descrição das alterações do BOOK SPB 2.0
    '-------------------------------------------------------------------------------------
    'Tipo                               BookAntigo               Book2
    '-------------------------------------------------------------------------------------
    'Numerico com decimais              '0,345'                  '0.345'
    'Data                               '20051231'               '2005-12-31'
    'AnoMes                             '200512'                 '2005-12'
    'Hora                               '185959'                 '18:59:59'
    'Data Hora                          '20051231185959'         '2005-12-31T18:59:59'
    '-------------------------------------------------------------------------------------

    Dim xmlConf                                     As MSXML2.DOMDocument40
    Dim objRegra                                    As MSXML2.IXMLDOMNode
    Dim objCampo                                    As MSXML2.IXMLDOMNode
    Dim strSQL                                      As String
    Dim strCodMsg                                   As String
    
    Dim strCampo                                    As String
    Dim strTipo                                     As String
    Dim dblDecimais                                 As Long
    
    'RATS 387 - OCOR 01 -------------------------------------------------------------------------
    If pxmlMensagem.selectSingleNode("//CodMsg") Is Nothing Then
        'Não é mensagem SPB
        Exit Function
    End If
    
    If UCase$(Right$(Trim$(pxmlMensagem.selectSingleNode("//CodMsg").Text), 1)) = "E" Then
        strCodMsg = Left$(Trim$(pxmlMensagem.selectSingleNode("//CodMsg").Text), 7)
    Else
        strCodMsg = Trim$(pxmlMensagem.selectSingleNode("//CodMsg").Text)
    End If
    'RATS 387 - OCOR 01 ------------------------------------------------------------------------
    
    If xmlConf Is Nothing Then
        'Cache das informaçoes de configuração, pode ficar em uma rotina separada, contanto
        'que esta rotina aqui acesse o 'xmlConf'.
        
        Set xmlConf = CreateObject("MSXML2.DOMDocument.4.0")
        
        strSQL = ""
        strSQL = strSQL & "SELECT       RTRIM(A.CO_MESG)       AS CO_MESG,          " & vbCrLf
        strSQL = strSQL & "             RTRIM(b.NO_TAG)        AS NO_TAG,           " & vbCrLf
        strSQL = strSQL & "             RTRIM(d.NO_TIPO_TAG)   AS NO_TIPO_TAG,      " & vbCrLf
        strSQL = strSQL & "             NVL(D.QT_CASA_DECI, 0) AS QT_CASA_DECI      " & vbCrLf
        strSQL = strSQL & "FROM                                                     " & vbCrLf
        strSQL = strSQL & "             A8.TB_MENSAGEM A,                           " & vbCrLf
        strSQL = strSQL & "             A8.TB_TAG B,                                " & vbCrLf
        strSQL = strSQL & "             A8.TB_TAG_MENSAGEM C ,                      " & vbCrLf
        strSQL = strSQL & "             A8.TB_TIPO_TAG D                            " & vbCrLf
        strSQL = strSQL & "WHERE                                                    " & vbCrLf
        strSQL = strSQL & "             (A.SQ_MESG       = C.SQ_MESG)               " & vbCrLf
        strSQL = strSQL & "AND          (C.SQ_TAG        = B.SQ_TAG)                " & vbCrLf
        strSQL = strSQL & "AND          (B.SQ_TIPO_TAG   = D.SQ_TIPO_TAG)           " & vbCrLf
        strSQL = strSQL & "AND          A.CO_MESG        = '" & strCodMsg & "'      " & vbCrLf
        strSQL = strSQL & "AND          (                                           " & vbCrLf
        strSQL = strSQL & "                (                                        " & vbCrLf
        strSQL = strSQL & "                    D.NO_TIPO_TAG  LIKE '%Data%'         " & vbCrLf
        strSQL = strSQL & "                 OR D.NO_TIPO_TAG  LIKE '%Hora%'         " & vbCrLf
        strSQL = strSQL & "                 OR D.NO_TIPO_TAG  LIKE '%Ano%'          " & vbCrLf
        strSQL = strSQL & "                 OR D.NO_TIPO_TAG  LIKE '%Mes%'          " & vbCrLf
        strSQL = strSQL & "                )                                        " & vbCrLf
        strSQL = strSQL & "             OR                                          " & vbCrLf
        strSQL = strSQL & "                D.QT_CASA_DECI >0                        " & vbCrLf
        strSQL = strSQL & "             )                                           " & vbCrLf
        strSQL = strSQL & "ORDER BY     1,2,3                                       " & vbCrLf
   
        xmlConf.loadXML fgQueryXMLLerTodos("CONF", strSQL, "CONF")
   
    End If
    
    For Each objRegra In xmlConf.selectNodes("//Grupo_CONF[./CO_MESG='" & strCodMsg & "']")
        'Para cada regra desta mensagem
        strCampo = objRegra.selectSingleNode("NO_TAG").Text
        strTipo = objRegra.selectSingleNode("NO_TIPO_TAG").Text
        dblDecimais = Val(objRegra.selectSingleNode("QT_CASA_DECI").Text)
        
        
        For Each objCampo In pxmlMensagem.selectNodes("//" & strCampo)
        
            If plngTipoTraducao = Book2_BookAntigo Then
                '//////////////////////////////////////////////////////////////////
                '//
                '//  Tradução do formato novo para o formato antigo
                '//
                '//////////////////////////////////////////////////////////////////
                If objCampo.Text = "" Then
                    'Se campo veio vazio, remove-o do XML
                    If objCampo.attributes.length = 0 Then
                        objCampo.parentNode.removeChild objCampo
                    End If
                Else
                    If dblDecimais > 0 Then
                        'Campo numérico com casas decimais
                        objCampo.Text = Replace(objCampo.Text, ".", ",")
                    Else
                        'Campo do tipo data ou hora
                        objCampo.Text = Replace( _
                                        Replace( _
                                        Replace( _
                                                    objCampo.Text, _
                                        "T", ""), _
                                        "-", ""), _
                                        ":", "")
                    End If
                End If
            ElseIf plngTipoTraducao = BookAntigo_Book2 Then
                '//////////////////////////////////////////////////////////////////
                '//
                '//  Tradução do formato antigo para o formato novo
                '//
                '//////////////////////////////////////////////////////////////////
                If objCampo.Text = "" Then
                    'Se campo veio vazio, remove-o do XML
                    If objCampo.attributes.length = 0 Then
                        objCampo.parentNode.removeChild objCampo
                    End If
                Else
                    If dblDecimais > 0 Then
                        'Campo numérico com casas decimais
                        objCampo.Text = Replace(objCampo.Text, ",", ".")
                    Else
                        Select Case strTipo
                            Case "AnoMes"
                                If Len(objCampo.Text) = 6 Then
                                    objCampo.Text = Format(objCampo.Text, "0000-00")
                                End If
                            Case "Data"
                                If Len(objCampo.Text) = 8 Then
                                    objCampo.Text = Format(objCampo.Text, "0000-00-00")
                                End If
                            Case "Data Hora"
                                If Len(objCampo.Text) = 14 Then
                                    objCampo.Text = Format(objCampo.Text, "0000-00-00T00:00:00")
                                End If
                            Case "Hora"
                                If Len(objCampo.Text) = 6 Then
                                    objCampo.Text = Format(objCampo.Text, "00:00:00")
                                End If
                        End Select
                    End If
                End If
            End If
        Next
    Next

    '#################################################
    'Remover todas as TAGS sem conteudo
    Dim objxmlNodeDel                              As MSXML2.IXMLDOMNode
    Dim objxmlNodeListDel                          As MSXML2.IXMLDOMNodeList
    
    Set objxmlNodeListDel = pxmlMensagem.selectNodes("//*[.='']")
    
    For Each objxmlNodeDel In objxmlNodeListDel
        If objxmlNodeDel.attributes.length = 0 Then
            objxmlNodeDel.parentNode.removeChild objxmlNodeDel
        End If
    Next
    '#################################################

End Function
