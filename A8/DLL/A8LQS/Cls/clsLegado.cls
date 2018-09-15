VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLegado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF8D2530249"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsável por montar as mensagens de envio para os sistemas legados.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

'Enviar mensagem de rejeição para o Legado
'   - Montar protocolo de integração A7
'   - Montar Remessa com os erros identificados

Public Function MontarMensagemRejeicao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                       ByRef xmlErros As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim intCount                                As Integer
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim lngTipoMensagem                         As Long
Dim strTipoMensagemRet                      As String

On Error GoTo ErrorHandler

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = UCase$(fgCompletaString(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, " ", 3, False))
    udtProtocolo.SiglaSistemaOrigem = UCase$(fgCompletaString("A8", " ", 3, False))

    If Not xmlRemessa.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        If xmlRemessa.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
            Call objTipoOperacao.ObterTipoOperacaoPorMsgBUS(xmlRemessa, _
                                                            lngTipoMensagem, _
                                                            "", _
                                                            strTipoMensagemRet, _
                                                            enumAcaoControleInterno.ErrorHandlerBypass)
            lngTipoMensagem = CLng("0" & strTipoMensagemRet)
        End If
    End If
            
    If lngTipoMensagem = 0 Then
        If xmlRemessa.selectSingleNode("//TP_MESG_RETN_INTE") Is Nothing Then
            lngTipoMensagem = objTipoOperacao.ObterTipoOperacaoRetorno(CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text))
        Else
            lngTipoMensagem = xmlRemessa.selectSingleNode("//TP_MESG_RETN_INTE").Text
        End If
    End If
    
    If Trim$(udtProtocolo.SiglaSistemaDestino) = "WZ" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "WZC" Then
        lngTipoMensagem = lngTipoMensagem + 2000
    Else
        lngTipoMensagem = lngTipoMensagem
    End If
    
    udtProtocolo.TipoMensagem = fgCompletaString(lngTipoMensagem, "0", 9, True)
    
    LSet udtProtocoloAux = udtProtocolo

    fgAppendNode xmlRemessa, "MESG", "TP_RETN", "2"
    fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", "0"

    'Adicionar o erro a mensagem
    intCount = 1
    For Each xmlNode In xmlErros.selectNodes("Erro/Grupo_ErrorInfo")
        fgAppendNode xmlRemessa, "MESG", "CO_ERRO" & intCount, xmlNode.selectSingleNode("Number").Text
        fgAppendNode xmlRemessa, "MESG", "DE_ERRO" & intCount, xmlNode.selectSingleNode("Description").Text
        intCount = intCount + 1
        If intCount > 3 Then Exit For
    Next

    MontarMensagemRejeicao = udtProtocoloAux.String & xmlRemessa.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemRejeicao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar mensagem de retorno para o Legado
'   - Montar protocolo de integração A7
'   - Montar Remessa
'   - Incluir a mensagem de retorno na tabela de Mensagem Interna

Public Function EnviarRetorno(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                     Optional ByRef xmlMensagem As MSXML2.DOMDocument40, _
                     Optional ByVal pstrHeaderNZ As String) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim lngTipoMensagem                         As Long
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim strTipoMensagem                         As String
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40

Dim intFormatoSaidaMensagem                 As Integer
Dim strTagsErroNegocioBacen                 As String
Dim domNode                                 As MSXML2.IXMLDOMNode
Dim intCountErroNegocioBacen                As Integer
Dim objLogErro                              As A6A7A8CA.clsLogErro
Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

    On Error GoTo ErrorHandler
    
    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    
    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")

    If Not xmlOperacao.selectSingleNode("//CO_ERRO1") Is Nothing Then
        If xmlOperacao.selectSingleNode("//CO_ERRO1").Text = "4007" Then 'Veículo Legal inválido"
            xmlOperacao.selectSingleNode("//CO_ERRO1").Text = vbNullString
    
            If Not xmlOperacao.selectSingleNode("//DE_ERRO1") Is Nothing Then
                xmlOperacao.selectSingleNode("//DE_ERRO1").Text = vbNullString
            End If
        End If
    End If
    
    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = UCase$(fgCompletaString(xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, " ", 3, False))
    udtProtocolo.SiglaSistemaOrigem = UCase$(fgCompletaString("A8", " ", 3, False))
    
    If xmlOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE") Is Nothing Then
        Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
        Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
        If Not xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
            xmlTipoOperacao.loadXML objTipoOperacao.Ler(CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text))
        ElseIf Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
            xmlTipoOperacao.loadXML objTipoOperacao.LerPorTipoMensagemRecebimento(CLng(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text))
        End If
        
        If Val(xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text) = 0 Then
            Exit Function
        End If
        
        lngTipoMensagem = CLng(xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text)
        
        Set objTipoOperacao = Nothing
        Set xmlTipoOperacao = Nothing
        
    Else
        lngTipoMensagem = CLng(xmlOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text)
    End If
    
    If Trim$(udtProtocolo.SiglaSistemaDestino) = "WZ" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "WZC" Then
        lngTipoMensagem = lngTipoMensagem + 2000
    Else
        lngTipoMensagem = lngTipoMensagem
    End If

    strTipoMensagem = xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text
    
    udtProtocolo.TipoMensagem = fgCompletaString(lngTipoMensagem, "0", 9, True)
    xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text = fgCompletaString(lngTipoMensagem, "0", 9, True)

    LSet udtProtocoloAux = udtProtocolo

    'inclui tag TP_RETN
    If Trim$(udtProtocolo.SiglaSistemaDestino) = "GPC" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "R2" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "BOL" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "HQ" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "CMX" Then  '//INCLUSÂO DO BOL e CMX (NICK)
        fgAppendNode xmlOperacao, "MESG", "TP_RETN", "2" 'exceção para sistemas legados GPC e R2
    Else
        fgAppendNode xmlOperacao, "MESG", "TP_RETN", "1" 'default
    End If

    'inclui tags (TX_XML_ERRO, CO_ERRO1, DE_ERRO1...) para reportar para o legado os erros gerados na rejeicao de mensagens no Bacen/NZ/PK
    If Trim$(udtProtocolo.SiglaSistemaDestino) = "GPC" _
    Or Trim$(udtProtocolo.SiglaSistemaDestino) = "R2" _
    Or Trim$(udtProtocolo.SiglaSistemaDestino) = "BOL" _
    Or Trim$(udtProtocolo.SiglaSistemaDestino) = "HQ" _
    Or Trim$(udtProtocolo.SiglaSistemaDestino) = "CMX" _
    Or Trim$(udtProtocolo.SiglaSistemaDestino) = "E2" Or Trim$(udtProtocolo.SiglaSistemaDestino) = "E2A" Then 'apenas para os sistemas legados de COMEX
        
        If Right(Trim(udtProtocoloErroNZ.CodigoMensagem), 1) = "E" Then 'mensagens de erro vindas do Bacen/Câmara
            For Each domNode In xmlMensagem.documentElement.selectNodes("//*/*")
                If Not domNode.attributes.getNamedItem("CodErro") Is Nothing Then
                    
                    'verifica se é o node raiz, ex.: CAM0021, CAM0022, etc. Caso seja raiz, envia em branco o espaço reservado para o ConteudoTag ( formatação da tag TX_XML_ERRO = Tag|CodigoErro|ConteudoTag; )
                    If domNode.parentNode.nodeName = "SISMSG" Then
                        strTagsErroNegocioBacen = strTagsErroNegocioBacen & domNode.nodeName & "|" & domNode.attributes.Item(0).Text & "| ;"
                    Else
                        strTagsErroNegocioBacen = strTagsErroNegocioBacen & domNode.nodeName & "|" & domNode.attributes.Item(0).Text & "|" & domNode.Text & ";"
                    End If
                    
                    intCountErroNegocioBacen = intCountErroNegocioBacen + 1
                    If intCountErroNegocioBacen <= 3 Then
                        If IsNumeric(Right(domNode.attributes.Item(0).Text, 4)) = True Then
                            fgAppendNode xmlOperacao, "MESG", "CO_ERRO" & intCountErroNegocioBacen, Val(Right(domNode.attributes.Item(0).Text, 4))
                            fgAppendNode xmlOperacao, "MESG", "DE_ERRO" & intCountErroNegocioBacen, "Campo " & domNode.nodeName & " - " & objLogErro.ObterDescErroNegocio(domNode.attributes.Item(0).Text) & " - Erro " & domNode.attributes.Item(0).Text
                        Else
                            fgAppendNode xmlOperacao, "MESG", "CO_ERRO" & intCountErroNegocioBacen, 0
                            fgAppendNode xmlOperacao, "MESG", "DE_ERRO" & intCountErroNegocioBacen, "Campo " & domNode.nodeName & " - Erro genérico" & " - " & domNode.attributes.Item(0).Text
                        End If
                    End If
                End If
            Next
            fgAppendNode xmlOperacao, "MESG", "TX_XML_ERRO", Left(strTagsErroNegocioBacen, 200) 'a definição do campo é de 200 no layout
        
        ElseIf Len(Trim(udtProtocoloErroNZ.CodigoMensagem)) = 7 Then 'mensagens de erro vindas do NZ/PK
            If Trim(udtProtocoloErroNZ.CodigoErro1) <> "" Then
                fgAppendNode xmlOperacao, "MESG", "CO_ERRO1", Left(Trim(udtProtocoloErroNZ.CodigoErro1), 5)
                fgAppendNode xmlOperacao, "MESG", "DE_ERRO1", "Campo " & Trim(udtProtocoloErroNZ.NomeDoCampo1) & objLogErro.ObterDescricaoErroNZPKPJ(Trim(udtProtocoloErroNZ.CodigoErro1)) & " - Erro " & Trim(udtProtocoloErroNZ.CodigoErro1)
            End If
            If Trim(udtProtocoloErroNZ.CodigoErro2) <> "" Then
                fgAppendNode xmlOperacao, "MESG", "CO_ERRO2", Left(Trim(udtProtocoloErroNZ.CodigoErro2), 5)
                fgAppendNode xmlOperacao, "MESG", "DE_ERRO2", "Campo " & Trim(udtProtocoloErroNZ.NomeDoCampo2) & objLogErro.ObterDescricaoErroNZPKPJ(Trim(udtProtocoloErroNZ.CodigoErro2)) & " - Erro " & Trim(udtProtocoloErroNZ.CodigoErro2)
            End If
            If Trim(udtProtocoloErroNZ.CodigoErro3) <> "" Then
                fgAppendNode xmlOperacao, "MESG", "CO_ERRO3", Left(Trim(udtProtocoloErroNZ.CodigoErro3), 5)
                fgAppendNode xmlOperacao, "MESG", "DE_ERRO3", "Campo " & Trim(udtProtocoloErroNZ.NomeDoCampo3) & objLogErro.ObterDescricaoErroNZPKPJ(Trim(udtProtocoloErroNZ.CodigoErro3)) & " - Erro " & Trim(udtProtocoloErroNZ.CodigoErro3)
            End If
            strTagsErroNegocioBacen = strTagsErroNegocioBacen & IIf(Trim(udtProtocoloErroNZ.CodigoErro1) <> "", Trim(udtProtocoloErroNZ.NomeDoCampo1) & "|" & Trim(udtProtocoloErroNZ.CodigoErro1) & "|" & Trim(udtProtocoloErroNZ.ConteúdoCampoErro1) & ";", "")
            strTagsErroNegocioBacen = strTagsErroNegocioBacen & IIf(Trim(udtProtocoloErroNZ.CodigoErro2) <> "", Trim(udtProtocoloErroNZ.NomeDoCampo2) & "|" & Trim(udtProtocoloErroNZ.CodigoErro2) & "|" & Trim(udtProtocoloErroNZ.ConteúdoCampoErro2) & ";", "")
            strTagsErroNegocioBacen = strTagsErroNegocioBacen & IIf(Trim(udtProtocoloErroNZ.CodigoErro3) <> "", Trim(udtProtocoloErroNZ.NomeDoCampo3) & "|" & Trim(udtProtocoloErroNZ.CodigoErro3) & "|" & Trim(udtProtocoloErroNZ.ConteúdoCampoErro3) & ";", "")
            strTagsErroNegocioBacen = strTagsErroNegocioBacen & IIf(Trim(udtProtocoloErroNZ.CodigoErro4) <> "", Trim(udtProtocoloErroNZ.NomeDoCampo4) & "|" & Trim(udtProtocoloErroNZ.CodigoErro4) & "|" & Trim(udtProtocoloErroNZ.ConteúdoCampoErro4) & ";", "")
            strTagsErroNegocioBacen = strTagsErroNegocioBacen & IIf(Trim(udtProtocoloErroNZ.CodigoErro5) <> "", Trim(udtProtocoloErroNZ.NomeDoCampo5) & "|" & Trim(udtProtocoloErroNZ.CodigoErro5) & "|" & Trim(udtProtocoloErroNZ.ConteúdoCampoErro5) & ";", "")
            fgAppendNode xmlOperacao, "MESG", "TX_XML_ERRO", Left(strTagsErroNegocioBacen, 200) 'a definição do campo é de 200 no layout
        End If
        
    End If
    Set objLogErro = Nothing

    'Estas TAGS foram criadas para retorno do Cetip e BMA, pois o tamanho é diferente em cada uma
    If Not xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_COMD_OPER_SELIC", xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text
        fgAppendNode xmlOperacao, "MESG", "NU_COMD_OPER_BMA", xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text
    End If

    'Esta TAG foi incluída para atender necessidade dos layouts de retorno 250 (Interbancário) e 254 (Arbitragem)
    If xmlOperacao.documentElement.selectSingleNode("CO_MESG_SPB") Is Nothing Then
        If Trim(udtProtocoloErroNZ.CodigoMensagem) <> "" Then
            fgAppendNode xmlOperacao, "MESG", "CO_MESG_SPB", Trim(udtProtocoloErroNZ.CodigoMensagem)
        Else
            fgAppendNode xmlOperacao, "MESG", "CO_MESG_SPB", "NAO_IDENT"
        End If
    Else
        If Trim(udtProtocoloErroNZ.CodigoMensagem) <> "" Then
            xmlOperacao.documentElement.selectSingleNode("CO_MESG_SPB").Text = Trim(udtProtocoloErroNZ.CodigoMensagem)
        End If
    End If

    'Inclui tag DtMovto quando ela não existe nem no xmlMensagem nem no xmlOperacao
    If xmlOperacao.documentElement.selectSingleNode("//DtMovto") Is Nothing And xmlMensagem Is Nothing Then
        If Not xmlOperacao.documentElement.selectSingleNode("//DT_MOVI") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "DtMovto", xmlOperacao.documentElement.selectSingleNode("//DT_MOVI").Text
        End If
    End If

    If Not xmlMensagem Is Nothing Then
        fgAppendXML xmlOperacao, "MESG", xmlMensagem.xml
    End If

    intFormatoSaidaMensagem = ObterFormatoMensagemRetorno(fgVlr_To_Xml(lngTipoMensagem), Trim$(udtProtocolo.SiglaSistemaDestino))
    
    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text, _
                                       enumTipoSolicitacao.RetornoLegado, _
                                       xmlOperacao, _
                                       intFormatoSaidaMensagem
    Set objOperacaoMensagemInterna = Nothing

    EnviarRetorno = udtProtocoloAux.String & xmlOperacao.xml

    xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    Exit Function
ErrorHandler:

    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set objLogErro = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRetorno Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter Formato de Saída da Mensagem de Retorno, de acordo com o sistema destino
Public Function ObterFormatoMensagemRetorno(ByVal pstrTipoMensagem As String, _
                                            ByVal pstrSiglaSistemaDestino As String) As Integer

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

    On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_FORM_MESG_SAID " & vbNewLine & _
             " FROM     A7.TB_REGR_SIST_DEST " & vbNewLine & _
             " WHERE    TP_MESG      = '" & pstrTipoMensagem & "'" & vbNewLine & _
             " AND      SG_SIST_DEST = '" & pstrSiglaSistemaDestino & "'" & vbNewLine & _
             " ORDER BY DH_INIC_VIGE_REGR_TRAP DESC"
             
    Set objRS = CreateObject("ADODB.Recordset")
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        objRS.MoveFirst
        ObterFormatoMensagemRetorno = objRS!TP_FORM_MESG_SAID
    Else
        ObterFormatoMensagemRetorno = 0
    End If
             
    Set objRS = Nothing
    
    Exit Function

ErrorHandler:
    Set objRS = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterFormatoMensagemRetorno Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

                     

