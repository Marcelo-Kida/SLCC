VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF75481038B"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelo envio de mensagens do sistema A7.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private objMQAX200                          As A6A7A8CA.clsMQAX200
Private ObjectContext                       As COMSVCSLib.ObjectContext
Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

'Método responsável por retirar a mensagem da fila MQSeries;
'Executar o método clsMQAX200.MQGet;
Public Function ReceberMensagemMQ(ByVal pstrNomeFila As String, _
                                  ByRef pstrRetornoLogErro As String, _
                                  ByRef pstrMessageIdHex As String, _
                                  ByRef plngBackOutCount As Long, _
                                  ByRef pstrCorrelationID As String) As Boolean
    
Dim lngLenMensagemRecebida                  As Long
Dim strMensagemRecebida                     As String
Dim strMessageDescriptor                    As String

On Error GoTo ErrorHandler
    
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    objMQAX200.MQInit
    
    objMQAX200.MQOpen pstrNomeFila, MQOO_INPUT_SHARED
     
    If objMQAX200.MQGet(lngLenMensagemRecebida, _
                        strMensagemRecebida, _
                        True, _
                        True, _
                        pstrMessageIdHex, _
                        pstrCorrelationID, _
                        strMessageDescriptor, _
                        plngBackOutCount) Then
        
        objMQAX200.MQClose
        
        Call flProcessaEnvioMensagem(strMensagemRecebida, pstrNomeFila)
        
        ReceberMensagemMQ = True
    Else
        ReceberMensagemMQ = False
    End If
        
    objMQAX200.MQEnd
    
    Set objMQAX200 = Nothing
       
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
Exit Function
ErrorHandler:

    'Caso ocorra algum erro não tratado, desfazer a transação e devolver o erro
    'para fazer o put na fila de erro para que o processo não pare

    pstrRetornoLogErro = Err.Description
    
    Set objMQAX200 = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If

End Function

'Método traduz e posta a mensagem recebida.
Public Function flProcessaEnvioMensagem(ByRef pstrXMLMensagem As String, _
                               Optional ByVal pstrNomeFila As String)

Dim objDomMensagem                          As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set objDomMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    objDomMensagem.loadXML pstrXMLMensagem
        
    'Traduzir Mensagem
    If flTraduzirMensagem(objDomMensagem, pstrNomeFila) Then
        'Postar Mensagem Traduzida
        Call flPostarMensagemTraduzida(objDomMensagem)
    End If
    
    Set objDomMensagem = Nothing

Exit Function
ErrorHandler:
    
    Set objDomMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaEnvioMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Traduzir a mensagem recebida no sistema A7.
Private Function flTraduzirMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                           Optional ByVal pstrNomeFila As String) As Boolean

Dim objTradutor                             As A7Server.clsTradutor
Dim blnTraducaoBemSucedida                  As Boolean

    On Error GoTo ErrorHandler
        
    Set objTradutor = CreateObject("A7Server.clsTradutor")
    blnTraducaoBemSucedida = objTradutor.TraduzirMensagem(pxmlMensagem, pstrNomeFila)
    Set objTradutor = Nothing
            
    If blnTraducaoBemSucedida Then
        'Se tradução OK
        Call flSalvarMensagemDadoSaida(pxmlMensagem)
        flTraduzirMensagem = True
    Else
        'Se tradução Não OK - Salva Mensagem Traduzida
        Call flSalvarMensagemDadoSaida(pxmlMensagem)
        Call flSalvarSituacaoMensagem(pxmlMensagem, enumOcorrencia.CanceladaErroTradução)
        
        'Pikachu - Damanda 1026
        'Enviar Resposta para sistema Legado - Layout de Resposta
        Call EnviarMensagemRejeicaoLegado(pxmlMensagem, enumOcorrencia.CanceladaErroTradução)
        
        flTraduzirMensagem = False
    End If
    
    Exit Function
ErrorHandler:

    Set objTradutor = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTraduzirMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Incluir nova situação para a mensagem transitada.
Private Function flSalvarSituacaoMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                          ByVal plngCodigoOcorrencia As Long) As Boolean
                                                 
Dim objSituacaoMensagem                   As A7Server.clsSituacaoMensagem
    
On Error GoTo ErrorHandler
    
    Set objSituacaoMensagem = CreateObject("A7Server.clsSituacaoMensagem")
    Call objSituacaoMensagem.SalvarSituacaoMensagem(pxmlMensagem, plngCodigoOcorrencia)
    Set objSituacaoMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objSituacaoMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarSituacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar a mensagem de saída (após a sua tradução).
Private Function flSalvarMensagemDadoSaida(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objMensagem                       As A7Server.clsMensagem

On Error GoTo ErrorHandler
    
    Set objMensagem = CreateObject("A7Server.clsMensagem")
    Call objMensagem.SalvarMensagemDadosSaida(pxmlMensagem)
    Set objMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarMensagemDadoSaida Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Preparar a memsagem recebida para o envio ao sistema destino.
Private Function flPostarMensagemTraduzida(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objSistema                              As MSXML2.DOMDocument40
Dim strSistema                              As String
Dim strAlerta                               As String
Dim strNomeFila                             As String

'Variaveis utilizadas pelo Carimbo NZ
Dim objCarimbador                           As A7Server.clsCarimbador
Dim xmlMensagemSPB                          As MSXML2.DOMDocument40
Dim strNumCtrlIF                            As String
Dim strValor                                As String
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
        
    strSistema = flObterSistema(pxmlMensagem)
    
    If Trim(strSistema) = vbNullString Then
        strAlerta = "Nome da Fila não cadastrada para o sistema: " & pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text
        Call fgGeraInformacaoAlerta(strAlerta, Format(Now, "dd/mm/yyyy HH:mm:ss"), "Gerenciador de envio", strAlerta)
        strNomeFila = ""
    Else
        
        If Trim(pxmlMensagem.documentElement.selectSingleNode("NO_FILA_MQSE_DEST").Text) = vbNullString Then
            Set objSistema = CreateObject("MSXML2.DOMDocument.4.0")
            objSistema.loadXML strSistema
            strNomeFila = Trim(objSistema.documentElement.selectSingleNode("Grupo_EnderecoFilaMqseries/NO_FILA_MQSE").Text)
            Set objSistema = Nothing
        Else
            strNomeFila = Trim(pxmlMensagem.documentElement.selectSingleNode("NO_FILA_MQSE_DEST").Text)
        End If
    
        If pxmlMensagem.selectSingleNode("//SG_SIST_DEST").Text = "NZ" Then
            'Enviar carimbo
            On Error Resume Next
            
            'Obtem a mensagem de entrada pois esta contem as tags necessarias com nomes já padronizados pelo A8
            Set xmlMensagemSPB = CreateObject("MSXML2.DOMDocument.4.0")
            xmlMensagemSPB.loadXML Mid$(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 21)
           
            'Obtem o Numero de Controle IF
            If Not xmlMensagemSPB.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
                strNumCtrlIF = xmlMensagemSPB.selectSingleNode("//NU_CTRL_IF").Text
            Else
                strNumCtrlIF = vbNullString
            End If
            
            'Obtem o Código da Mensagem
            If Not xmlMensagemSPB.selectSingleNode("//CO_MESG") Is Nothing Then
                strCodigoMensagem = xmlMensagemSPB.selectSingleNode("//CO_MESG").Text
            ElseIf Not IsNumeric(Left$(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 9)) Then
                strCodigoMensagem = Trim$(Left$(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 9))
            Else
                strCodigoMensagem = vbNullString
            End If
            
            'Obtem o Valor da Mensagem
            If Not xmlMensagemSPB.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                strValor = xmlMensagemSPB.selectSingleNode("//VA_OPER_ATIV").Text
            Else
                strValor = vbNullString
            End If
            
            Set objCarimbador = CreateObject("A7Server.clsCarimbador")
            objCarimbador.IncluirAvisoMonitor objMQAX200, _
                                              CLng(pxmlMensagem.selectSingleNode("//CO_EMPR_ORIG").Text), _
                                              strNumCtrlIF, _
                                              strCodigoMensagem, _
                                              FLUXO_MONITOR_NORMAL, _
                                              MONITOR_ENVIADA_NZ, _
                                              "A8", "NZ", _
                                              strNomeFila, _
                                              fgDataHoraServidor(enumFormatoDataHoraAux.DataAux), _
                                              strValor
            Set objCarimbador = Nothing
            Set xmlMensagemSPB = Nothing
        
            Err.Clear
            On Error GoTo 0
        
        End If
    
        If strNomeFila = vbNullString Then
            strAlerta = "Nome da Fila não cadastrada para o sistema: " & pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text
            Call fgGeraInformacaoAlerta(strAlerta, Format(Now, "dd/mm/yyyy HH:mm:ss"), "Gerenciador de envio", strAlerta)
            strNomeFila = ""
        End If
    
    End If
   
    Call flPostarMQSeries(pxmlMensagem, strNomeFila)
    
    Exit Function
ErrorHandler:
    
    Set objSistema = Nothing
    Set objCarimbador = Nothing
    Set xmlMensagemSPB = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPostarMensagemTraduzida Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os dados referentes ao sistema destino da mensagem recebida.
Private Function flObterSistema(ByRef pxmlMensagem As MSXML2.DOMDocument40) As String

Dim objSistema                              As A7Server.clsParametrosSistema

On Error GoTo ErrorHandler

    Set objSistema = CreateObject("A7Server.clsParametrosSistema")
    
    flObterSistema = objSistema.Ler(pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text, _
                               CLng(pxmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text))

    Set objSistema = Nothing

    Exit Function
ErrorHandler:

    Set objSistema = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterSistema Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar a mensagem traduzida para a fila do seu sistema destino.
Private Function flPostarMQSeries(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                  ByVal pstrNomeFila As String) As Boolean


Dim strMessageDescriptor                    As String
Dim strMensagemTraduzida                    As String
Dim strCorrelationID                        As String
Dim strMessageId                            As String

Dim xmlMensagemPostada                      As DOMDocument40

Dim lngPrioridade                           As Long

On Error GoTo ErrorHandler

    Set xmlMensagemPostada = New DOMDocument40
    
    xmlMensagemPostada.loadXML pxmlMensagem.selectSingleNode("//TX_CNTD_SAID").Text
    
    'Pika
    'É aqui que vc vai decidir se é boleto ou Operação
    
    lngPrioridade = 9
    
    If Not xmlMensagemPostada.selectSingleNode("//TP_SOLI") Is Nothing Then
        
        If pxmlMensagem.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemEntrada.MensagemIDADV And _
           pxmlMensagem.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemEntrada.MensagemIDABG Then
            
            Select Case CLng(xmlMensagemPostada.selectSingleNode("//TP_SOLI").Text)
                Case 1
                    pstrNomeFila = "A8Q.E.ENTRADA_BOLETO"
                Case 2
                    pstrNomeFila = "A8Q.E.ENTRADA_OPERACAO"
                Case Else
                    pstrNomeFila = "A8Q.E.ENTRADA"
            End Select
            
            
        End If
        
    Else
        If Not xmlMensagemPostada.selectSingleNode("//TP_MESG") Is Nothing Then
    
            Select Case CLng(xmlMensagemPostada.selectSingleNode("//TP_MESG").Text)
                Case enumTipoMensagemEntrada.MensagemNZA8, _
                     enumTipoMensagemEntrada.MensagemErroNZA8, _
                     enumTipoMensagemEntrada.MensagemPZErro, _
                     enumTipoMensagemEntrada.MensagemPZR1, _
                     enumTipoMensagemEntrada.MensagemPZR2, _
                     enumTipoMensagemEntrada.MensagemSTR0010R2PZA8
                
                    pstrNomeFila = "A8Q.E.ENTRADA_BACEN"
                
                Case enumTipoMensagemEntrada.MensagemRetornoDV, _
                     enumTipoMensagemEntrada.MensagemRetornoBG
                
                    pstrNomeFila = "A8Q.E.ENTRADA_CC"
                
            End Select
        
        End If
    
    End If
    
    Set xmlMensagemPostada = Nothing
    
    strMessageDescriptor = ""
    strCorrelationID = ""
   
    With pxmlMensagem.documentElement
        
        If Left(UCase(Trim(pstrNomeFila)), 3) <> "AEA" Then
            strMensagemTraduzida = .selectSingleNode("TX_CNTD_SAID").Text
        Else
            strMensagemTraduzida = flObterMensagemBG(.selectSingleNode("TX_CNTD_SAID").Text)
        End If
                        
        objMQAX200.MQOpen pstrNomeFila, MQOO_OUTPUT
        
        If Left(UCase(Trim(pstrNomeFila)), 2) = "NZ" Or _
           Left(UCase(Trim(pstrNomeFila)), 2) = "PZ" Or _
           Left(UCase(Trim(pstrNomeFila)), 2) = "DV" Or _
           Trim(UCase(pstrNomeFila)) = "A7Q.E.SAIDAMV" Or _
           Trim(UCase(pstrNomeFila)) = "A7Q.E.SAIDASIG" Then
                
            objMQAX200.MQPut strMensagemTraduzida, _
                             Len(strMensagemTraduzida), _
                             enumPutOptions.MainFrame, _
                             "A7Q.E.REPORT", _
                             "", _
                             strMessageId, , , _
                             9
        
        ElseIf Left(UCase(Trim(pstrNomeFila)), 3) = "AEA" Then

            objMQAX200.MQPut strMensagemTraduzida, _
                             Len(strMensagemTraduzida), _
                             enumPutOptions.MainFrame, _
                              , _
                             "", _
                             strMessageId
                             
        Else
        
            objMQAX200.MQPut strMensagemTraduzida, _
                             Len(strMensagemTraduzida), _
                             enumPutOptions.Binario, _
                             "A7Q.E.REPORT", _
                             strCorrelationID, _
                             strMessageId, , , _
                             lngPrioridade
        End If
        
        objMQAX200.MQClose
            
        Call flAtualizaMessageIdHex(CLng(.selectSingleNode("CO_MESG").Text), strMessageId)
        
        If Left(UCase(Trim(pstrNomeFila)), 2) = "DV" Or _
           Left(UCase(Trim(pstrNomeFila)), 3) = "AEA" Then
            
            If Not .selectSingleNode("TX_CNTD_ENTR") Is Nothing Then
                Call flAtualizaControleIntegracaoDV(Mid(.selectSingleNode("TX_CNTD_ENTR").Text, 21), strMessageId)
            End If
            
        End If
                
        Call flSalvarSituacaoMensagem(pxmlMensagem, enumOcorrencia.PostagemBemSucedida)
    
    End With
    
    Exit Function

ErrorHandler:
    
    Set xmlMensagemPostada = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPostarMQSeries Function #NK#", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Atualizar o MessageID da mensagem enviada.
Private Function flAtualizaMessageIdHex(ByVal plngCodigoMensagem As Long, _
                                        ByVal pstrMessageIdHex As String) As Boolean
                                                 
Dim objMensagem                       As A7Server.clsMensagem

On Error GoTo ErrorHandler
    
    Set objMensagem = CreateObject("A7Server.clsMensagem")
    Call objMensagem.AtualizaMessageID(plngCodigoMensagem, pstrMessageIdHex)
    Set objMensagem = Nothing

    Exit Function
ErrorHandler:

    Set objMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAtualizaMessageIdHex Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAtualizaControleIntegracaoDV(ByVal pxmlMensagemEntrada As String, _
                                                ByVal pstrMessageId As String) As Boolean
                                                 
Dim objLancamentoCCContabil                   As Object ' A8LQS.clsLancamentoCCContabil
Dim xmlMensagem                               As MSXML2.DOMDocument40
Dim lngTipoLancamentoIntegracao             As Long
    
On Error GoTo ErrorHandler
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML pxmlMensagemEntrada
    
    lngTipoLancamentoIntegracao = 1
    If Not xmlMensagem.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
        lngTipoLancamentoIntegracao = Val(xmlMensagem.selectSingleNode("//TP_LANC_ITGR").Text)
    End If
    
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    Call objLancamentoCCContabil.GravarControleIntegracao(xmlMensagem.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                          lngTipoLancamentoIntegracao, _
                                                          Trim(pstrMessageId))
    
    Set objLancamentoCCContabil = Nothing
    Set xmlMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objLancamentoCCContabil = Nothing
    Set xmlMensagem = Nothing
    
    Err.Clear

End Function

'Obter a mensagem de integração com o sistema BG através do correlationID.
Private Function flObterMensagemBG(ByVal pstrMensagemTraduzida As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML pstrMensagemTraduzida
    
    flObterMensagemBG = xmlMensagem.selectSingleNode("//requestMsg").xml
    
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:

    Set xmlMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterMensagemBG Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar teste de conexão a câmaras participantes do SPB através da mensagem GEN0001.
Public Sub EnviarTesteConexao(ByVal pstrMensagem As String)

Dim objMQAX200                              As A6A7A8CA.clsMQAX200

On Error GoTo ErrorHandler
    
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
     
    With objMQAX200
        .MQInit
        .MQOpen "A7Q.E.ENTRADA", MQOO_OUTPUT
        .MQPut pstrMensagem, Len(pstrMensagem), Binario
        .MQClose
        .MQEnd
    End With
     
    Set objMQAX200 = Nothing
         
    Exit Sub
ErrorHandler:
    
    Set objMQAX200 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarTesteConexao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub



'Enviar mensagem de rejeição para o Legado
'   - Montar protocolo de integração A7
'   - Montar Remessa com os erros identificados

Public Function EnviarMensagemRejeicaoLegado(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByVal plngCodigoOcorrencia As Long) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim intCount                                As Integer
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objErro                                 As A6A7A8CA.clsLogErro
Dim lngTipoMensagem                         As Long
Dim xmlMensagemSaida                        As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strTipoMensagemRetorno                  As String
Dim strSiglaSistemaOrigem                   As String

On Error GoTo ErrorHandler
    
    Set xmlMensagemSaida = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagemSaida.loadXML xmlMensagem.selectSingleNode("//TX_CNTD_SAID").Text
    xmlMensagemSaida.loadXML xmlMensagemSaida.selectSingleNode("//SaidaXML").xml
    
    
    strSiglaSistemaOrigem = UCase$(Trim$(xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text))
    
    If strSiglaSistemaOrigem = "WZ" Then
        
        If Not xmlMensagemSaida.selectSingleNode("//TP_SOLI") Is Nothing Then
            
            If Trim$(xmlMensagemSaida.selectSingleNode("//TP_SOLI").Text) <> vbNullString And _
               IsNumeric(xmlMensagemSaida.selectSingleNode("//TP_SOLI").Text) Then
                
                If Val(xmlMensagemSaida.selectSingleNode("//TP_SOLI").Text) = enumTipoSolicitacao.Inclusao Then
                    Exit Function
                End If
                
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If


    Select Case strSiglaSistemaOrigem
        Case "A8", "NZ", "BG", "DV"
            Exit Function
    End Select

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")

    udtProtocolo.CodigoEmpresa = Format(CLng(xmlMensagem.documentElement.selectSingleNode("//CO_EMPR_ORIG").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = UCase$(fgCompletaString(xmlMensagem.documentElement.selectSingleNode("//SG_SIST_ORIG").Text, " ", 3, False))
    udtProtocolo.SiglaSistemaOrigem = UCase$(fgCompletaString("A8", " ", 3, False))
    
    
    fgAppendNode xmlMensagemSaida, _
                 xmlMensagemSaida.documentElement.nodeName, _
                 "TP_MESG", _
                 Val(xmlMensagem.selectSingleNode("//TP_MESG").Text), _
                 ""
    
    Call objTipoOperacao.ObterTipoOperacaoPorMsgBUS(xmlMensagemSaida, lngTipoMensagem, "", strTipoMensagemRetorno)
    
    lngTipoMensagem = Val(strTipoMensagemRetorno)
    
    If Trim$(udtProtocolo.SiglaSistemaDestino) = "WZ" Then
        lngTipoMensagem = lngTipoMensagem + 2000
    Else
        lngTipoMensagem = lngTipoMensagem
    End If
    
    udtProtocolo.TipoMensagem = fgCompletaString(lngTipoMensagem, "0", 9, True)
    
    LSet udtProtocoloAux = udtProtocolo

    fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "TP_RETN", "2"
    fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "CO_ULTI_SITU_PROC", "99"

    Set objErro = CreateObject("A6A7A8CA.clsLogErro")

    Select Case plngCodigoOcorrencia
        
        Case enumOcorrencia.CanceladaErroTradução
            'Erro de Tradução no sistema A7.
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "CO_ERRO1", 1017
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "DE_ERRO1", xmlMensagem.selectSingleNode("//TX_DTLH_OCOR_ERRO").Text
        
        Case enumOcorrencia.RejeicaoNaoAutenticidade
        
            'Erro de não autenticidade no sistema A7.
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "CO_ERRO1", 1018
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "DE_ERRO1", objErro.ObterDescErroNegocio(1018)
    
    End Select
    
    Set objErro = Nothing
    
    objMQAX200.MQOpen "A7Q.E.ENTRADA", enumMQOO_Open.MQOO_OUTPUT
    
    objMQAX200.MQPut udtProtocoloAux.String & xmlMensagemSaida.xml, _
                     Len(udtProtocoloAux.String & xmlMensagemSaida.xml), _
                     enumPutOptions.Binario
    
    objMQAX200.MQClose

    Exit Function
ErrorHandler:
    
    Set objTipoOperacao = Nothing
    Set objErro = Nothing
    
    Err.Clear

'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemRejeicao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



