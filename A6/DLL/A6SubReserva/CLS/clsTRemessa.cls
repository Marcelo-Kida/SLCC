VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo
' tratamento em geral de remessas enviadas ao A6.

Option Explicit

Public objValidaRemessa                     As Object 'A6A8ValidaRemessa.clsValidaRemessa

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Implements ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pelo recebimento de mensagens MQ.

Public Function ReceberMensagemMQ(ByVal pstrQName As String, _
                                  ByRef pstrLogErro As String, _
                                  ByRef pstrMessageIDHexOut As String, _
                                  ByRef plngBackOutCount As Long, _
                                  ByRef pstrCorrelationID As String, _
                                  ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim strMessageDescriptor                    As String
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim strErros                                As String
Dim blnEnviaParaRemesssaRejeitada           As Boolean

On Error GoTo ErrorHandler

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")

'    fgGravaArquivo "LOGA6_clsTRemessa_ReceberMensagemMQ_", vbNullString
    
    objMQAX200.MQInit
    objMQAX200.MQOpen pstrQName, MQOO_INPUT_SHARED

    lngBufferLen = Len(strBuffer)
    If objMQAX200.MQGet(lngBufferLen, _
                        strBuffer, _
                        True, _
                        Binario, _
                        pstrMessageIDHexOut, _
                        pstrCorrelationID, _
                        strMessageDescriptor, _
                        plngBackOutCount) Then
        
        objMQAX200.MQClose

'       Chamada do ValidarRemessa
        If flValidarRemessa(strBuffer, xmlRemessa, strErros) Then
            flVerificarRemessa xmlRemessa, strErros
        End If

        If strErros <> vbNullString Then

             Err.Description = strErros
            GoTo ErrorHandler

        End If

        ReceberMensagemMQ = True
    Else
        ReceberMensagemMQ = False
    End If

    objMQAX200.MQClose
    objMQAX200.MQEnd

    Set objMQAX200 = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    pstrLogErro = Err.Description
    
    Set objMQAX200 = Nothing

    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

End Function

' Este método é responsável pela validação da remessa (ativa ValidaRemessa.exe ActiveX Component).

Private Function flValidarRemessa(ByRef pstrBuffer As String, _
                                  ByRef pxmlRemessa As MSXML2.DOMDocument40, _
                                  ByRef pstrErros As String) As Boolean

'Dim objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa

On Error GoTo ErrorHandler

    Set pxmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    
    'If objValidaRemessa Is Nothing Then
        Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
    'End If
    
    flValidarRemessa = False
    
    If pxmlRemessa.loadXML(pstrBuffer) Then
    
'        fgGravaArquivo "LOGA6_clsTRemessa_flValidarRemessa_", vbNullString
    
        pstrErros = objValidaRemessa.ValidarMensagemA6(pstrBuffer)
        Set objValidaRemessa = Nothing

        If pstrErros = vbNullString Then
            Call pxmlRemessa.loadXML(pstrBuffer)
            flValidarRemessa = True
        End If
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    

Exit Function
ErrorHandler:

    flValidarRemessa = False
    
    Set objValidaRemessa = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarRemessa Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pelo encaminhamento da mensagem MQ, conforme o tipo da remessa.

Public Function flVerificarRemessa(ByRef pxmlRemessa As MSXML2.DOMDocument40, _
                                   ByRef pstrErro As String) As Boolean

Dim objControleRemessa                      As A6SubReserva.clsControleRemessa
Dim objMovimentoFuturo                      As A6SubReserva.clsMovimentoFuturo
Dim objMovimentoSubReserva                  As A6SubReserva.clsMovimentoSubReserva
Dim objSaldoFechSubReserva                  As A6SubReserva.clsSaldoFechSubReserva

    On Error GoTo ErrorHandler

    With pxmlRemessa.documentElement
        
'        fgGravaArquivo "LOGA6_clsTRemessa_flVerificarRemessa_TP_MESG_" & .selectSingleNode("TP_MESG").Text, vbNullString
    
        If Val(.selectSingleNode("TP_MESG").Text) = enumEvento.EstatisticaCaixaFuturo Then
            
            Set objControleRemessa = CreateObject("A6SubReserva.clsControleRemessa")
            Call objControleRemessa.Salvar(pxmlRemessa)
            Set objControleRemessa = Nothing
        
        ElseIf Val(.selectSingleNode("TP_MESG").Text) = enumEvento.DetalheOperacaoCaixaFuturo Then
            
            Set objMovimentoFuturo = CreateObject("A6SubReserva.clsMovimentoFuturo")
            Call objMovimentoFuturo.Salvar(pxmlRemessa)
            Set objMovimentoFuturo = Nothing
        
            If Not .selectSingleNode("CO_ITEM_CAIX_SUB_RESE") Is Nothing Then
                Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
                Call objMovimentoSubReserva.Salvar(pxmlRemessa)
                Set objMovimentoSubReserva = Nothing
            End If
        
        ElseIf Val(.selectSingleNode("TP_MESG").Text) = enumEvento.CaixaSubReservaSaldoFechamento Then
            
            Set objSaldoFechSubReserva = CreateObject("A6SubReserva.clsSaldoFechSubReserva")
            If .selectSingleNode("TP_SINA_SALD_FECH").Text = "2" Then
                .selectSingleNode("VA_SALD_FECH").Text = fgVlr_To_Xml(fgVlrXml_To_Interface(.selectSingleNode("VA_SALD_FECH").Text) * -1)
            End If
            Call objSaldoFechSubReserva.Salvar(pxmlRemessa)
            Set objSaldoFechSubReserva = Nothing
        
        ElseIf Val(.selectSingleNode("TP_MESG").Text) = enumEvento.MovimentoSubReserva Then
        
            If .selectSingleNode("DE_ATIV_MERC") Is Nothing Then
                Call fgAppendNode(pxmlRemessa, .baseName, "DE_ATIV_MERC", "")
            End If
            
            If .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text <> enumTipoMovimento.PrevistoCompromIda And _
               .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text <> enumTipoMovimento.EstornoPrevistoCompromIda Then
               
                Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
            
                If .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.RealizadoConfirmado Then
                    If objMovimentoSubReserva.AtualizarSituacaoMovimento(pxmlRemessa) = 0 Then
                        Call objMovimentoSubReserva.Salvar(pxmlRemessa)
                    End If
                Else
                    Call objMovimentoSubReserva.Salvar(pxmlRemessa)
                End If
                
                Set objMovimentoSubReserva = Nothing
                
            End If
        
            Set objMovimentoFuturo = CreateObject("A6SubReserva.clsMovimentoFuturo")
            If .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.Previsto Or _
               .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.PrevistoCompromIda Then
                
                Call objMovimentoFuturo.Salvar(pxmlRemessa)
            
            ElseIf .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoPrevisto Or _
                   .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoPrevistoCompromIda Then
                
                .selectSingleNode("//VA_OPER_ATIV").Text = "-" & Trim$(.selectSingleNode("//VA_OPER_ATIV").Text)
                Call objMovimentoFuturo.Salvar(pxmlRemessa)
            
            End If
            Set objMovimentoFuturo = Nothing
        
        End If
    End With

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    pstrErro = fgLimpaCaracterEspecial(Err.Description)
    
    Set objControleRemessa = Nothing
    Set objMovimentoFuturo = Nothing
    Set objSaldoFechSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

End Function

'Inclusão da remessa na tabela de remessa rejeitada e geração de alerta

Public Function RemessaRejeitada(ByVal pstrQName As String, _
                                 ByRef pstrErros As String, _
                                 ByRef pstrMessageIDHexOut As String, _
                                 ByRef plngBackOutCount As Long, _
                                 ByVal xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim objRemessaRejeitada                     As A6SubReserva.clsRemessaRejeitada
Dim xmlErros                                As MSXML2.DOMDocument40
Dim xmlRemessaRejeitada                     As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strMessageDescriptor                    As String
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim strErros                                As String


On Error GoTo ErrorHandler

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")

    objMQAX200.MQInit
    objMQAX200.MQOpen pstrQName, MQOO_INPUT_SHARED

    lngBufferLen = Len(strBuffer)
    
    If objMQAX200.MQGet(lngBufferLen, _
                        strBuffer, _
                        True, _
                        enumPutOptions.Binario, _
                        pstrMessageIDHexOut, _
                        vbNullString, _
                        strMessageDescriptor, _
                        plngBackOutCount) Then

        objMQAX200.MQClose
        

        Set xmlRemessaRejeitada = CreateObject("MSXML2.DOMDocument.4.0")
        
        fgAppendNode xmlRemessaRejeitada, "", "REMESSAREJEITADA", ""
        fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "TX_XML_REME", ""
        
        Set xmlNode = xmlRemessaRejeitada.documentElement.selectSingleNode("TX_XML_REME")
        xmlNode.appendChild fgCreateCDATASection(xmlRemessa.xml)
        
        fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "TX_XML_ERRO", ""
        
        Set xmlNode = xmlRemessaRejeitada.documentElement.selectSingleNode("TX_XML_ERRO")
        xmlNode.appendChild fgCreateCDATASection(pstrErros)
        
        fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "TP_MESG", xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
        fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "CO_EMPR", xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text
        fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "SG_SIST_INFO", xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text
        fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "CO_VEIC_LEGA", xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text

        Set objRemessaRejeitada = CreateObject("A6SubReserva.clsRemessaRejeitada")
        objRemessaRejeitada.Salvar xmlRemessaRejeitada

        Set xmlRemessaRejeitada = Nothing
        Set objRemessaRejeitada = Nothing
            
        'Pikachu - 26/01/2005
        'Demanada 1025
        Call flEnviarEmail(xmlRemessa)
            
    End If

    pstrErros = vbNullString

    objMQAX200.MQEnd
    Set objMQAX200 = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    pstrErros = Err.Description

    Set xmlRemessaRejeitada = Nothing
    Set objRemessaRejeitada = Nothing

    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If
    
End Function

Private Function flEnviarEmail(ByRef pxmlRemessa As MSXML2.DOMDocument40)

Dim objAlerta                               As Object 'A8LQS.clsAlerta

On Error GoTo ErroHandler

    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    Call objAlerta.GerarAlertaA6(pxmlRemessa, _
                                 enumFatorGeradorAlerta.RemessaRejeitadaA6)
    Set objAlerta = Nothing
   
    Exit Function
ErroHandler:
    
    Set objAlerta = Nothing
    
    Err.Clear
    
End Function

