VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF7760B01F4"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelo recebimento de mensagens COA e COD e atualização do status das mensagens transitadas no sistema A7.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Private objMQAX200                           As A6A7A8CA.clsMQAX200
Private ObjectContext                        As COMSVCSLib.ObjectContext
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
    
Dim lngLenFeedBack                           As Long
Dim strFeedBack                              As String
Dim strMessageDescriptor                     As String
Dim blnTemMensagemFila                       As Boolean
Dim lngCont                                  As Long

On Error GoTo ErrorHandler
    
    blnTemMensagemFila = True
     
    lngCont = 0
     
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    Call objMQAX200.MQInit
    Call objMQAX200.MQOpen(pstrNomeFila, MQOO_INPUT_SHARED)
     
    While blnTemMensagemFila
               
        strMessageDescriptor = ""
        strFeedBack = ""
        pstrMessageIdHex = ""
        pstrCorrelationID = ""
        
        lngCont = lngCont + 1
        
        blnTemMensagemFila = objMQAX200.MQGet(lngLenFeedBack, _
                                              strFeedBack, _
                                              True, _
                                              True, _
                                              pstrMessageIdHex, _
                                              pstrCorrelationID, _
                                              strMessageDescriptor, _
                                              plngBackOutCount)
        
        If blnTemMensagemFila Then
            flProcessaEntregaMensagem strFeedBack, strMessageDescriptor
        End If
            
        If lngCont = 100 Then
            blnTemMensagemFila = False
        End If
    Wend
    
    Call objMQAX200.MQEnd
    Set objMQAX200 = Nothing
    
    ReceberMensagemMQ = False
       
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

'Tratar as mensagens recebidas para a atualização das mensagens transitadas.
Private Function flProcessaEntregaMensagem(ByVal pstrMensagemRecebida As String, _
                                           ByVal pstrMessageDescriptor As String)
    
Dim xmlMessageDescriptor                     As MSXML2.DOMDocument40
Dim xmlMensagem                              As MSXML2.DOMDocument40
Dim strMessageIdHex                          As String
Dim strCorrelationIdHex                      As String
Dim strxmlMensagem                           As String
Dim lngFeedBack                              As Long
Dim lngCodigoOcorrencia                      As Long

On Error GoTo ErrorHandler
    
    Set xmlMessageDescriptor = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMessageDescriptor.loadXML pstrMessageDescriptor
        
    With xmlMessageDescriptor.documentElement
        strCorrelationIdHex = .selectSingleNode("CorrelationIdHex").Text
        strMessageIdHex = .selectSingleNode("MessageIdHex").Text
        lngFeedBack = .selectSingleNode("Feedback").Text
    End With
    
    Set xmlMessageDescriptor = Nothing
    
    strxmlMensagem = flObterMensagem(strCorrelationIdHex)
    
    If strxmlMensagem <> vbNullString Then
    
         Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
         
         xmlMensagem.loadXML strxmlMensagem
        
         If lngFeedBack = enumMQFeedBack.MQFB_COA Then
             lngCodigoOcorrencia = enumOcorrencia.ConfirmacaoEntregaRecebida
         ElseIf lngFeedBack = enumMQFeedBack.MQFB_COD Then
             lngCodigoOcorrencia = enumOcorrencia.ConfirmacaoRetiradaRecebida
         Else
              Call flSalvarMensagemRejeitada(pstrMensagemRecebida, _
                                             pstrMessageDescriptor, _
                                             "Mensagem Report não tratada", _
                                             "A7Q.E.ENTREGA")
         End If
         
         flSalvarSituacao xmlMensagem, lngCodigoOcorrencia
            
        End If
        
        Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlMensagem = Nothing
    Set xmlMessageDescriptor = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaEntregaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Inlcui nova situação a mensagem transitada.
Private Function flSalvarSituacao(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarSituacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter a mensagem transitada através do seu correlationID.
Private Function flObterMensagem(ByVal pstrCorrelationIDHex As String) As String
                                                 
Dim strSQL                                  As String
Dim strxmlMensagem                          As String
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagemRecebida                     As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
    
On Error GoTo ErrorHandler

    strSQL = " SELECT  /*+ INDEX(TB_MESG,IE_MESG_01) */ " & _
            "           CO_MESG,                 " & _
            "           DH_MESG,                 " & _
            "           TP_MESG,                 " & _
            "           CO_EMPR_ORIG,            " & _
            "           SG_SIST_ORIG,            " & _
            "           '' AS SG_SIST_DEST,      " & _
            "           DH_INIC_VIGE_REGR_TRAP,  " & _
            "           CO_CMPO_ATRB_IDEF_MESG,  " & _
            "           CO_MESG_MQSE,            " & _
            "           CO_TEXT_XML_ENTR,        " & _
            "           '' AS  TX_DTLH_OCOR_ERRO " & _
            "  FROM     A7.TB_MESG               " & _
            "  WHERE    CO_MESG_MQSE = '" & Trim(pstrCorrelationIDHex) & "'"
            
    strxmlMensagem = fgQueryXMLLer("Grupo_Mensagem", strSQL, App.EXEName & "." & TypeName(Me))
        
            
    If strxmlMensagem = vbNullString Then
        'Não existe mensagem associada a esta mensagem de REPORT
        'lngCodigoErroNegocio = 1012
        'GoTo ErrorHandler
        flObterMensagem = vbNullString
    Else
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagem.loadXML strxmlMensagem
        strMensagemRecebida = fgSelectVarchar4000(xmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_TEXT_XML_ENTR").Text)
        udtProtocoloAux.String = Mid(strMensagemRecebida, 1, Len(udtProtocolo))
        LSet udtProtocolo = udtProtocoloAux
        xmlMensagem.selectSingleNode("//Grupo_Mensagem/SG_SIST_DEST").Text = udtProtocolo.SiglaSistemaDestino
        flObterMensagem = strxmlMensagem
    End If

    Exit Function
ErrorHandler:
   

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterCodigoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar mensagens rejeitadas.
Private Sub flSalvarMensagemRejeitada(ByVal pstrMensagemRecebida As String, _
                                      ByVal pstrMessageDescriptor As String, _
                                      ByVal pstrDetalheOcorrenciaErro As String, _
                             Optional ByVal pstrNomeFila As String = vbNullString)

Dim xmlMessageDescriptor                     As MSXML2.DOMDocument40
Dim objMensagemRejeitada                     As A7Server.clsMensagemRejeitada
Dim objFilaErro                              As A7Server.clsFilaErro

Dim strNomeAplicativo                        As String
Dim strMessageIdHex                          As String
Dim dtmDataPut                               As Date

On Error GoTo ErrorHandler

    Set xmlMessageDescriptor = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMessageDescriptor.loadXML pstrMessageDescriptor
    
    With xmlMessageDescriptor.documentElement
        dtmDataPut = fgDtHrStr_To_DateTime(.selectSingleNode("PutDateTime").Text)
        strNomeAplicativo = Trim(.selectSingleNode("PutApplicationName").Text)
        strMessageIdHex = Trim(.selectSingleNode("MessageIdHex").Text)
    End With
    
    Set xmlMessageDescriptor = Nothing
    
    Set objMensagemRejeitada = CreateObject("A7Server.clsMensagemRejeitada")
    
    pstrMensagemRecebida = IIf(Trim(pstrMensagemRecebida), "", " ")
    
    Call objMensagemRejeitada.SalvarMensagemRejeitada(pstrMensagemRecebida, _
                                                      enumOcorrencia.RejeicaoNaoAutenticidade, _
                                                      strMessageIdHex, _
                                                      strNomeAplicativo, _
                                                      dtmDataPut, _
                                                      pstrDetalheOcorrenciaErro)
    Set objMensagemRejeitada = Nothing

    Set objFilaErro = CreateObject("A7Server.clsFilaErro")
    
    Call objFilaErro.PutFilaErro(pstrNomeFila, _
                                 "A7Q.E.ERRO", _
                                 pstrMensagemRecebida, _
                                 pstrDetalheOcorrenciaErro, _
                                 strMessageIdHex, _
                                 True)
    
    Set objFilaErro = Nothing

    Exit Sub
ErrorHandler:
    
    Set xmlMessageDescriptor = Nothing
    Set objMensagemRejeitada = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarMensagemRejeitada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub


