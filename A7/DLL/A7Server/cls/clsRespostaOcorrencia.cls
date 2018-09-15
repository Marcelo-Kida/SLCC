VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRespostaOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo envio de respostas de ocorrências e envia e-mail de notificação da ocorrência.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

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

'Enviar resposta de ocorrência configuradas nos parâmetros de notificação de ororrência.
Public Function EnviarRespostaOcorrencia(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                         ByVal plngCodigoOcorrencia As Long) As Boolean

Dim xmlParamRespostaOcorrencia              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlParamRespostaOcorrencia = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParamRespostaOcorrencia.loadXML flObterParamRespostaOcorrencia(pxmlMensagem, plngCodigoOcorrencia)
            
    'Verifica se enciar resposta a ocorrencia
    If xmlParamRespostaOcorrencia.documentElement.selectSingleNode("Grupo_RespostaOcorrencia/IN_ENVI_RSPT").Text = enumIndicadorSimNao.Sim Then
        flEnviarRespostaOcorrencia pxmlMensagem, xmlParamRespostaOcorrencia
    End If
        
    'Verifica se envia e-mail a ocorrencia
    If CLng(xmlParamRespostaOcorrencia.documentElement.selectSingleNode("Grupo_Email/IN_ENVI_EMAIL").Text) = enumIndicadorSimNao.Sim Then
        flEnviarEmail pxmlMensagem, xmlParamRespostaOcorrencia
    End If
        
    Set xmlParamRespostaOcorrencia = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlParamRespostaOcorrencia = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRespostaOcorrencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obtém os parmâmetros da resposta a ocorrência.
Private Function flObterParamRespostaOcorrencia(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                                ByVal plngCodigoOcorrencia As Long) As String

Dim strSQL                                  As String
Dim strLer                                  As String
Dim xmlParamRespostaOcorrencia              As MSXML2.DOMDocument40
Dim objRS                                   As ADODB.Recordset


On Error GoTo ErrorHandler
    
    Set xmlParamRespostaOcorrencia = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlParamRespostaOcorrencia, "", "Grupo_ParamResposta", "")
        
    Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_ParamResposta", "Grupo_RespostaOcorrencia", "")
    
    strSQL = " SELECT  A.NO_FILA_MQSE " & _
             "   FROM  A7.TB_ENDE_FILA_MQSE A, " & _
             "         A7.TB_RSPT_OCOR_SIST B " & _
             "  WHERE  A.SG_SIST_DEST = B.SG_SIST " & _
             "    AND  A.CO_EMPR_DEST = B.CO_EMPR " & _
             "    AND  B.SG_SIST = '" & Trim(pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text) & "'" & _
             "    AND  B.CO_EMPR =  " & CLng(pxmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text) & _
             "    AND  B.CO_OCOR_MESG = " & plngCodigoOcorrencia
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_RespostaOcorrencia", "IN_ENVI_RSPT", enumIndicadorSimNao.Nao)
        Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_RespostaOcorrencia", "NO_FILA_MQSE", "")
    Else
        Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_RespostaOcorrencia", "IN_ENVI_RSPT", enumIndicadorSimNao.Sim)
        Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_RespostaOcorrencia", "NO_FILA_MQSE", IIf(IsNull(objRS.fields(0)), "", objRS.fields(0)))
    End If
    
    objRS.Close
    
    strSQL = " SELECT CO_OCOR_MESG      , " & _
             "        DE_OCOR_MESG      , " & _
             "        DE_ABRV_OCOR_MESG , " & _
             "        IN_ENVI_EMAIL     , " & _
             "        TX_ENDE_EMAIL_NOTI  " & _
             "   FROM A7.TB_OCOR_MESG     " & _
             "  WHERE CO_OCOR_MESG = " & plngCodigoOcorrencia
    
    Set objRS = fgQuerySQL(strSQL)
    
    Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_ParamResposta", "Grupo_Email", "")
    
    Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_Email", "CO_OCOR_MESG", Trim(objRS!CO_OCOR_MESG))
    Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_Email", "DE_OCOR_MESG", Trim(objRS!DE_OCOR_MESG))
    Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_Email", "DE_ABRV_OCOR_MESG", Trim(objRS!DE_ABRV_OCOR_MESG))
    
    If UCase(Trim(objRS!IN_ENVI_EMAIL)) = "S" Then
        Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_Email", "IN_ENVI_EMAIL", enumIndicadorSimNao.Sim)
    Else
        Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_Email", "IN_ENVI_EMAIL", enumIndicadorSimNao.Nao)
    End If
    
    Call fgAppendNode(xmlParamRespostaOcorrencia, "Grupo_Email", "TX_ENDE_EMAIL_NOTI", IIf(IsNull(objRS!TX_ENDE_EMAIL_NOTI), "", Trim(objRS!TX_ENDE_EMAIL_NOTI)))
        
    flObterParamRespostaOcorrencia = xmlParamRespostaOcorrencia.xml
    
    Set xmlParamRespostaOcorrencia = Nothing
    
    Exit Function
ErrorHandler:

    Set xmlParamRespostaOcorrencia = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterParamRespostaOcorrencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar email de notificação da ocorrência.
Private Function flEnviarEmail(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                               ByRef pxmlRespostaOcorrencia As MSXML2.DOMDocument40) As String

'<Email>
'   <EmailTo></EmailTo>                 'Lista de Email separados por ;
'   <EmailSender></EmailSender>         'Nome do Computador (Servidor)
'   <EmailSubject></EmailSubject>       'Asssunto do Email
'   <EmailTextBody></EmailTextBody>     'Texto do e-mail
'</Email>

Dim xmlEmail                                As MSXML2.DOMDocument40
Dim objEmail                                As A6A7A8CA.clsEmail
Dim strTextBody                             As String

On Error GoTo ErrorHandler
    
    Set xmlEmail = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlEmail, "", "Email", "")
    Call fgAppendNode(xmlEmail, "Email", "EmailTo", pxmlRespostaOcorrencia.documentElement.selectSingleNode("Grupo_Email/TX_ENDE_EMAIL_NOTI").Text)
    Call fgAppendNode(xmlEmail, "Email", "EmailSender", "")
    
    If pxmlRespostaOcorrencia.selectSingleNode("//CO_OCOR_MESG").Text = enumOcorrencia.UsuarioInativo Then
        Call fgAppendNode(xmlEmail, "Email", "EmailSubject", "Relatório de Usuários Inativos - SLCC")
    Else
        Call fgAppendNode(xmlEmail, "Email", "EmailSubject", "Notificação de Ocorrência - Sistema A7 - Bus de Interface")
    End If
    
    strTextBody = flMontaMensagemEmail(pxmlMensagem, pxmlRespostaOcorrencia)
    
    Call fgAppendNode(xmlEmail, "Email", "EmailTextBody", strTextBody)
    
    Set objEmail = CreateObject("A6A7A8CA.clsEmail")
    
    objEmail.EnviaEmail xmlEmail.xml
    
    Set objEmail = Nothing
    Set xmlEmail = Nothing

    Exit Function
ErrorHandler:
    
    Set objEmail = Nothing
    Set xmlEmail = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarEmail Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar notificação de ocorrência através de fila MQSeries.
Private Function flEnviarRespostaOcorrencia(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                            ByRef pxmlRespostaOcorrencia As MSXML2.DOMDocument40) As String

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim strMessageDescriptor                    As String
Dim strCorrelationID                        As String
Dim strMessageId                            As String
Dim strMensagemOcorrencia                   As String

On Error GoTo ErrorHandler
    
    strMensagemOcorrencia = flMontaMensagemOcorrencia(pxmlMensagem, pxmlRespostaOcorrencia)

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    
    objMQAX200.MQInit
    objMQAX200.MQOpen pxmlRespostaOcorrencia.documentElement.selectSingleNode("Grupo_RespostaOcorrencia/NO_FILA_MQSE").Text, MQOO_OUTPUT
    
    objMQAX200.MQPut strMensagemOcorrencia, _
                     Len(strMensagemOcorrencia), _
                     enumPutOptions.Binario, _
                     "A7Q.E.REPORT", _
                     strCorrelationID, _
                     strMessageId, _
                     strMessageDescriptor
    
    objMQAX200.MQClose
    
    Set objMQAX200 = Nothing
     
    Exit Function

ErrorHandler:

    Set objMQAX200 = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRespostaOcorrencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar mensagem de notificação de ororrência.
Private Function flMontaMensagemOcorrencia(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pxmlRespostaOcorrencia As MSXML2.DOMDocument40) As String
                                

Dim xmlResposta                             As MSXML2.DOMDocument40
Dim strProtocolo                            As String
Dim udtProtocolo                            As udtProtocolo
Dim strMensagemEntrada                      As String

On Error GoTo ErrorHandler
    
    Set xmlResposta = CreateObject("MSXML2.DOMDocument.4.0")
        
    strMensagemEntrada = pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR").Text
    strProtocolo = Mid(strMensagemEntrada, 1, Len(udtProtocolo))
    
    Call fgAppendNode(xmlResposta, "", "MESG", "")
    Call fgAppendNode(xmlResposta, "MESG", "TX_PRTC", strProtocolo)
    Call fgAppendNode(xmlResposta, "MESG", "CO_MESG", pxmlMensagem.documentElement.selectSingleNode("CO_MESG").Text)
    Call fgAppendNode(xmlResposta, "MESG", "DE_OCOR_MESG", pxmlRespostaOcorrencia.documentElement.selectSingleNode("Grupo_Email/DE_OCOR_MESG").Text)
    Call fgAppendNode(xmlResposta, "MESG", "DE_SITU_MESG", pxmlRespostaOcorrencia.documentElement.selectSingleNode("Grupo_Email/DE_ABRV_OCOR_MESG").Text)
    
    flMontaMensagemOcorrencia = xmlResposta.xml
    
    Set xmlResposta = Nothing
    
    Exit Function
ErrorHandler:
        
    Set xmlResposta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaMensagemOcorrencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar mensagem de ocorrência com formatação para envio via email.
Private Function flMontaMensagemEmail(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                      ByRef pxmlRespostaOcorrencia As MSXML2.DOMDocument40) As String
                                

On Error GoTo ErrorHandler
    
    If pxmlRespostaOcorrencia.selectSingleNode("//CO_OCOR_MESG").Text = enumOcorrencia.UsuarioInativo Then
    
        flMontaMensagemEmail = _
        " " & String$(70, "*") & vbCrLf & _
        " Relatório de Usuários Inativos - SLCC" & vbCrLf & _
        " Data do Relatório     : " & fgDataHoraServidor(DataHoraAux) & vbCrLf & _
        " " & String$(70, "*") & vbCrLf & vbCrLf & _
        pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").Text
    
    Else
        flMontaMensagemEmail = _
        " *********************************************************" & vbCrLf & _
        " Notificação de Ocorrência - Sistema A7 - Bus de Interface" & vbCrLf & _
        " Código da Mensagem    : " & pxmlMensagem.documentElement.selectSingleNode("CO_MESG").Text & vbCrLf & _
        " Ocorrência            : " & pxmlRespostaOcorrencia.documentElement.selectSingleNode("Grupo_Email/DE_OCOR_MESG").Text & vbCrLf & _
        " Complemento           : " & pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").Text & vbCrLf & _
        " *********************************************************"
    End If
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaMensagemEmail Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




