VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Objeto responsável pelo envio de e-mails a usuários.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

Private strConstruct                         As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31
Private Const NORMAL_PRIORITY_CLASS = &H20&

Implements COMSVCSLib.IObjectConstruct

'Obter atrvés do object contruct o nome do Servidor de SMTP

Private Sub IObjectConstruct_Construct(ByVal pobjConstruct As Object)

On Error GoTo ErrHandler
    
    strConstruct = pobjConstruct.ConstructString
    
    Exit Sub
    
ErrHandler:
    Err.Clear
End Sub

'Enviar e-mail
'Definição do XML de Email
'<Email>
'   <EmailTo></EmailTo>                 'Lista de Email separados por ;
'   <EmailSender></EmailSender>         'Nome do Computador (Servidor)
'   <EmailSubject></EmailSubject>       'Asssunto do Email
'   <EmailTextBody></EmailTextBody>     'Texto do e-mail
'</Email>

Public Sub EnviaEmail(ByVal pstrXMLEmail As String)


Dim objCDOMessage                           As Object
Dim objEmail                                As MSXML2.DOMDocument40

On Error GoTo ErrHandler
    
    Set objEmail = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objEmail.loadXML(pstrXMLEmail) Then
        'Call fgErroLoadXML(objEmail, App.EXEName, TypeName(Me), "EnviaEmail")
    End If
    
    Set objCDOMessage = CreateObject("CDO.Message")
    
    With objCDOMessage.Configuration
        .Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  'cdoSendUsingMethod 'cdoSendUsingPort
        
        If Len(strConstruct) > 0 Then
            .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strConstruct 'cdoSMTPServer
        Else
            .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "192.168.255.1" 'cdoSMTPServer
        End If

        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 'cdoSMTPServerPort
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0 'cdoSMTPAuthenticate 'cdoAnonymous
        .Fields.Update
    End With
    
    With objCDOMessage
        .To = objEmail.documentElement.selectSingleNode("EmailTo").Text
        .From = objEmail.documentElement.selectSingleNode("EmailSender").Text
        .Sender = flObterEstacaoTrabalho & "@santander.com.br"
        .Subject = objEmail.documentElement.selectSingleNode("EmailSubject").Text
        .TextBody = objEmail.documentElement.selectSingleNode("EmailTextBody").Text
        .send
    End With
    
    Set objEmail = Nothing
    Set objCDOMessage = Nothing
    
    Exit Sub

ErrHandler:
    
    Set objEmail = Nothing
    Set objCDOMessage = Nothing
    
    Err.Clear
    
End Sub


'Obter o nome do Servidor

Private Function flObterEstacaoTrabalho() As String

Dim strEstacao                               As String
Dim lngLen                                   As Long

On Error GoTo ErrorHandler
    
    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    strEstacao = String(lngLen, "X")
    
    GetComputerName strEstacao, lngLen
    strEstacao = Left(strEstacao, lngLen)
    
    flObterEstacaoTrabalho = strEstacao

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "modStartScreen", "flObterEstacaoTrabalho Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

