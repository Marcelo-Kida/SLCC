Attribute VB_Name = "Module1"
Option Explicit

Public Sub MAIN()

Dim strLogErro As String
Dim strMensagem As String
Dim x As Long
Dim strCorrelID As String
Dim xmlRemessa As MSXML2.DOMDocument40
Dim o As New A8LQS.clsRemessa
Dim o As New A7Server.clsGerenciadorRecebimento
Dim p As New A7Server.clsGerenciadorEnvio

    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
'    o.ReceberMensagemMQ "A6Q.E.REMESSASUBRESERVA", strLogErro, strMensagem, x, strCorrelID
    'o.ReceberMensagemMQ "A7Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
    'p.ReceberMensagemMQ "A7Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
    o.ReceberMensagemMQ "A8Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
'    o.ReceberMensagemMQ "A7Q.E.ERRO", strLogErro, strMensagem, x, strCorrelID
'    o.ReceberMensagemMQ "A7Q.E.ERRO", strLogErro, strMensagem, x, strCorrelID, xmlRemessa
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description

End Sub


