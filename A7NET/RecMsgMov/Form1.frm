VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2295
      Left            =   720
      TabIndex        =   0
      Top             =   390
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Dim o As New A6SubReserva.clsRemessa
'Dim o As New A7Server.clsGerenciadorRecebimento
'Dim p As New A7Server.clsGerenciadorEnvio
Dim o As New A8LQS.clsRemessa
'Dim o As New A8LQS.clsTRemessa

Dim strLogErro As String
Dim strMensagem As String
Dim x As Long
Dim strCorrelID As String
Dim xmlRemessa As MSXML2.DOMDocument40
'Dim p As New A7Server.clsGerenciadorEnvio

    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
'    o.ReceberMensagemMQ "A6Q.E.REMESSASUBRESERVA", strLogErro, strMensagem, x, strCorrelID
    'o.ReceberMensagemMQ "A7Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
'    p.ReceberMensagemMQ "A7Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
'    o.ReceberMensagemMQ "A6Q.E.REMESSASUBRESERVA", strLogErro, strMensagem, x, strCorrelID
    o.ReceberMensagemMQ "A8Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
'    p.ReceberMensagemMQ "A7Q.E.ENTRADA", strLogErro, strMensagem, x, strCorrelID
'    o.ReceberMensagemMQ "A8Q.E.ENTRADA_BACEN", strLogErro, strMensagem, x, strCorrelID
'    o.ReceberMensagemMQ "A6Q.E.ERRO", strLogErro, strMensagem, x, strCorrelID
'    o.ReceberMensagemMQ "A6Q.E.ERRO", strLogErro, strMensagem, x, strCorrelID, xmlRemessa
'    o.ReceberMensagemMQ "A8Q.E.ENTRADA_BOLETO", strLogErro, strMensagem, x, strCorrelID
'    o.ReceberMensagemMQ "A8Q.E.ENTRADA_OPERACAO", strLogErro, strMensagem, x, strCorrelID
   ' o.ReceberMensagemMQ "A8Q.E.MSG_R1", strLogErro, strMensagem, x, strCorrelID
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    
End Sub

Public Sub Executar()

Dim objExecutar                             As New A8LQS.clsLiquidacaoFutura

    Set objExecutar = Nothing

End Sub
