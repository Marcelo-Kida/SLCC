VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim o As New A6A7A8ServicoCOM.clsProcesso_Batch
'Dim o As New A6A7A8ServicoCOM.clsProcesso_A7



o.Processar Clipboard.GetText, Format(DateSerial(1900, 1, 1), "YYYYMMDDHHmmss")
'o.ReceberMensagemMQ Clipboard.GetText, Format(DateSerial(1900, 1, 1), "YYYYMMDDHHmmss")


End Sub
