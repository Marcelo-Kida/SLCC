VERSION 5.00
Begin VB.Form frmVerificaServer 
   Caption         =   "A7 - Verifica Server"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   Icon            =   "frmVerificaServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFuncoesVerificaServer 
      Height          =   1860
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmVerificaServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        
    lstFuncoesVerificaServer.AddItem "Verificar MQSeries"
    lstFuncoesVerificaServer.AddItem "Verificar Banco de Dados SLCC"
    lstFuncoesVerificaServer.AddItem "Verificar Banco de Dados PJ/PK"
    
    DoEvents
    Me.Show
       
End Sub

