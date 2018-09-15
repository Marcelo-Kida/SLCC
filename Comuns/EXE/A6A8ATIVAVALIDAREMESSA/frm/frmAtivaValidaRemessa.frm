VERSION 5.00
Begin VB.Form frmAtivaValidaRemessa 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAtivaValidaRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa

Private Sub Form_Load()
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
End Sub
