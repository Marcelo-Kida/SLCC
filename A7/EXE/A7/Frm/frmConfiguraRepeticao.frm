VERSION 5.00
Begin VB.Form frmConfiguraRepeticao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   2465
      TabIndex        =   2
      Top             =   270
      Width           =   465
   End
   Begin VB.TextBox txtQtdRepe 
      Height          =   285
      Left            =   45
      MaxLength       =   4
      TabIndex        =   1
      Top             =   270
      Width           =   2355
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2985
   End
   Begin VB.Label lblCriterio 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade de Repetições:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   1950
   End
End
Attribute VB_Name = "frmConfiguraRepeticao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto auxiliar, reponsável pela configuração da quantidade de repetições de um atributo em uma mensagem.
'Utilizado no formulário de cadastrop de tipos de mensagem.
Option Explicit

Public plngQtdRepe                          As Long
Public Event QuantidadeRepeticoes(ByVal plngQtdRepeticao As Long)

Private Sub Form_Activate()
    
    txtQtdRepe.Text = plngQtdRepe
    txtQtdRepe.SetFocus
    SendKeys "{Home}+{End}"
        
End Sub

Private Sub Form_LostFocus()
    txtQtdRepe.Text = vbNullString
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    RaiseEvent QuantidadeRepeticoes(CLng(IIf(txtQtdRepe.Text = vbNullString, 0, txtQtdRepe.Text)))
    Me.Hide

End Sub

Private Sub txtQtdRepe_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Case 13
            RaiseEvent QuantidadeRepeticoes(CLng(IIf(txtQtdRepe.Text = vbNullString, 0, txtQtdRepe.Text)))
            Me.Hide
        Case Else
            KeyAscii = 0
    End Select
    
End Sub
