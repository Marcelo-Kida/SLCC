VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIncluirNumComandoAcao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inclusão Número Comando"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumeroComandoAcao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtNumeroComando 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   2820
      TabIndex        =   1
      Top             =   960
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIncluirNumComandoAcao.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIncluirNumComandoAcao.frx":031A
            Key             =   "Salvar"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNumeroComando 
      Alignment       =   1  'Right Justify
      Caption         =   "Número Comando "
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número Comando"
      Height          =   195
      Left            =   1140
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmIncluirNumComandoAcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Formulário para manutenção do Número de Comando

Public NumeroComando                        As Long
Public NumeroComandoAcao                    As Long
Public Acao                                 As String

Private Sub Form_Load()
On Error GoTo ErrorHandler

    Call fgCursor
    Set Me.Icon = mdiLQS.Icon
    txtNumeroComando.Text = NumeroComando
    Me.Caption = Me.Caption & " " & Acao
    Me.lblNumeroComando.Caption = Me.lblNumeroComando.Caption & " " & Acao
    fgCenterMe Me

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler

    Select Case Button.Key
        Case gstrSalvar
            If Trim$(txtNumeroComandoAcao.Text) = "0" Or Trim$(txtNumeroComandoAcao.Text) = vbNullString Then
                MsgBox "Informe o número do comando de " & Acao
                Exit Sub
            End If
            NumeroComandoAcao = txtNumeroComandoAcao.Text
            Me.Hide

        Case gstrSair
            NumeroComandoAcao = 0
            Me.Hide
    End Select

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbCadastro_ButtonClick"
End Sub

Private Sub txtNumeroComandoACAO_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler

    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtNumeroComandoACAO_KeyPress"
End Sub
