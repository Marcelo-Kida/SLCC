VERSION 5.00
Begin VB.Form frmTrocarSenha 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Trocar Senha"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form4"
   ScaleHeight     =   4080
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConfirmacao 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   1515
   End
   Begin VB.TextBox txtNovaSenha 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3180
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   315
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3180
      Width           =   915
   End
   Begin VB.TextBox txtSenhaAtual 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   1515
   End
   Begin VB.TextBox txtUsuario 
      Height          =   300
      Left            =   2280
      MaxLength       =   13
      TabIndex        =   0
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmação"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1140
      TabIndex        =   9
      Top             =   2580
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nova Senha"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha Atual"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1140
      TabIndex        =   7
      Top             =   1860
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1140
      TabIndex        =   6
      Top             =   1500
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      X1              =   480
      X2              =   4260
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      X1              =   480
      X2              =   4260
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   1260
      Picture         =   "frmTrocarSenha.frx":0000
      Top             =   420
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Height          =   3435
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   300
      Visible         =   0   'False
      Width           =   4275
   End
End
Attribute VB_Name = "frmTrocarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event SenhaTrocada(ByVal psSenhaAntiga As String, ByVal psNovaSenha As String)

Private Sub cmdOK_Click()

Dim objPerfil                               As SLCCSeguranca.clsPerfil

On Error GoTo ErrorHandler

    fgCursor True
    
    Set objPerfil = CreateObject("SLCCSeguranca.clsPerfil")
    objPerfil.TrocarSenha txtUsuario, txtSenhaAtual, txtNovaSenha, txtConfirmacao
    
    MsgBox "Troca de Senha Realizada com Sucesso.", vbInformation, Me.Caption
    RaiseEvent SenhaTrocada(txtSenhaAtual, txtNovaSenha)
        
    Set objPerfil = Nothing
    Unload Me
    
fgCursor
    Exit Sub
ErrorHandler:
fgCursor
    Set objPerfil = Nothing
    frmStartScreen.uctLogErros1.MostrarErros Err

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

Dim llRet                                   As Long
Dim llRgn                                   As Long

fgCursor True
    CenterMe Me
    
    llRgn = CreateRoundRectRgn(20, 40, 310, 275, 60, 60)
    llRet = SetWindowRgn(hwnd, llRgn, True)

    DoEvents
    
    txtUsuario = gstrUsuario
    
fgCursor
    Exit Sub
ErrorHandler:
fgCursor
    frmStartScreen.uctLogErros1.MostrarErros Err

End Sub

Private Sub txtConfirmacao_GotFocus()
    txtConfirmacao.SelStart = 0
    txtConfirmacao.SelLength = Len(Trim(txtConfirmacao))
End Sub

Private Sub txtNovaSenha_GotFocus()
    txtNovaSenha.SelStart = 0
    txtNovaSenha.SelLength = Len(Trim(txtNovaSenha))
End Sub

Private Sub txtSenhaAtual_GotFocus()
    txtSenhaAtual.SelStart = 0
    txtSenhaAtual.SelLength = Len(Trim(txtSenhaAtual))
End Sub

Private Sub txtUsuario_GotFocus()
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = Len(Trim(txtUsuario))
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
