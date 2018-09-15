VERSION 5.00
Begin VB.Form frmLogon 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Logar"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form4"
   ScaleHeight     =   3360
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Logar"
      Default         =   -1  'True
      Height          =   315
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   315
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   915
   End
   Begin VB.TextBox txtSenha 
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1140
      TabIndex        =   5
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1140
      TabIndex        =   4
      Top             =   1500
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      X1              =   480
      X2              =   4260
      Y1              =   2220
      Y2              =   2220
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
      Picture         =   "frmLogon.frx":0000
      Top             =   420
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Height          =   2775
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   300
      Visible         =   0   'False
      Width           =   4275
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gbTrocarUsuario                     As Boolean
Private WithEvents objTrocarSenha           As frmTrocarSenha
Attribute objTrocarSenha.VB_VarHelpID = -1

Private Sub cmdLogar_Click()
On Error GoTo ErrorHandler
Dim objPerfil                               As PJPKSeguranca.clsPerfil

fgCursor True
    Set objPerfil = CreateObject("PJPKSeguranca.clsPerfil")
    
    If objPerfil.Logar(txtUsuario, txtSenha, gstrUsuarioRede, gstrEstacaoTrabalho, gbTrocarUsuario) Then
        gstrUsuario = txtUsuario
        
        frmStartScreen.Show
        Unload Me
    End If
    
    Set objPerfil = Nothing

fgCursor
    Exit Sub
ErrorHandler:
fgCursor
    Set objPerfil = Nothing
    uctLogErros1.MostrarErros Err
    
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
        
    llRgn = CreateRoundRectRgn(20, 40, 310, 230, 60, 60)
    llRet = SetWindowRgn(hwnd, llRgn, True)

    DoEvents

fgCursor
    Exit Sub
ErrorHandler:
fgCursor
    uctLogErros1.MostrarErros Err

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
    
    Set objTrocarSenha = Nothing
    Exit Sub
ErrorHandler:
    Cancel = True
    uctLogErros1.MostrarErros Err

End Sub

Private Sub objTrocarSenha_SenhaTrocada(ByVal psSenhaAntiga As String, ByVal psNovaSenha As String)
On Error GoTo ErrorHandler
Dim objSeguranca                            As PJPKSeguranca.clsPerfil
    
    Set objSeguranca = CreateObject("PJPKSeguranca.clsPerfil")
    objSeguranca.Logar gstrUsuario, psNovaSenha, gstrUsuarioRede, gstrEstacaoTrabalho
    Set objSeguranca = Nothing
    
    frmStartScreen.Show
    Unload Me
    
    Exit Sub
ErrorHandler:
    Set objSeguranca = Nothing
    uctLogErros1.MostrarErros Err
    
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(Trim(txtSenha))
End Sub

Private Sub txtUsuario_GotFocus()
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = Len(Trim(txtUsuario))
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub ShowLogon()
On Error GoTo ErrorHandler

    gbTrocarUsuario = False
    Me.Show

    Exit Sub
ErrorHandler:
    uctLogErros1.MostrarErros Err
    
End Sub

Public Sub ShowTrocarUsuario()
On Error GoTo ErrorHandler

    gbTrocarUsuario = True
    Me.Show

    Exit Sub
ErrorHandler:
    uctLogErros1.MostrarErros Err

End Sub

Private Sub uctLogErros1_ErroGerado(ErrNumber As Long, ErrDescription As String, Cancel As Boolean)
On Error GoTo ErrorHandler

Dim objSeguranca                            As PJPKSeguranca.clsPerfil

    Select Case ErrNumber
        Case ERR_USUARIONAOLOGADO
            Set objSeguranca = CreateObject("PJPKSeguranca.clsPerfil")
            If objSeguranca.Reconectar(gstrUsuario, _
                                       gstrUsuarioRede, _
                                       gstrEstacaoTrabalho) Then
                            
                ErrNumber = ErrNumber + 15
                ErrDescription = "Seu usuário foi reconectado ao sistema" & vbCr & _
                                 "Reinicie sua operação"
            Else
                ErrNumber = ErrNumber + 25
                ErrDescription = "Seu usuário foi desconectado" & vbCr & _
                                 "É necessário reiniciar o sistema"
            End If
            
            Set objSeguranca = Nothing
        
        Case ERR_SENHAEXPIRADA
            Cancel = True
            gstrUsuario = txtUsuario
            
            If MsgBox("Senha expirada." & vbNewLine & _
                      "Deseja realizar a troca imediatamente ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                            
                If objTrocarSenha Is Nothing Then
                    Set objTrocarSenha = New frmTrocarSenha
                End If
                objTrocarSenha.Show
            End If
    
    End Select

    Exit Sub
ErrorHandler:
    'Caso ocorra algum erro na reconecção,
    '   esse erro não é tratado, mostrando apenas o primeiro erro gerado
    
    Set objSeguranca = Nothing
    Set objTrocarSenha = Nothing
    Err.Clear
    
End Sub
