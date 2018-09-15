VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStartScreen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SLCC"
   ClientHeight    =   6135
   ClientLeft      =   2190
   ClientTop       =   3255
   ClientWidth     =   1725
   ClipControls    =   0   'False
   Icon            =   "frmStartScreen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin A6A7A8StartScreen.ctlErrorMessage uctLogErros 
      Left            =   600
      Top             =   4800
      _ExtentX        =   1058
      _ExtentY        =   979
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   5790
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Ambiente"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin A6A7A8StartScreen.ctlSysTray SLCCSystray 
      Left            =   1245
      Top             =   4920
      _ExtentX        =   688
      _ExtentY        =   688
      InTray          =   0   'False
      TrayIcon        =   "frmStartScreen.frx":0442
      TrayTip         =   "SLCC"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":257C
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":2896
            Key             =   "CC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":4228
            Key             =   "BUS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":4542
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":485C
            Key             =   "SBR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":4B76
            Key             =   "LQS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":4E90
            Key             =   "Auditoria"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartScreen.frx":51AA
            Key             =   "CC2"
         EndProperty
      EndProperty
   End
   Begin A6A7A8StartScreen.ctlBarraVertical SLCCMenu 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   10186
      Tabs            =   $"frmStartScreen.frx":5E84
      ImageListName   =   "ImageList"
   End
   Begin VB.Menu mnuSPB 
      Caption         =   "mnuSPB"
      Visible         =   0   'False
      Begin VB.Menu mnuMinimizar 
         Caption         =   "Minimizar"
      End
      Begin VB.Menu mnuRestaurar 
         Caption         =   "Restaurar"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "Fechar"
      End
   End
End
Attribute VB_Name = "frmStartScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Controlar a execução dos Sistemas do Projeto SLCC (A6, A7 e A8)

Option Explicit

Private luA6ProcessInfo                     As PROCESS_INFORMATION
Private luA6CCProcessInfo                   As PROCESS_INFORMATION
Private luA7ProcessInfo                     As PROCESS_INFORMATION
Private luA8ProcessInfo                     As PROCESS_INFORMATION

Private luA6AuditoriaProcessInfo            As PROCESS_INFORMATION
Private luA6CCAuditoriaProcessInfo          As PROCESS_INFORMATION
Private luA7AuditoriaProcessInfo            As PROCESS_INFORMATION
Private luA8AuditoriaProcessInfo            As PROCESS_INFORMATION

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 Then
        mnuMinimizar_Click
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

#If EnableSoap = 1 Then
    Dim objUsuario                          As MSSOAPLib30.SoapClient30
#Else
    Dim objUsuario                          As A6A7A8Miu.clsUsuario
#End If

On Error GoTo ErrorHandler
    
    fgTerminateProcess luA6ProcessInfo
    fgTerminateProcess luA6CCProcessInfo
    fgTerminateProcess luA7ProcessInfo
    fgTerminateProcess luA8ProcessInfo
    
    fgTerminateProcess luA6AuditoriaProcessInfo
    fgTerminateProcess luA6CCAuditoriaProcessInfo
    fgTerminateProcess luA7AuditoriaProcessInfo
    fgTerminateProcess luA8AuditoriaProcessInfo
    
    
    Set objUsuario = fgCriarObjetoMIU("A6A7A8Miu.clsUsuario")
    objUsuario.Logoff gstrUsuarioRede
    Set objUsuario = Nothing

    'Limpa os TLBs
    fgDesregistraComponentes '          <-- Inibido temporariamente para os testes com SOAP
    
    Set frmStartScreen = Nothing
    
    End
    
    Exit Sub
    
ErrorHandler:
    
    uctLogErros.MostrarErros Err, "frmStartScreen - Form_Unload"
    Set objUsuario = Nothing
    Set frmStartScreen = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub SLCCSystray_MouseDblClick(Button As Integer, Id As Long)

    If Button = vbLeftButton Then
        mnuRestaurar_Click
    End If
    
    PostMessage Me.hwnd, WM_USER, 0&, 0&                    ' Update form...

End Sub

Private Sub SLCCSystray_MouseDown(Button As Integer, Id As Long)
    
    If Button = vbRightButton Then
        ShowPopupMenu
    End If
    
    PostMessage Me.hwnd, WM_USER, 0&, 0&                    ' Update form...

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True
    
    StatusBar.Panels("Ambiente").Text = gstrAmbiente
    
    With SLCCMenu
        .ActiveTab = SLCCMenu.Tabs("Sistemas").Index
    End With
    
    With Me
        Move Screen.Width - 2000, Screen.Height - 6850
        Show
        Refresh
    End With
     
    'Componente resposavel pelo controle do systray
    '   adiciona um icone para o SPB no systray do windows
    With SLCCSystray
        .InTray = True
        Set .TrayIcon = Icon
        .TrayTip = "SLCC - " & gstrAmbiente
    End With
    
    'Força o form sempre visivel, estando sempre a frente das outras janelas
    AlwaysOnTop Me, True

    fgCursor False
    
    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    uctLogErros.MostrarErros Err, "frmStartScreen - Form_Load"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ShowPopupMenu
    End If
    
    Exit Sub
ErrorHandler:
    
    uctLogErros.MostrarErros Err, "frmStartScreen - Form_MouseDown"
    
End Sub

Private Sub mnuFechar_Click()
    
    PostMessage Me.hwnd, WM_CLOSE, 0&, 0&           ' Close window(Unload Me will GPF)

End Sub

Private Sub mnuMinimizar_Click()

On Error GoTo ErrorHandler
    
    Me.Visible = Not Me.Visible

    Exit Sub
ErrorHandler:
    uctLogErros.MostrarErros Err, "frmStartScreen - mnuMinimizar_Click"

End Sub

Private Sub mnuRestaurar_Click()

On Error GoTo ErrorHandler

    fgCursor True
    Me.Visible = Not Me.Visible
    fgCursor False
    
    Exit Sub
    
ErrorHandler:
    
    fgCursor False
    
    uctLogErros.MostrarErros Err, "frmStartScreen - mnuRestaurar_Click"

End Sub

'Mostra o menu Popup

Private Sub ShowPopupMenu()

On Error GoTo ErrorHandler

    Select Case Me.Visible
    
        Case True
            mnuMinimizar.Enabled = True
            mnuRestaurar.Enabled = False

        Case False
            mnuMinimizar.Enabled = False
            mnuRestaurar.Enabled = True

    End Select
    
    Me.PopupMenu mnuSPB

    Exit Sub
    
ErrorHandler:
    
    uctLogErros.MostrarErros Err, "frmStartScreen - ShowPopupMenu"

End Sub

Private Sub SLCCMenu_Click(ByVal TabIndex As Integer, ByVal TabKey As String, ByVal ButtonIndex As Integer, ByVal ButtonKey As String)

Dim strParametros                           As String

Const OWNER_A6                              As String = "OWNERA6"
Const OWNER_A6_CC                           As String = "OWNERA6COLI"
Const OWNER_A7                              As String = "OWNERA7"
Const OWNER_A8                              As String = "OWNERA8"
Dim lngRetorno                              As Long
Dim strTitle                                As String

On Error GoTo ErrorHandler
        
    fgCursor True
    'Parâmetros para A6, A7 e A8
    strParametros = gstrAmbiente & ";" & _
                    gstrSource & ";" & _
                    gstrUsuario & ";" & _
                    "ON;" & _
                    Abs(gblnRegistraTLB) & ";" & _
                    gstrURLWebService & ";" & _
                    glngTimeOut & ";" & _
                    gstrPrint
        
    Select Case ButtonKey
        
        
        Case "A6CC"
        
            strTitle = "A6 - Caixa Coligadas - " & gstrAmbiente
            strParametros = strParametros & ";" & gstrHelpFileA6
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathA6CC, strParametros, PER_SISTEMA_A6CC, luA6CCProcessInfo
            End If
        
        Case "A6"
        
            strTitle = "A6 - Sub-Reserva - " & gstrAmbiente
            strParametros = strParametros & ";" & gstrHelpFileA6
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathA6, strParametros, PER_SISTEMA_A6, luA6ProcessInfo
            End If
            
        Case "A7"
        
            strTitle = "A7 - BUS de Interface - " & gstrAmbiente
            strParametros = strParametros & ";" & gstrHelpFileA7
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathA7, strParametros, PER_SISTEMA_A7, luA7ProcessInfo
            End If
            
        Case "A8"
            
            strTitle = "A8 - Sistema de Liquidação e Controle das Câmaras - " & gstrAmbiente
            strParametros = strParametros & ";" & gstrHelpFileA8
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathA8, strParametros, PER_SISTEMA_A8, luA8ProcessInfo
            End If
            
        Case "AuditoriaA6CC"
            
            strTitle = "SLCC Trilha de Auditoria"
            
            strParametros = gstrAmbiente & ";" & _
                            gstrUsuario & ";" & _
                            OWNER_A6_CC & ";" & _
                            gstrURLWebService & ";" & _
                            glngTimeOut
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathAuditoria, strParametros, PER_AUDITORIA_A6CC, luA6CCAuditoriaProcessInfo
            End If
            
        Case "AuditoriaA6"
            
            strTitle = "SLCC Trilha de Auditoria"
            
            strParametros = gstrAmbiente & ";" & _
                            gstrUsuario & ";" & _
                            OWNER_A6 & ";" & _
                            gstrURLWebService & ";" & _
                            glngTimeOut
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathAuditoria, strParametros, PER_AUDITORIA_A6, luA6AuditoriaProcessInfo
            End If
            
        Case "AuditoriaA7"
            
            strTitle = "SLCC Trilha de Auditoria"
            
            strParametros = gstrAmbiente & ";" & _
                            gstrUsuario & ";" & _
                            OWNER_A7 & ";" & _
                            gstrURLWebService & ";" & _
                            glngTimeOut
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathAuditoria, strParametros, PER_AUDITORIA_A7, luA7AuditoriaProcessInfo
            End If
            
        Case "AuditoriaA8"
            
            strTitle = "SLCC Trilha de Auditoria"
            
            strParametros = gstrAmbiente & ";" & _
                            gstrUsuario & ";" & _
                            OWNER_A8 & ";" & _
                            gstrURLWebService & ";" & _
                            glngTimeOut
            
            If fgFindWindow(strTitle) > 0 Then
                fgSetFocus strTitle
            Else
                fgShowApplication gstrPathAuditoria, strParametros, PER_AUDITORIA_A8, luA8AuditoriaProcessInfo
            End If
        
    End Select

    AlwaysOnTop Me, True
    
fgCursor
    
    Exit Sub

ErrorHandler:
fgCursor
    
    uctLogErros.MostrarErros Err, "frmStartScreen - SLCCMenu_Click"

End Sub

Private Sub SLCCMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ShowPopupMenu
    End If
    
    Exit Sub
    
ErrorHandler:
    uctLogErros.MostrarErros Err, "frmStartScreen - SPBSystray_MouseDown"

End Sub

Private Sub uctLogErros_ErroGerado(ErrNumber As Long, ErrDescription As String, Cancel As Boolean)

On Error GoTo ErrorHandler

#If EnableSoap = 1 Then
    Dim objUsuario                          As MSSOAPLib30.SoapClient30
#Else
    Dim objUsuario                          As A6A7A8Miu.clsUsuario
#End If

    Select Case ErrNumber
        Case ERR_USUARIONAOLOGADO
            Set objUsuario = fgCriarObjetoMIU("A6A7A8Miu.clsUsuario")
            
            If objUsuario.Reconectar(gstrUsuario, _
                                       gstrUsuarioRede, _
                                       gstrEstacaoTrabalho) Then
                            
                ErrNumber = ErrNumber + 15
                ErrDescription = "Seu usuário foi reconectado ao sistema." & vbCr & _
                                 "Reinicie sua operação."
            Else
                ErrNumber = ErrNumber + 25
                ErrDescription = "Seu usuário foi desconectado." & vbCr & _
                                 "É necessário reiniciar o sistema."
            End If
            
            Set objUsuario = Nothing
    
        Case ERR_SEMACESSO
            ErrNumber = ErrNumber
            ErrDescription = "Seu usuário não possui acesso ao diretório." & vbCr & _
                             "Acionar Gerência SLCC."
        
    End Select

    Exit Sub
ErrorHandler:
    'Caso ocorra algum erro na reconecção,
    '   esse erro não é tratado, mostrando apenas o primeiro erro gerado
    
    Set objUsuario = Nothing
    
    Err.Clear
    
End Sub
