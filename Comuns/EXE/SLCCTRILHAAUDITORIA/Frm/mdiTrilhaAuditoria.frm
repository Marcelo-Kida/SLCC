VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiTrilhaAuditoria 
   BackColor       =   &H8000000C&
   Caption         =   "SLCC Trilha de Auditoria"
   ClientHeight    =   5730
   ClientLeft      =   2325
   ClientTop       =   1830
   ClientWidth     =   10245
   Icon            =   "mdiTrilhaAuditoria.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbTrilhaAuditoria 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Ambiente:"
            TextSave        =   "Ambiente:"
         EndProperty
      EndProperty
   End
   Begin A6A7A8TrilhaAuditoria.ctlErrorMessage uctLogErros 
      Left            =   1140
      Top             =   2625
      _ExtentX        =   1085
      _ExtentY        =   1032
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuTrilha 
      Caption         =   "Trilha "
      Begin VB.Menu mnuLogTabelas 
         Caption         =   "Log Tabelas"
      End
   End
End
Attribute VB_Name = "mdiTrilhaAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    
Dim strSistema                              As String

    If gstrOwnerSLCC = "OWNERA6" Then
       strSistema = "A6 SBR"
    ElseIf gstrOwnerSLCC = "OWNERA7" Then
       strSistema = "A7 - Bus de Interface"
    ElseIf gstrOwnerSLCC = "OWNERA8" Then
       strSistema = "A8 LQS"
    ElseIf gstrOwnerSLCC = "OWNERA6COLI" Then
       strSistema = "A6 Coligadas"
       
    End If
    
    stbTrilhaAuditoria.Panels(1).Text = "Sistema: " & strSistema
    stbTrilhaAuditoria.Panels(2).Text = "Ambiente: " & gstrAmbiente
    
    
End Sub

Private Sub mnuLogTabelas_Click()

On Error GoTo ErrHandler
       
    If gstrOwnerSLCC = "OWNERA6" Then
       frmLogTabela.Owner = enumSLCCOwner.OwnerA6
    ElseIf gstrOwnerSLCC = "OWNERA7" Then
       frmLogTabela.Owner = enumSLCCOwner.OwnerA7
    ElseIf gstrOwnerSLCC = "OWNERA8" Then
        frmLogTabela.Owner = enumSLCCOwner.OwnerA8
    ElseIf gstrOwnerSLCC = "OWNERA6COLI" Then
       frmLogTabela.Owner = enumSLCCOwner.OwnerA6Coli
    End If
    
    frmLogTabela.Show
    
    Exit Sub
    
ErrHandler:

    mdiTrilhaAuditoria.uctLogErros.MostrarErros Err, "mdiTrilhaAuditoria - mnuLogTabelas_Click"

End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

