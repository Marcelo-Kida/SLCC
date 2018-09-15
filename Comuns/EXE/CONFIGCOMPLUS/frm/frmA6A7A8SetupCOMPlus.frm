VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmA6A7A8SetupCOMPlus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A6A7A8SetupCOMPlus - Instalação e Configuração dos Componentes"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   ClipControls    =   0   'False
   Icon            =   "frmA6A7A8SetupCOMPlus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemover 
      Caption         =   "R&emover Componentes"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   4800
      Width           =   1875
   End
   Begin VB.Frame fraInterface 
      Caption         =   "Status da Configuração"
      Height          =   4695
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   12195
      Begin MSComctlLib.ListView lvwStatus 
         Height          =   3855
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Elemento"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ação"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Complemento"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ProgressBar prgStatus 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   4320
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Visible         =   0   'False
         Width           =   5715
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdInstalar 
      Caption         =   "&Instalar"
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   4800
      Width           =   855
   End
End
Attribute VB_Name = "frmA6A7A8SetupCOMPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Empresa        : Regerbanc
'Pacote         :
'Classe         : frmA6A7A8SetupCOMPlus
'Data Criação   : 10/10/2003
'Objetivo       : Instalar e configurar os componentes do SLCC
'
'Analista       : Adilson Gonçalves Damasceno
'
'Programador    : Eder Andrade
'Data           : 14/10/2003
'
'Teste          :
'Autor          :
'
'Data Alteração : 05/11/2003
'Autor          : Eder Andrade
'Objetivo       : Exibição e gravação dos logs das operações
'
'Data Alteração : 04/12/2003
'Autor          : Eder Andrade
'Objetivo       : Implementação da rotina que desregistra os componentes anteriores


Option Explicit

Private Sub cmdInstalar_Click()
Dim xml                                     As New MSXML2.DOMDocument40
Dim lstItem                                 As ListItem
    
On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass

    cmdRemover.Enabled = False
    cmdInstalar.Enabled = False
    cmdSair.Enabled = False
    
    lvwStatus.ListItems.Clear
    flconfigurarListView
    
    fgInstalar
        
    If Not xmlLog Is Nothing Then
        xmlLog.save App.Path & "\" & App.EXEName & "InstLog.xml"
    End If
    
    MsgBox "Componentes A6A7A8, Instalados e configurados com sucesso no COM+", vbOKOnly, App.EXEName
        
    cmdRemover.Enabled = True
    cmdInstalar.Enabled = True
    cmdSair.Enabled = True
    
    Call fgBarraStatus(1, 0, vbNullString)
    
    Screen.MousePointer = vbNormal
    
    lvwStatus.Refresh
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbNormal
    
    cmdRemover.Enabled = True
    cmdInstalar.Enabled = True
    cmdSair.Enabled = True
    
    Call fgBarraStatus(1, 0, vbNullString)
    
    If Not xmlLog Is Nothing Then
        xmlLog.save App.Path & "\" & App.EXEName & "Log.xml"
    End If
    
    MsgBox Err.Description

End Sub

Private Sub cmdRemover_Click()
Dim xml                                     As New MSXML2.DOMDocument40
Dim lstItem                                 As ListItem
    
On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    cmdRemover.Enabled = False
    cmdInstalar.Enabled = False
    cmdSair.Enabled = False
    
    lvwStatus.ListItems.Clear
    flconfigurarListView
    
    fgDesinstalar
        
    If Not xmlLog Is Nothing Then
        xmlLog.save App.Path & "\" & App.EXEName & "RemLog.xml"
    End If
    
    MsgBox "Componentes A6A7A8, removidos com sucesso no COM+", vbOKOnly, App.EXEName
    
    cmdRemover.Enabled = True
    cmdInstalar.Enabled = True
    cmdSair.Enabled = True
    
    Call fgBarraStatus(1, 0, vbNullString)
    
    Screen.MousePointer = vbNormal
    
    lvwStatus.Refresh
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbNormal
    
    cmdRemover.Enabled = True
    cmdInstalar.Enabled = True
    cmdSair.Enabled = True
    
    Call fgBarraStatus(1, 0, vbNullString)
    
    If Not xmlLog Is Nothing Then
        xmlLog.save App.Path & "\" & App.EXEName & "Log.xml"
    End If
    
    MsgBox Err.Description
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    
    flMensagemInicial
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub flMensagemInicial()

Dim objListItem                             As MSComctlLib.ListItem

    With lvwStatus
        With .ColumnHeaders
            .Clear
            .Add , , "Informação", 12000
        End With
        With .ListItems
            .Add , , "Para atualizar a versão do SLCC:"
            .Add , , "1 - Execute a opção 'Remover Componentes' antes de atualizar os arquivos da nova versão(Desregistrar componentes)."
            .Add , , "2 - Copiar os novos arquivos para a pasta do SLCC "
            .Add , , "3 - Executar a opção 'Instalar'"
        End With
    End With
    
    For Each objListItem In lvwStatus.ListItems
        objListItem.ForeColor = vbRed
    Next objListItem


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not cmdInstalar.Enabled Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmA6A7A8SetupCOMPlus = Nothing
End Sub

Private Sub flconfigurarListView()
    
    lvwStatus.ColumnHeaders.Clear
    lvwStatus.ColumnHeaders.Add 1, "Elemento", "Elemento", 4500
    lvwStatus.ColumnHeaders.Add 2, "Ação", "Ação", 3000
    lvwStatus.ColumnHeaders.Add 3, "Complemento", "Complemento", 4500
        
End Sub
