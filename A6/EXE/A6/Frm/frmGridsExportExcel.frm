VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGridsExportExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione os Grids a serem exportados para o Excel"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   60
      TabIndex        =   8
      Top             =   3360
      Width           =   5025
      Begin VB.Frame Frame2 
         Height          =   1425
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   4755
         Begin VB.OptionButton optAdionarPastas 
            Caption         =   "Criar novas pastas para a exportação dos dados"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Value           =   -1  'True
            Width           =   3735
         End
         Begin VB.OptionButton optSobreporPastas 
            Caption         =   "Sobrepor os dados das pastas existentes na exportação"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   1050
            Width           =   4275
         End
         Begin VB.TextBox txtCaminhoPlanilha 
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   270
            Width           =   3165
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.OptionButton optPlanilhaExistente 
         Caption         =   "Exportar para uma planilha Excel existente"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   690
         Width           =   3285
      End
      Begin VB.OptionButton optPlanilhaNova 
         Caption         =   "Exportar para uma planilha Excel nova"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   330
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin MSComctlLib.ListView lvwGrids 
      Height          =   3225
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   5689
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlIcons"
      SmallIcons      =   "imlIcons"
      ColHdrIcons     =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Grid"
         Object.Width           =   8467
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Height          =   330
      Left            =   2790
      TabIndex        =   6
      Top             =   5910
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ButtonWidth     =   2011
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exportar"
            Key             =   "exportar"
            ImageKey        =   "exportar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r       "
            Key             =   "sair"
            ImageKey        =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":0000
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":031A
            Key             =   "exportar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":195E
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":1A70
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":1B6A
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGridsExportExcel.frx":1C64
            Key             =   "Leaf"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgPath 
      Left            =   570
      Top             =   5790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGridsExportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Objeto responsável pela exibição e seleção de grids de um determinado formulário
' a serem exportados para o Excel.

Option Explicit

Public objMyForm                            As Form

Private colListaGrids                       As New Collection

Private Sub flCarregarListaGrids()

Dim objControle                             As Object
Dim intContControles                        As Integer
Dim strDescricao                            As String

    On Error GoTo ErrorHandler
    
    For Each objControle In objMyForm.Controls
        
        If TypeOf objControle Is MSFlexGrid Or _
           TypeOf objControle Is MSComctlLib.ListView Or _
           TypeOf objControle Is vaSpread Then
           
            intContControles = intContControles + 1
            strDescricao = IIf(objControle.Tag = vbNullString, "Propriedade 'Tag' do Grid está em branco", objControle.Tag)
            
            colListaGrids.Add objControle
            lvwGrids.ListItems.Add , , intContControles & " - " & strDescricao
        
        End If
    
    Next
    
    If intContControles = 0 Then
        lvwGrids.ListItems.Add , , "Nenhum Grid a ser exportado para o Excel, neste formulário"
        tlbComandos.Buttons("exportar").Enabled = False
    End If
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - flCarregarListaGrids"

End Sub

Private Function flExportar() As Boolean

Dim objListItem                             As MSComctlLib.ListItem
Dim colGridsExport                          As New Collection
Dim blnSobrepor                             As Boolean

    On Error GoTo ErrorHandler
    
    For Each objListItem In lvwGrids.ListItems
        If objListItem.Checked Then
            colGridsExport.Add colListaGrids.Item(objListItem.Index)
        End If
    Next
    
    If colGridsExport.Count = 0 Then
        MsgBox "Nenhum Grid foi selecionado para a exportação.", vbExclamation, "Exportação para o Excel"
        flExportar = False
    ElseIf optPlanilhaExistente.Value And txtCaminhoPlanilha.Text = vbNullString Then
        MsgBox "Informe o caminho da planilha a ser utilizada para a exportação.", vbExclamation, "Exportação para o Excel"
        flExportar = False
    Else
        fgCursor True
        Call fgExportaExcel(colGridsExport, optPlanilhaNova.Value, txtCaminhoPlanilha.Text, optSobreporPastas.Value)
        fgCursor
        flExportar = True
    End If
    
    Exit Function

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - flExportar"

End Function

Private Sub cmdBuscar_Click()

    On Error GoTo ErrorHandler
    
    With dlgPath
        .DialogTitle = "Selecione a planilha Excel a ser utilizada"
        .InitDir = "C:\"
        .Filter = "*.xls"
        .ShowOpen
        txtCaminhoPlanilha.Text = .FileName
    End With
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_Load"

End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorHandler
    
    Me.Icon = mdiSBR.Icon
    Call flCarregarListaGrids
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGridsExportExcel = Nothing
End Sub

Private Sub optPlanilhaExistente_Click()
    txtCaminhoPlanilha.Enabled = True
    cmdBuscar.Enabled = True
    optAdionarPastas.Enabled = True
    optSobreporPastas.Enabled = True
    txtCaminhoPlanilha.SetFocus
End Sub

Private Sub optPlanilhaNova_Click()
    txtCaminhoPlanilha.Enabled = False
    cmdBuscar.Enabled = False
    optAdionarPastas.Enabled = False
    optSobreporPastas.Enabled = False
End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "sair"
            Set objMyForm = Nothing
            Unload Me
        Case "exportar"
            If flExportar Then
                Set objMyForm = Nothing
                Unload Me
            End If
    End Select
            
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick"

End Sub
