VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmListaOperacoes 
   Caption         =   "Lista de Operações"
   ClientHeight    =   6450
   ClientLeft      =   465
   ClientTop       =   2235
   ClientWidth     =   13320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   13320
   Begin MSComctlLib.ListView lstOperacoes 
      Height          =   2265
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   3995
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   6120
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   582
      ButtonWidth     =   1667
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechar "
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListaOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela Listagem de Operações
'

Option Explicit

Private Const strFuncionalidade             As String = "frmConciliacaoTitulos"

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    flMontaCampos

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    lstOperacoes.Width = Me.ScaleWidth - (lstOperacoes.Left * 2)
    lstOperacoes.Height = Me.ScaleHeight - lstOperacoes.Top - (lstOperacoes.Left) - tlbFiltro.Height
        
    On Error GoTo 0

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

'Formata as colunas da lista
Sub flMontaCampos()
    
    With lstOperacoes.ColumnHeaders
        .Clear
    
        .Add , , "Veículo Legal", 1600
        .Add , , "C/V", 1500
        .Add , , "ID", 2150
        .Add , , "Data Vencimento", 1440
        .Add , , "Quantidade", 1440, lvwColumnRight
        .Add , , "PU", 1440, lvwColumnRight
        .Add , , "Valor", 1440, lvwColumnRight
        .Add , , "Núm. Comando", 1440
        .Add , , "Data Liquidação", 1440
        .Add , , "Data Operação", 1440
        
    End With

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "Fechar" Then
        Unload Me
    End If

End Sub
