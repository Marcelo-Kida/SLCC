VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComposicaoNetOperacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Composição do Net de Operações"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwComposicaoNet 
      Height          =   3345
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtIdentPartCamaraContraparte 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   450
      Width           =   1485
   End
   Begin VB.TextBox txtIdentPartCamara 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   60
      Width           =   1485
   End
   Begin VB.TextBox txtContraparte 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   450
      Width           =   4845
   End
   Begin VB.TextBox txtVeiculoLegal 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   4845
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2430
      Top             =   3930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":031A
            Key             =   "Padrao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":076C
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":0A86
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":0DA0
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":10BA
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":150C
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComposicaoNetOperacoes.frx":195E
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Height          =   330
      Left            =   10290
      TabIndex        =   8
      Top             =   4245
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ButtonWidth     =   1191
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar Parametrização Padrão"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lblComposicaoNet 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Identificador Participante Câmara Contraparte"
      Height          =   195
      Index           =   3
      Left            =   6210
      TabIndex        =   6
      Top             =   510
      Width           =   3210
   End
   Begin VB.Label lblComposicaoNet 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Identificador Participante Câmara"
      Height          =   195
      Index           =   2
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label lblComposicaoNet 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Contraparte"
      Height          =   195
      Index           =   1
      Left            =   255
      TabIndex        =   2
      Top             =   510
      Width           =   825
   End
   Begin VB.Label lblComposicaoNet 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Veículo Legal"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmComposicaoNetOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição da composição do Net de Operações

Option Explicit

Private Const COL_OP_ID_TITULO              As Integer = 0
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 1
Private Const COL_OP_CV                     As Integer = 2
Private Const COL_OP_VALOR                  As Integer = 3
Private Const COL_OP_QUANTIDADE             As Integer = 4
Private Const COL_OP_CODIGO                 As Integer = 5

Private Const KEY_NU_SEQU_OPER_ATIV         As Integer = 1

Public strXmlOperacoes                      As String
Public strCondicaoNavegacaoXml              As String

Private xmlOperacoes                        As MSXML2.DOMDocument40

'Carregar composição do Net de Operações

Private Sub flCarregarLista()

Dim objDomNode                              As MSXML2.IXMLDOMNode
    
    On Error GoTo ErrorHandler
    
    Set xmlOperacoes = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlOperacoes.loadXML(strXmlOperacoes)
    
    For Each objDomNode In xmlOperacoes.selectNodes(strCondicaoNavegacaoXml)

        With lvwComposicaoNet.ListItems.Add(, "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            .Text = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
            .SubItems(COL_OP_CODIGO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            .SubItems(COL_OP_CV) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
            .SubItems(COL_OP_QUANTIDADE) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("QT_ATIV_MERC").Text)
            .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
            
            If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
            End If
            
        End With

    Next
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flCarregarLista", Me.Caption
    
End Sub

'Inicializar colunas da lista de operações

Private Sub flInicializarLvwOperacao()

    On Error GoTo ErrorHandler

    With Me.lvwComposicaoNet.ColumnHeaders
        .Clear
        .Add , , "ID Título", 1500
        .Add , , "Data Vencimento", 1800
        .Add , , "D/C", 1000
        .Add , , "Valor Operação", 2000, lvwColumnRight
        .Add , , "Quantidade", 2000, lvwColumnRight
        .Add , , "Código", 2200
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwOperacao", 0

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    DoEvents
    
    fgCursor True
    Call flInicializarLvwOperacao
    Call flCarregarLista
    fgCursor
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlOperacoes = Nothing
    Set frmComposicaoNetOperacoes = Nothing
End Sub

Private Sub lvwComposicaoNet_DblClick()

    On Error GoTo ErrorHandler

    If Not lvwComposicaoNet.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Split(lvwComposicaoNet.SelectedItem.Key, "|")(KEY_NU_SEQU_OPER_ATIV)
            .Show vbModal
        End With
    End If

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwComposicaoNet_DblClick", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "sair"
            Unload Me
    End Select
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwComposicaoNet_DblClick", Me.Caption

End Sub
