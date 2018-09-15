VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaVeiculoLegal 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   14145
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCadastro 
      Caption         =   "Cadastro"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   9615
      Begin VB.TextBox txtDataUltimaAtualizacao 
         Height          =   315
         Left            =   9000
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   2550
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.ComboBox cboTipoTiularBMA 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ListBox lstTipoBCKAux 
         Height          =   255
         Left            =   6480
         TabIndex        =   32
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox txtTipoBCK 
         BackColor       =   &H80000011&
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox cboSistema 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   315
         Left            =   6480
         MaxLength       =   18
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtNomeReduzido 
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   15
         Top             =   2550
         Width           =   2415
      End
      Begin VB.TextBox txtNome 
         Height          =   315
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtVeiculoLegal 
         Height          =   315
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin NumBox.Number numIdCETIP 
         Height          =   315
         Left            =   6480
         TabIndex        =   19
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelecao     =   0   'False
         AceitaNegativo  =   0   'False
      End
      Begin NumBox.Number numPadraoSELIC 
         Height          =   315
         Left            =   6480
         TabIndex        =   21
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelecao     =   0   'False
         AceitaNegativo  =   0   'False
      End
      Begin NumBox.Number numCodigoBMA 
         Height          =   315
         Left            =   6480
         TabIndex        =   25
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelecao     =   0   'False
         AceitaNegativo  =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
         Height          =   315
         Left            =   6480
         TabIndex        =   27
         Top             =   2160
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         Format          =   56098817
         CurrentDate     =   37622
         MaxDate         =   73050
         MinDate         =   37622
      End
      Begin MSComCtl2.DTPicker dtpDataFimVigencia 
         Height          =   315
         Left            =   6480
         TabIndex        =   29
         Top             =   2550
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   56098817
         CurrentDate     =   37622
         MaxDate         =   73050
         MinDate         =   37622
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fim &Vigência"
         Height          =   195
         Index           =   13
         Left            =   4560
         TabIndex        =   28
         Top             =   2550
         Width           =   1290
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Início Vigência"
         Height          =   195
         Index           =   12
         Left            =   4560
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Titular &BMA"
         Height          =   195
         Index           =   11
         Left            =   4560
         TabIndex        =   24
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Tit&ular BMA"
         Height          =   195
         Index           =   10
         Left            =   4560
         TabIndex        =   22
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Padrã&o SELIC"
         Height          =   195
         Index           =   9
         Left            =   4560
         TabIndex        =   20
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Identificador CETIP"
         Height          =   195
         Index           =   8
         Left            =   4560
         TabIndex        =   18
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CNPJ"
         Height          =   195
         Index           =   7
         Left            =   4560
         TabIndex        =   16
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome &Reduzido"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   2550
         Width           =   1140
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo Back-Office"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo Veículo Legal"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Empresa"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sistema"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Veículo Legal"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   990
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3840
      TabIndex        =   31
      Top             =   6465
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwVeiculoLegal 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5318
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
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   30
      Top             =   6480
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   582
      ButtonWidth     =   3043
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro          "
            Key             =   "DefinirFiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   10560
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   9840
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":19F2
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":1D0C
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":2026
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":2340
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":265A
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":2AAC
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaVeiculoLegal.frx":2EFE
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'' Exibe todos os dados do cadastro de Veículo Legal

Option Explicit

Private intRefresh                          As Integer
Private strOperacao                         As String

Private xmlLer                              As MSXML2.DOMDocument40
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLerTodosTipoBCK                  As MSXML2.DOMDocument40
Private blnEditMode                         As Boolean
Private strCNPJAtual                        As String

Private Const strFuncionalidade             As String = "frmConsultaVeiculoLegal"

Private Const COL_VEICULO_LEGAL             As Integer = 0
Private Const COL_SIGLA_SISTEMA             As Integer = 1
Private Const COL_EMPRESA                   As Integer = 2
Private Const COL_GRUPO_VEICULO_LEGAL       As Integer = 3
Private Const COL_TIPO_BACKOFFICE           As Integer = 4
Private Const COL_NOME                      As Integer = 5
Private Const COL_NOME_REDUZIDO             As Integer = 6
Private Const COL_CNPJ                      As Integer = 7
Private Const COL_INDENTIFICADOR_CETIP      As Integer = 8
Private Const COL_CONTA_PADRAO_SELIC        As Integer = 9
Private Const COL_TIPO_TITULAR_BMA          As Integer = 10
Private Const COL_CODIGO_TITULAR_BMA        As Integer = 11
Private Const COL_DATA_INICIO_VIG           As Integer = 12
Private Const COL_DATA_FINAL_VIG            As Integer = 13

Dim WithEvents objFiltro                    As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Public blnConsulta                          As Boolean

Private lngIndexClassifList                 As Long

'' Formatas os títulos das colunas do grid
Private Sub flPreencherHeadersLvw()

On Error GoTo ErrorHandler

    With lvwVeiculoLegal.ColumnHeaders
        .Clear
        .Add 1, , "Veiculo Legal", 900
        .Add 2, , "Sigla do Sistema", 585
        .Add 3, , "Empresa", 2759
        .Add 4, , "Grupo Veículo Legal", 2039
        .Add 5, , "Tipo BackOffice", 1260
        .Add 6, , "Nome", 3075
        .Add 7, , "Nome Reduzido", 2489
        .Add 8, , "CNPJ do Veículo Legal", 1830
        .Add 9, , "Identificador Participante CETIP", 2459
        .Add 10, , "Conta Própria Custódia SELIC", 2489
        .Add 11, , "Tipo Titular BMA", 1725
        .Add 12, , "Código Titular BMA", 1725
        .Add 13, , "Data Início Vigência", 400
        .Add 14, , "Data Fim Vigência", 400
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPreencherHeadersLvw", 0

End Sub

'' Carrega todos os veículos legais na lista
Private Sub flCarregarLista(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objVeiculoLegal     As MSSOAPLib30.SoapClient30
#Else
    Dim objVeiculoLegal     As A8MIU.clsVeiculoLegal
#End If

Dim xmlRetorno              As MSXML2.DOMDocument40
Dim strRetorno              As String
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim objListItem             As ListItem
Dim dtTmp                   As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    fgCursor True
    fgLockWindow Me.hwnd
    
    Set objVeiculoLegal = fgCriarObjetoMIU("A8MIU.clsVeiculoLegal")
    
    strRetorno = objVeiculoLegal.ObterDetalheVeiculoLegal(pstrFiltro, _
                                                          vntCodErro, _
                                                          vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lvwVeiculoLegal.ListItems.Clear
    
    If strRetorno = vbNullString Then
        fgCursor
        fgLockWindow
        Exit Sub
    End If
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    xmlRetorno.loadXML strRetorno

    For Each objDomNode In xmlRetorno.documentElement.childNodes
    
        Set objListItem = lvwVeiculoLegal.ListItems.Add(, "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & objDomNode.selectSingleNode("SG_SIST").Text & "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
        objListItem.Text = fgSelectSingleNode(objDomNode, "CO_VEIC_LEGA").Text
        objListItem.SubItems(COL_SIGLA_SISTEMA) = fgSelectSingleNode(objDomNode, "SG_SIST").Text
        objListItem.SubItems(COL_EMPRESA) = fgSelectSingleNode(objDomNode, "CO_EMPR").Text & " - " & fgSelectSingleNode(objDomNode, "NO_EMPR").Text
        objListItem.SubItems(COL_GRUPO_VEICULO_LEGAL) = fgSelectSingleNode(objDomNode, "CO_GRUP_VEIC_LEGA").Text & " - " & fgSelectSingleNode(objDomNode, "NO_GRUP_VEIC_LEGA").Text
        objListItem.SubItems(COL_TIPO_BACKOFFICE) = fgSelectSingleNode(objDomNode, "TP_BKOF").Text & " - " & fgSelectSingleNode(objDomNode, "DE_BKOF").Text
        objListItem.SubItems(COL_NOME) = fgSelectSingleNode(objDomNode, "NO_VEIC_LEGA").Text
        objListItem.SubItems(COL_NOME_REDUZIDO) = fgSelectSingleNode(objDomNode, "NO_REDU_VEIC_LEGA").Text
        objListItem.SubItems(COL_CNPJ) = fgFormataCnpj(fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text)
        objListItem.SubItems(COL_INDENTIFICADOR_CETIP) = fgSelectSingleNode(objDomNode, "ID_PART_CAMR_CETIP").Text
        objListItem.SubItems(COL_CONTA_PADRAO_SELIC) = fgSelectSingleNode(objDomNode, "CO_CNTA_CUTD_PADR_SELIC").Text
        objListItem.SubItems(COL_TIPO_TITULAR_BMA) = fgSelectSingleNode(objDomNode, "TP_TITL_BMA").Text
        objListItem.SubItems(COL_CODIGO_TITULAR_BMA) = fgSelectSingleNode(objDomNode, "CO_TITL_BMA").Text
        objListItem.SubItems(COL_DATA_INICIO_VIG) = fgDtXML_To_Interface(fgSelectSingleNode(objDomNode, "DT_INIC_VIGE").Text)
        dtTmp = fgSelectSingleNode(objDomNode, "DT_FIM_VIGE").Text
        
        If dtTmp <> "00:00:00" And dtTmp <> "" Then
            objListItem.SubItems(COL_DATA_FINAL_VIG) = fgDtXML_To_Interface(dtTmp)
        Else
            objListItem.SubItems(COL_DATA_FINAL_VIG) = ""
        End If
        
    Next objDomNode

    Call fgClassificarListview(Me.lvwVeiculoLegal, lngIndexClassifList, True)
    
    Set objVeiculoLegal = Nothing
    Set xmlRetorno = Nothing

    fgLockWindow 0
    fgCursor

Exit Sub
ErrorHandler:
    fgLockWindow 0
    fgCursor
    Set objVeiculoLegal = Nothing
    Set xmlRetorno = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0
End Sub

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    If cboEmpresa.ListIndex <> -1 Then
        
        fgCursor True
        flCarregaComboSistema
        fgCursor False
        
    End If

Exit Sub
ErrorHandler:
   fgCursor False
    
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboEmpresa_Click"
End Sub

Private Sub cboGrupo_Click()

On Error Resume Next

    txtTipoBCK.Text = lstTipoBCKAux.List(cboGrupo.ListIndex) & " - " & _
    xmlLerTodosTipoBCK.selectSingleNode("//Repeat_TipoBackOffice/Grupo_TipoBackOffice/DE_BKOF[../TP_BKOF='" & lstTipoBCKAux.List(cboGrupo.ListIndex) & "']").Text
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("refresh"))
        
        Call fgCursor(False)
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyDown"

End Sub

Private Sub Form_Load()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlDomSistema As MSXML2.DOMDocument40
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant
    
On Error GoTo ErrorHandler
    
    lstTipoBCKAux.Visible = False
    lstTipoBCKAux.Enabled = False
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao

    fgCursor True
    fgCenterMe Me
    
    Set Me.Icon = mdiLQS.Icon
    If Me.blnConsulta Then
        Me.Caption = "Consulta Veiculo Legal"
        flOcultarControleCadastro
    Else
        Me.Caption = "Cadastro Veiculo Legal"
    End If
    Me.Show
    DoEvents
        
    flInicializar
    flPreencherHeadersLvw
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmFiltro", "Form_Load")
    End If

    If Not Me.blnConsulta Then
        flCarregarTipoTitularBMA
        fgCarregarCombos cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR", False
        flCarregaComboVeiculoLegal
    End If
        
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaVeiculoLegal
    Load objFiltro
    objFiltro.fgCarregarPesquisaAnterior

    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    Dim lngAlturaLista As Long
    
    If Not Me.blnConsulta Then
        tlbCadastro.Top = Me.ScaleHeight - tlbCadastro.Height
        fraCadastro.Top = tlbCadastro.Top - fraCadastro.Height - 60
        fraCadastro.Width = Me.ScaleWidth - fraCadastro.Left
        lngAlturaLista = fraCadastro.Top
    Else
        lngAlturaLista = Me.ScaleHeight - tlbFiltro.Height
    End If
    
    With lvwVeiculoLegal
        .Top = 0
        .Left = 0
        .Width = Me.Width - 100
        '.Height = fraCadastro.Top - 120
        .Height = lngAlturaLista - 120
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub lvwVeiculoLegal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwVeiculoLegal, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

    Exit Sub

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwVeiculoLegal_ColumnClick"
    
End Sub

Private Sub lvwVeiculoLegal_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error Resume Next
 
    blnEditMode = True
    txtVeiculoLegal.Locked = True
    cboEmpresa.Locked = True
    cboSistema.Locked = True
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    
    With Item
        txtVeiculoLegal.Text = .Text 'COL_VEICULO_LEGAL
        cboEmpresa.ListIndex = flRetornaIndiceCombo(cboEmpresa, .ListSubItems(COL_EMPRESA))
        cboSistema.ListIndex = flRetornaIndiceCombo(cboSistema, .ListSubItems(COL_SIGLA_SISTEMA), True)
        cboGrupo.ListIndex = flRetornaIndiceCombo(cboGrupo, .ListSubItems(COL_GRUPO_VEICULO_LEGAL))
        'txtTipoBCK.Text = COL_TIPO_BACKOFFICE
        txtNome.Text = .ListSubItems(COL_NOME)
        txtNomeReduzido.Text = .ListSubItems(COL_NOME_REDUZIDO)
        txtCNPJ.Text = .ListSubItems(COL_CNPJ)
        strCNPJAtual = txtCNPJ
        numIdCETIP.Valor = .ListSubItems(COL_INDENTIFICADOR_CETIP)
        numPadraoSELIC = .ListSubItems(COL_CONTA_PADRAO_SELIC)
        cboTipoTiularBMA.ListIndex = flRetornaIndiceCombo(cboTipoTiularBMA, .ListSubItems(COL_TIPO_TITULAR_BMA), True)
        numCodigoBMA.Valor = .ListSubItems(COL_CODIGO_TITULAR_BMA)
        dtpDataInicioVigencia.MinDate = CDate(.ListSubItems(COL_DATA_INICIO_VIG))
        dtpDataInicioVigencia.value = CDate(.ListSubItems(COL_DATA_INICIO_VIG))
        If Trim(.ListSubItems(COL_DATA_FINAL_VIG)) <> "" Then
            dtpDataFimVigencia.MinDate = CDate(.ListSubItems(COL_DATA_FINAL_VIG))
            dtpDataFimVigencia.value = CDate(.ListSubItems(COL_DATA_FINAL_VIG))
        Else
            dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.value
            dtpDataFimVigencia.value = Empty
        End If
        txtDataUltimaAtualizacao.Text = Split(.Key, "|")(2)
    End With
        
End Sub

Private Sub lvwVeiculoLegal_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
    
        Call fgCursor(True)
        
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("refresh"))
        
        Call fgCursor(False)
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwVeiculoLegal_KeyDown"

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

On Error GoTo ErrorHandler

    flCarregarLista xmlDocFiltros

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - objFiltro_AplicarFiltro"
End Sub

Private Sub numCodigoBMA_Change()
    If Len(numCodigoBMA.Valor) > 18 Then
        numCodigoBMA.Valor = Val(Mid(numCodigoBMA.Valor, 1, 18))
    End If
End Sub

Private Sub numIdCETIP_Change()
    If Len(numIdCETIP.Valor) > 8 Then
        numIdCETIP.Valor = Val(Mid(numIdCETIP.Valor, 1, 8))
    End If
End Sub

Private Sub numPadraoSELIC_Change()
    If Len(numPadraoSELIC.Valor) > 9 Then
        numPadraoSELIC.Valor = Val(Mid(numPadraoSELIC.Valor, 1, 9))
    End If
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strSelecaoFiltro                        As String
Dim strResultadoConfirmacao                 As String

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
            
        Case "DefinirFiltro"
            objFiltro.Show vbModal
            
        Case "refresh"
            objFiltro.fgCarregarPesquisaAnterior
            
    End Select
    fgCursor
    
Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaVeiculoLegal - tlbFiltro_ButtonClick", Me.Caption

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case gstrSalvar
            strOperacao = IIf(blnEditMode, gstrOperAlterar, gstrOperIncluir)
            Call flSalvar
        Case gstrOperExcluir
        
            If Not blnEditMode Then
                MsgBox "Selecione um item a ser excluído.", vbOKCancel + vbInformation, Me.Caption
                Exit Sub
            End If
            
            If MsgBox("Confirma a exclusão do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
               strOperacao = gstrOperExcluir
               Call flSalvar
            End If
            
        Case gstrSair
            fgCursor False
            Unload Me
            Exit Sub
    End Select
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmTipoJustificativaConciliacao - tlbCadastro_ButtonClick", Me.Caption
    Call flCarregarLista("")
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    End If

End Sub

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()
        
On Error GoTo ErrorHandler

    Dim ctrl As Control
    
    blnEditMode = False
    txtVeiculoLegal.Locked = False
    cboEmpresa.Locked = False
    cboSistema.Locked = False
    strCNPJAtual = ""
    
    For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is Number Then
            ctrl.Valor = 0
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
        End If
    Next
        
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(Data)
    dtpDataInicioVigencia.value = fgDataHoraServidor(Data)
    
    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = Null
    
    dtpDataInicioVigencia.Enabled = True
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    
Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

'' É acionado através no botão 'Salvar' da barra de ferramentas.
''
'' Tem como função, encaminhar a solicitação (Atualização dos dados na tabela) à
'' camada controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strRetorno              As String
Dim strPropriedades         As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
    
    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
    
    Call fgCursor(True)
    
    Call flInterfaceToXml

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    If objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro) Then
        'If strOperacao <> gstrOperExcluir Then
        '    xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        '    xmlLer.loadXML objMIU.Executar(xmlLer.xml)
        '   strOperacao = gstrOperAlterar
        'Else
        '    flLimpaCampos
        'End If
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objMIU = Nothing
        
        Call flLimpaCampos
        Call flCarregarLista("")
        
        MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    End If
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'Valida o preenchimento dos campos
Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If Trim(txtVeiculoLegal.Text) = "" Then
        flValidarCampos = "Digite o código do Veículo Legal."
        txtVeiculoLegal.SetFocus
        Exit Function
    End If
    
    If cboEmpresa.ListIndex = -1 Then
        flValidarCampos = "Informe a Empresa."
        cboEmpresa.SetFocus
        Exit Function
    End If
    
    If cboSistema.ListIndex = -1 Then
        flValidarCampos = "Informe a Sigla do Sistema."
        cboSistema.SetFocus
        Exit Function
    End If
        
    If cboGrupo.ListIndex = -1 Then
        flValidarCampos = "Informe o Grupo do Veículo Legal."
        cboGrupo.SetFocus
        Exit Function
    End If
    
    If Trim(txtNome.Text) = "" Then
        flValidarCampos = "Informe o Nome do Veículo Legal."
        txtNome.SetFocus
        Exit Function
    End If
    
    If Trim(txtNomeReduzido.Text) = "" Then
        flValidarCampos = "Informe o Nome Reduzido do Veículo Legal."
        txtNomeReduzido.SetFocus
        Exit Function
    End If
    
    If Not IsNull(dtpDataFimVigencia.value) Then
        If dtpDataFimVigencia.value < dtpDataInicioVigencia.value Then
            flValidarCampos = "Data final da vigência anterior à data inicial."
            dtpDataFimVigencia.SetFocus
            Exit Function
        End If
    End If
    
    If strCNPJAtual <> txtCNPJ.Text Then
        If Not fgValidaCNPJ_CPF(txtCNPJ.Text) Then
            flValidarCampos = "CNPJ inválido."
            txtCNPJ.SetFocus
            Exit Function
        End If
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Preenche o conteúdo do XML com o conteúdo dos campos apresentados em tela
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = strOperacao
        .selectSingleNode("CO_VEIC_LEGA").Text = fgLimpaCaracterInvalido(Trim(txtVeiculoLegal.Text))
        .selectSingleNode("SG_SIST").Text = fgObterCodigoCombo(cboSistema.Text)
        .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa.Text)
        .selectSingleNode("CO_GRUP_VEIC_LEGA").Text = fgObterCodigoCombo(cboGrupo.Text)
        .selectSingleNode("TP_BKOF").Text = fgObterCodigoCombo(txtTipoBCK.Text)
        .selectSingleNode("NO_VEIC_LEGA").Text = fgLimpaCaracterInvalido(Trim(txtNome.Text))
        .selectSingleNode("NO_REDU_VEIC_LEGA").Text = fgLimpaCaracterInvalido(Trim(txtNomeReduzido.Text))
        .selectSingleNode("CO_CNPJ_VEIC_LEGA").Text = fgLimpaCaracteresCNPJ(txtCNPJ.Text)
        .selectSingleNode("DT_INIC_VIGE").Text = fgDate_To_DtXML(dtpDataInicioVigencia.value)
        If Not IsNull(dtpDataFimVigencia.value) Then
            If fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text) <> dtpDataFimVigencia.value Then
                If MsgBox("Deseja desativar o registro a partir da data: " & dtpDataFimVigencia.value, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                    .selectSingleNode("DT_FIM_VIGE").Text = fgDate_To_DtXML(dtpDataFimVigencia.value)
                Else
                    .selectSingleNode("DT_FIM_VIGE").Text = ""
                End If
            End If
        Else
            .selectSingleNode("DT_FIM_VIGE").Text = ""
        End If
        .selectSingleNode("ID_PART_CAMR_CETIP").Text = numIdCETIP.Valor
        .selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text = numPadraoSELIC.Valor
        .selectSingleNode("TP_TITL_BMA").Text = fgObterCodigoCombo(cboTipoTiularBMA.Text)
        .selectSingleNode("CO_TITL_BMA").Text = numCodigoBMA.Valor
        .selectSingleNode("DH_ULTI_ATLZ").Text = txtDataUltimaAtualizacao.Text
    End With
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function

'' Carrega as propriedades necessárias a interface frmCadastroRegra, através da
'' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If

    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_VeiculoLegal").xml
    End If
    
    Set xmlLerTodosTipoBCK = CreateObject("MSXML2.DOMDocument.4.0")
    
    With xmlLerTodosTipoBCK
        .loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice").xml
        .documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
        .documentElement.selectSingleNode("TP_VIGE").Text = "S"
        .documentElement.selectSingleNode("TP_SEGR").Text = "S"
        .loadXML objMIU.Executar(.xml, vntCodErro, vntMensagemErro)
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    End With
    
    Set objMIU = Nothing
        
    Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Carrega o conteúdo dos combos
Private Sub flCarregaComboSistema()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    cboSistema.Clear
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa)
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/@Operacao").Text = "LerTodos"
    
    Call xmlLerTodos.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema").xml, _
                                             vntCodErro, _
                                             vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If Not xmlLerTodos.xml = Empty Then
    
        For Each xmlDomNode In xmlLerTodos.selectSingleNode("//Repeat_Sistema").childNodes
            
            With xmlDomNode
                cboSistema.AddItem Trim(.selectSingleNode("SG_SIST").Text) & " - " & .selectSingleNode("NO_SIST").Text
            End With
        Next
    
    End If

    Set xmlLerTodos = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaComboSistema", 0
    
End Sub

Private Sub flCarregaComboVeiculoLegal()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set xmlDomNode = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_SEGR").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_BKOF").Text = gintTipoBackoffice
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal").xml
    strLerTodos = objMIU.Executar(strPropriedades, _
                                  vntCodErro, _
                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)
    
    fgCarregarCombos cboGrupo, xmlLerTodos, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA", False, lstTipoBCKAux, "TP_BKOF"
    
    Set xmlLerTodos = Nothing

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmConsultaVeiculoLegal", "flCarregaComboVeiculoLegal", 0

End Sub

Private Function flRetornaIndiceCombo(ByRef pCombo As ComboBox, ByVal pstrDescricao As String, Optional pblnObterPorDesc As Boolean = False) As Integer

On Error Resume Next
    
    Dim intContador As Integer
    Dim strID As String
    
    If pCombo.ListCount = 0 Then Exit Function
    flRetornaIndiceCombo = -1
    
    If Not pblnObterPorDesc Then
        strID = fgObterCodigoCombo(pstrDescricao)
    Else
        strID = Trim(pstrDescricao)
    End If
    
    For intContador = 0 To pCombo.ListCount - 1
        If Trim(fgObterCodigoCombo(pCombo.List(intContador))) = Trim(strID) Then
            flRetornaIndiceCombo = intContador
            Exit For
        End If
    Next
        
End Function

Private Sub txtCNPJ_LostFocus()
    txtCNPJ.Text = fgFormataCnpj(fgLimpaCaracteresCNPJ(txtCNPJ.Text))
End Sub

Private Sub txtVeiculoLegal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub flOcultarControleCadastro()
    fraCadastro.Visible = False
    tlbCadastro.Visible = False
End Sub

Private Sub flCarregarTipoTitularBMA()

#If EnableSoap = 1 Then
    Dim objVeiculoLegal     As MSSOAPLib30.SoapClient30
#Else
    Dim objVeiculoLegal     As A8MIU.clsMensagem
#End If

Dim xmlRetorno              As MSXML2.DOMDocument40
Dim strRetorno              As String
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objVeiculoLegal = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetorno = objVeiculoLegal.ObterDominioSPB("TpTitlar", _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboTipoTiularBMA.AddItem "", 0
    
    If xmlRetorno.loadXML(strRetorno) Then
        
        For Each objDomNode In xmlRetorno.selectNodes("//Grupo_DominioAtributo")
            cboTipoTiularBMA.AddItem objDomNode.selectSingleNode("CO_DOMI").Text & _
                " - " & objDomNode.selectSingleNode("DE_DOMI").Text
        Next
    
    End If

Exit Sub
ErrorHandler:
    fgLockWindow 0
    fgCursor
    Set objVeiculoLegal = Nothing
    Set xmlRetorno = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarTipoTitularBMA", 0
End Sub
