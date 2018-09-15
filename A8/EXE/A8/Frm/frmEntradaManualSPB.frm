VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEntradaManualSPB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ferramentas - Entrada Manual - Mensagem SPB"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   13725
   Begin VB.Frame fraFiltro 
      Height          =   7170
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         ItemData        =   "frmEntradaManualSPB.frx":0000
         Left            =   135
         List            =   "frmEntradaManualSPB.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   450
         Width           =   4635
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1095
         Width           =   4635
      End
      Begin VB.ComboBox cboServico 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1695
         Width           =   4650
      End
      Begin VB.ComboBox cboEvento 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2310
         Width           =   4650
      End
      Begin MSComctlLib.TreeView treMensagem 
         Height          =   4440
         Left            =   120
         TabIndex        =   6
         Top             =   2685
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   7832
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   406
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgLstMensagem"
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   14
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   13
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Serviço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Evento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   11
         Top             =   2070
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7170
      Left            =   4875
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      Begin VB.TextBox txtDescrticaoMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   6690
      End
      Begin MSComCtl2.DTPicker dtpHoraAgendamento 
         Height          =   315
         Left            =   6870
         TabIndex        =   1
         Top             =   390
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   97452034
         UpDown          =   -1  'True
         CurrentDate     =   38141.4316203704
      End
      Begin FPSpread.vaSpread sprMensagem 
         Height          =   6345
         Left            =   105
         TabIndex        =   16
         Top             =   735
         Width           =   8460
         _Version        =   196608
         _ExtentX        =   14923
         _ExtentY        =   11192
         _StockProps     =   64
         AutoCalc        =   0   'False
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   9
         MaxRows         =   1
         NoBorder        =   -1  'True
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmEntradaManualSPB.frx":0004
         Appearance      =   1
         TextTip         =   1
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin VB.Label Label5 
         Caption         =   "Descrição da Mensagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   195
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Agendamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6870
         TabIndex        =   3
         Top             =   150
         Width           =   1515
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   75
      Top             =   7095
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
            Picture         =   "frmEntradaManualSPB.frx":04CA
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":07E4
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":0AFE
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":0E18
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":1132
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":1584
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":19D6
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   11160
      TabIndex        =   15
      Top             =   7200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ButtonWidth     =   1826
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Enviar"
            Key             =   "Enviar"
            Object.ToolTipText     =   "Enviar"
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
   Begin MSComctlLib.ImageList imgLstMensagem 
      Left            =   735
      Top             =   7095
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":1E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":2142
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSPB.frx":2A1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Atributos em negrito são obrigatórios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5010
      TabIndex        =   17
      Top             =   7290
      Width           =   3285
   End
End
Attribute VB_Name = "frmEntradaManualSPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'' Objeto responsável pelo envio da solicitação (Entrada manual de mensagem SPB) à
'' camada controladora de caso de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsControleAcessDado
''      A8MIU.clsMiu
''      A8MIU.clsMensagem
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlTagMensagem                      As MSXML2.DOMDocument40
Private strOperacao                         As String

Private Const strFuncionalidade             As String = "frmEntradaManual"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private strSepMilhar                        As String
Private strSepDecimal                       As String

'Controle Spread Mensagem------------------------------------------------------------
Private mintControle                        As Integer
Private intUltimaLinha                      As Integer
Private udtTag()                            As udtGrupoRepeticao
Private intTotalRepetNivel()                As Integer

Private Type udtGrupoRepeticao
    Tag                                     As String
    DescricaoTag                            As String
    NivelTag                                As Integer
    OrdemTag                                As Integer
    TipoTag                                 As Long
End Type

'Contantes Colunas Spread
Private Const intColNivelTag                As Integer = 2
Private Const intColOrdemTag                As Integer = 3
Private Const intColNomeTag                 As Integer = 4
Private Const intColTipoTag                 As Integer = 5
Private Const intColDescricaoTag            As Integer = 6
Private Const intPosicaoRelativaGrupo       As Integer = 5

Private strNumCtrlIF                         As String
Private lngCodigoEmpresa                     As Long

'Constants Spread
' Action property settings
Private Const SS_ACTION_ACTIVE_CELL = 0
Private Const SS_ACTION_GOTO_CELL = 1
Private Const SS_ACTION_SELECT_BLOCK = 2
Private Const SS_ACTION_CLEAR = 3
Private Const SS_ACTION_DELETE_COL = 4
Private Const SS_ACTION_DELETE_ROW = 5
Private Const SS_ACTION_INSERT_COL = 6
Private Const SS_ACTION_INSERT_ROW = 7
Private Const SS_ACTION_RECALC = 11
Private Const SS_ACTION_CLEAR_TEXT = 12
Private Const SS_ACTION_PRINT = 13
Private Const SS_ACTION_DESELECT_BLOCK = 14
Private Const SS_ACTION_DSAVE = 15
Private Const SS_ACTION_SET_CELL_BORDER = 16
Private Const SS_ACTION_ADD_MULTISELBLOCK = 17
Private Const SS_ACTION_GET_MULTI_SELECTION = 18
Private Const SS_ACTION_COPY_RANGE = 19
Private Const SS_ACTION_MOVE_RANGE = 20
Private Const SS_ACTION_SWAP_RANGE = 21
Private Const SS_ACTION_CLIPBOARD_COPY = 22
Private Const SS_ACTION_CLIPBOARD_CUT = 23
Private Const SS_ACTION_CLIPBOARD_PASTE = 24
Private Const SS_ACTION_SORT = 25
Private Const SS_ACTION_COMBO_CLEAR = 26
Private Const SS_ACTION_COMBO_REMOVE = 27
Private Const SS_ACTION_RESET = 28
Private Const SS_ACTION_SEL_MODE_CLEAR = 29
Private Const SS_ACTION_VMODE_REFRESH = 30
Private Const SS_ACTION_SMARTPRINT = 32

' Appearance property settings
Private Const SS_APPEARANCE_FLAT = 0
Private Const SS_APPEARANCE_3D = 1
Private Const SS_APPEARANCE_3DWITHBORDER = 2

' BackColorStyle property settings
Private Const SS_BACKCOLORSTYLE_OVERGRID = 0
Private Const SS_BACKCOLORSTYLE_UNDERGRID = 1
Private Const SS_BACKCOLORSTYLE_OVERHORZGRIDONLY = 2
Private Const SS_BACKCOLORSTYLE_OVERVERTGRIDONLY = 3

' ButtonDrawMode property settings
Private Const SS_BDM_ALWAYS = 0
Private Const SS_BDM_CURRENT_CELL = 1
Private Const SS_BDM_CURRENT_COLUMN = 2
Private Const SS_BDM_CURRENT_ROW = 4
Private Const SS_BDM_ALWAYS_BUTTON = 8
Private Const SS_BDM_ALWAYS_COMBO = 16

' CellBorderStyle property settings
Private Const SS_BORDER_STYLE_DEFAULT = 0
Private Const SS_BORDER_STYLE_SOLID = 1
Private Const SS_BORDER_STYLE_DASH = 2
Private Const SS_BORDER_STYLE_DOT = 3
Private Const SS_BORDER_STYLE_DASH_DOT = 4
Private Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Private Const SS_BORDER_STYLE_BLANK = 6
Private Const SS_BORDER_STYLE_FINE_SOLID = 11
Private Const SS_BORDER_STYLE_FINE_DASH = 12
Private Const SS_BORDER_STYLE_FINE_DOT = 13
Private Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Private Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' CellBorderType property settings
Private Const SS_BORDER_TYPE_NONE = 0
Private Const SS_BORDER_TYPE_OUTLINE = 16
Private Const SS_BORDER_TYPE_LEFT = 1
Private Const SS_BORDER_TYPE_RIGHT = 2
Private Const SS_BORDER_TYPE_TOP = 4
Private Const SS_BORDER_TYPE_BOTTOM = 8

' CellType property settings
Private Const SS_CELL_TYPE_DATE = 0
Private Const SS_CELL_TYPE_EDIT = 1
Private Const SS_CELL_TYPE_FLOAT = 2
Private Const SS_CELL_TYPE_INTEGER = 3
Private Const SS_CELL_TYPE_PIC = 4
Private Const SS_CELL_TYPE_STATIC_TEXT = 5
Private Const SS_CELL_TYPE_TIME = 6
Private Const SS_CELL_TYPE_BUTTON = 7
Private Const SS_CELL_TYPE_COMBOBOX = 8
Private Const SS_CELL_TYPE_PICTURE = 9
Private Const SS_CELL_TYPE_CHECKBOX = 10
Private Const SS_CELL_TYPE_OWNER_DRAWN = 11

' ClipboardOptions property settings
Private Const SS_CLIP_NOHEADERS = 0
Private Const SS_CLIP_COPYROWHEADERS = 1
Private Const SS_CLIP_PASTEROWHEADERS = 2
Private Const SS_CLIP_COPYCOLHEADERS = 4
Private Const SS_CLIP_PASTECOLHEADERS = 8
Private Const SS_CLIP_COPYPASTEALLHEADERS = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Private Const SS_HEADER_BLANK = 0
Private Const SS_HEADER_NUMBERS = 1
Private Const SS_HEADER_LETTERS = 2

' CursorStyle property settings
Private Const SS_CURSOR_STYLE_USER_DEFINED = 0
Private Const SS_CURSOR_STYLE_DEFAULT = 1
Private Const SS_CURSOR_STYLE_ARROW = 2
Private Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Private Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Private Const SS_CURSOR_TYPE_DEFAULT = 0
Private Const SS_CURSOR_TYPE_COLRESIZE = 1
Private Const SS_CURSOR_TYPE_ROWRESIZE = 2
Private Const SS_CURSOR_TYPE_BUTTON = 3
Private Const SS_CURSOR_TYPE_GRAYAREA = 4
Private Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Private Const SS_CURSOR_TYPE_COLHEADER = 6
Private Const SS_CURSOR_TYPE_ROWHEADER = 7
Private Const SS_CURSOR_TYPE_DRAGDROPAREA = 8
Private Const SS_CURSOR_TYPE_DRAGDROP = 9

' DAutoSize property settings
Private Const SS_AUTOSIZE_NO = 0
Private Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Private Const SS_AUTOSIZE_BEST_GUESS = 2

' EditEnterAction property settings
Private Const SS_CELL_EDITMODE_EXIT_NONE = 0
Private Const SS_CELL_EDITMODE_EXIT_UP = 1
Private Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Private Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Private Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Private Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Private Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Private Const SS_CELL_EDITMODE_EXIT_SAME = 7
Private Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' OperationMode property settings
Private Const SS_OP_MODE_NORMAL = 0
Private Const SS_OP_MODE_READONLY = 1
Private Const SS_OP_MODE_ROWMODE = 2
Private Const SS_OP_MODE_SINGLE_SELECT = 3
Private Const SS_OP_MODE_MULTI_SELECT = 4
Private Const SS_OP_MODE_EXT_SELECT = 5

' Position property settings
Private Const SS_POSITION_UPPER_LEFT = 0
Private Const SS_POSITION_UPPER_CENTER = 1
Private Const SS_POSITION_UPPER_RIGHT = 2
Private Const SS_POSITION_CENTER_LEFT = 3
Private Const SS_POSITION_CENTER_CENTER = 4
Private Const SS_POSITION_CENTER_RIGHT = 5
Private Const SS_POSITION_BOTTOM_LEFT = 6
Private Const SS_POSITION_BOTTOM_CENTER = 7
Private Const SS_POSITION_BOTTOM_RIGHT = 8

' PrintOrientation property settings
Private Const SS_PRINTORIENT_DEFAULT = 0
Private Const SS_PRINTORIENT_PORTRAIT = 1
Private Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Private Const SS_PRINT_ALL = 0
Private Const SS_PRINT_CELL_RANGE = 1
Private Const SS_PRINT_CURRENT_PAGE = 2
Private Const SS_PRINT_PAGE_RANGE = 3

' ScrollBars property settings
Private Const SS_SCROLLBAR_NONE = 0
Private Const SS_SCROLLBAR_H_ONLY = 1
Private Const SS_SCROLLBAR_V_ONLY = 2
Private Const SS_SCROLLBAR_BOTH = 3

' ScrollBarTrack property settings
Private Const SS_SCROLLBARTRACK_OFF = 0
Private Const SS_SCROLLBARTRACK_VERTICAL = 1
Private Const SS_SCROLLBARTRACK_HORIZONTAL = 2
Private Const SS_SCROLLBARTRACK_BOTH = 3

' SelBackColor property settings
Private Const SPREAD_COLOR_NONE = &H8000000B

' SelectBlockOptions property settings
Private Const SS_SELBLOCKOPT_COLS = 1
Private Const SS_SELBLOCKOPT_ROWS = 2
Private Const SS_SELBLOCKOPT_BLOCKS = 4
Private Const SS_SELBLOCKOPT_ALL = 8

' SortBy property settings
Private Const SS_SORT_BY_ROW = 0
Private Const SS_SORT_BY_COL = 1

' SortKeyOrder property settings
Private Const SS_SORT_ORDER_NONE = 0
Private Const SS_SORT_ORDER_ASCENDING = 1
Private Const SS_SORT_ORDER_DESCENDING = 2

' TextTip property settings
Private Const SS_TEXTTIP_OFF = 0
Private Const SS_TEXTTIP_FIXED = 1
Private Const SS_TEXTTIP_FLOATING = 2
Private Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Private Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

' TypeButtonAlign property settings
Private Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Private Const SS_CELL_BUTTON_ALIGN_TOP = 1
Private Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Private Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' TypeButtonType property settings
Private Const SS_CELL_BUTTON_NORMAL = 0
Private Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeCheckTextAlign property settings
Private Const SS_CHECKBOX_TEXT_LEFT = 0
Private Const SS_CHECKBOX_TEXT_RIGHT = 1

' TypeCheckType property settings
Private Const SS_CHECKBOX_NORMAL = 0
Private Const SS_CHECKBOX_THREE_STATE = 1

' TypeDateFormat property settings
Private Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Private Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Private Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Private Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Private Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Private Const SS_CELL_EDIT_CASE_NO_CASE = 1
Private Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Private Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Private Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Private Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Private Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeHAlign property settings
Private Const SS_CELL_H_ALIGN_LEFT = 0
Private Const SS_CELL_H_ALIGN_RIGHT = 1
Private Const SS_CELL_H_ALIGN_CENTER = 2

' TypeTextAlignVert property settings
Private Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Private Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Private Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Private Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Private Const SS_CELL_TIME_24_HOUR_CLOCK = 1

' TypeVAlign property settings
Private Const SS_CELL_V_ALIGN_TOP = 0
Private Const SS_CELL_V_ALIGN_BOTTOM = 1
Private Const SS_CELL_V_ALIGN_VCENTER = 2

' UnitType property settings
Private Const SS_CELL_UNIT_NORMAL = 0
Private Const SS_CELL_UNIT_VGA = 1
Private Const SS_CELL_UNIT_TWIPS = 2

' UserResize property settings
Private Const SS_USER_RESIZE_COL = 1
Private Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Private Const SS_USER_RESIZE_DEFAULT = 0
Private Const SS_USER_RESIZE_ON = 1
Private Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Private Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Private Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Private Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' ActionKey method settings
Private Const SS_KBA_CLEAR = 0
Private Const SS_KBA_CURRENT = 1
Private Const SS_KBA_POPUP = 2

' AddCustomFunctionExt method Flags parameter settings
Private Const SS_CUSTFUNC_WANTCELLREF = 1
Private Const SS_CUSTFUNC_WANTRANGEREF = 2

' CFGetParamInfo method Type parameter settings
Private Const SS_VALUE_TYPE_LONG = 0
Private Const SS_VALUE_TYPE_DOUBLE = 1
Private Const SS_VALUE_TYPE_STR = 2
Private Const SS_VALUE_TYPE_CELL = 3
Private Const SS_VALUE_TYPE_RANGE = 4

' CFGetParamInfo method Status parameter settings
Private Const SS_VALUE_STATUS_OK = 0
Private Const SS_VALUE_STATUS_ERROR = 1
Private Const SS_VALUE_STATUS_EMPTY = 2

' GetRefStyle/SetRefStyle methods return values/parameter settings
Private Const SS_REFSTYLE_DEFAULT = 0
Private Const SS_REFSTYLE_A1 = 1
Private Const SS_REFSTYLE_R1C1 = 2

' PrintOptions method PageOrder parameter settings
Private Const SS_PAGEORDER_AUTO = 0
Private Const SS_PAGEORDER_DOWNTHENOVER = 1
Private Const SS_PAGEORDER_OVERTHENDOWN = 2

' TextTipFetch method MultiLine parameter settings
Private Const SS_TT_MULTILINE_SINGLE = 0
Private Const SS_TT_MULTILINE_MULTI = 1
Private Const SS_TT_MULTILINE_AUTO = 2

' *************************  PrintPreview Settings *************************

' GrayAreaMarginType property values
Private Const SPV_GRAYAREAMARGINTYPE_SCALED = 0
Private Const SPV_GRAYAREAMARGINTYPE_ACTUAL = 1

' MousePointer property values
Private Const SPV_MOUSEPOINTER_DEFAULT = 0
Private Const SPV_MOUSEPOINTER_ARROW = 1
Private Const SPV_MOUSEPOINTER_CROSS = 2
Private Const SPV_MOUSEPOINTER_I_BEAM = 3
Private Const SPV_MOUSEPOINTER_ICON = 4
Private Const SPV_MOUSEPOINTER_SIZE = 5
Private Const SPV_MOUSEPOINTER_SIZE_NE_SW = 6
Private Const SPV_MOUSEPOINTER_SIZE_N_S = 7
Private Const SPV_MOUSEPOINTER_SIZE_NW_SE = 8
Private Const SPV_MOUSEPOINTER_SIZE_W_E = 9
Private Const SPV_MOUSEPOINTER_UP_ARROW = 10
Private Const SPV_MOUSEPOINTER_HOURGLASS = 11
Private Const SPV_MOUSEPOINTER_NO_DROP = 12

' PageViewType property values
Private Const SPV_PAGEVIEWTYPE_WHOLE_PAGE = 0
Private Const SPV_PAGEVIEWTYPE_NORMAL_SIZE = 1
Private Const SPV_PAGEVIEWTYPE_PERCENTAGE = 2
Private Const SPV_PAGEVIEWTYPE_PAGE_WIDTH = 3
Private Const SPV_PAGEVIEWTYPE_PAGE_HEIGHT = 4
Private Const SPV_PAGEVIEWTYPE_MULTIPLE_PAGES = 5

' ScrollBarH property values
Private Const SPV_SCROLLBARH_SHOW = 0
Private Const SPV_SCROLLBARH_AUTO = 1
Private Const SPV_SCROLLBARH_HIDE = 2

' ScrollBarV property values
Private Const SPV_SCROLLBARV_SHOW = 0
Private Const SPV_SCROLLBARV_AUTO = 1
Private Const SPV_SCROLLBARV_HIDE = 2

' ZoomState property values
Private Const SPV_ZOOMSTATE_INDETERMINATE = 0
Private Const SPV_ZOOMSTATE_IN = 1
Private Const SPV_ZOOMSTATE_OUT = 2
Private Const SPV_ZOOMSTATE_SWITCH = 3

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao        As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmEntradaManual", "flInicializar")
    End If
    
    Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmAtributo", "flInit", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Function flCarregarComboGrupoMensagem() As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim xmlNode                 As MSXML2.IXMLDOMNode
Dim xmlMensagem             As MSXML2.DOMDocument40
Dim strMensagem             As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboGrupo.Clear
    
    strMensagem = objMensagem.LerTodosGrupoMensagem(False, _
                                                    vntCodErro, _
                                                    vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem = "" Then Exit Function
    
    If Not xmlMensagem.loadXML(strMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregarComboGrupoMensagem"
    End If
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        cboGrupo.AddItem xmlNode.selectSingleNode("CO_GRUP").Text & " - " & xmlNode.selectSingleNode("NO_GRUP").Text
        cboGrupo.ItemData(cboGrupo.NewIndex) = xmlNode.selectSingleNode("SQ_GRUP").Text
    Next
    
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboGrupoMensagem", 0
End Function

Private Function flCarregarComboServico() As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboServico.Clear
    cboEvento.Clear
    
    strMensagem = objMensagem.LerTodosServico(cboGrupo.ItemData(cboGrupo.ListIndex), _
                                              False, _
                                              vntCodErro, _
                                              vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem = "" Then Exit Function
    
    If Not xmlMensagem.loadXML(strMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregarComboServico"
    End If
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        cboServico.AddItem xmlNode.selectSingleNode("NO_SERV").Text
        cboServico.ItemData(cboServico.NewIndex) = xmlNode.selectSingleNode("SQ_SERV").Text
    Next
    
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboServico", 0
    
End Function

Private Function flCarregarComboEvento() As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboEvento.Clear
    
    strMensagem = objMensagem.LerTodosEvento(cboServico.ItemData(cboServico.ListIndex), _
                                             False, _
                                             vntCodErro, _
                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem = "" Then Exit Function
    
    If Not xmlMensagem.loadXML(strMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregarComboEvento"
    End If
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        cboEvento.AddItem xmlNode.selectSingleNode("CO_EVEN").Text & " - " & xmlNode.selectSingleNode("NO_EVEN").Text
        cboEvento.ItemData(cboEvento.NewIndex) = xmlNode.selectSingleNode("SQ_EVEN").Text
    Next
    
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboEvento", 0
    
End Function

Private Sub flCarregaTreeViewMensagem(Optional plngSequenciaGrupo As Long, _
                                      Optional plngSequenciaServico As Long, _
                                      Optional plngSequenciaEvento As Long)
                               

#If EnableSoap = 1 Then
    Dim objMensagem                             As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                             As A8MIU.clsMensagem
#End If

Dim xmlMensagem                                 As MSXML2.DOMDocument40
Dim xmlNode                                     As MSXML2.IXMLDOMNode
Dim strXMLMensagem                              As String

Dim objNode                                     As Node
Dim strGrupo                                    As String
Dim strServico                                  As String
Dim strEvento                                   As String
Dim strMensagem                                 As String
Dim vntCodErro                                  As Variant
Dim vntMensagemErro                             As Variant

On Error GoTo ErrorHandler
    
    treMensagem.Nodes.Clear
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLMensagem = objMensagem.LerTodosMensagem(plngSequenciaGrupo, _
                                                  plngSequenciaServico, _
                                                  plngSequenciaEvento, _
                                                  False, _
                                                  "", _
                                                  vntCodErro, _
                                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
        
    If Not xmlMensagem.loadXML(strXMLMensagem) Then
        Exit Sub
        'fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregaTreeViewMensagem"
        
    End If
        
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        
        strGrupo = "G" & fgCompletaString(xmlNode.selectSingleNode("SQ_GRUP").Text, "0", 10, True)
        Set objNode = treMensagem.Nodes.Add(, tvwChild, "G" & strGrupo, xmlNode.selectSingleNode("CO_GRUP").Text, 1)
        objNode.Expanded = True
        objNode.Tag = "G"

        strServico = "S" & fgCompletaString(xmlNode.selectSingleNode("SQ_SERV").Text, "0", 10, True)
        Set objNode = treMensagem.Nodes.Add("G" & strGrupo, tvwChild, "S" & strGrupo & strServico, xmlNode.selectSingleNode("NO_SERV").Text, 2)
        objNode.Expanded = True
        objNode.Tag = "S"
        
        strEvento = "E" & fgCompletaString(xmlNode.selectSingleNode("SQ_EVEN").Text, "0", 10, True)
            
        strMensagem = "M" & fgCompletaString(xmlNode.selectSingleNode("SQ_MESG").Text, "0", 10, True) & Trim(xmlNode.selectSingleNode("NO_TAG_PRIN_MESG").Text)
        Set objNode = treMensagem.Nodes.Add("S" & strGrupo & strServico, tvwChild, "M" & strGrupo & strServico & strEvento & strMensagem, xmlNode.selectSingleNode("CO_MESG").Text & "-" & xmlNode.selectSingleNode("NO_MESG").Text, 3)
        objNode.Tag = "M"
    Next
    
Exit Sub
ErrorHandler:
    
    If Err.Number = 35602 Then
         Resume Next
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'' É acionado pelo botão 'Enviar' da barra de ferramentas.Tem como função,
'' encaminhar a solicitação (Envio de mensagem SPB manualmente) à camada
'' controladora de caso de uso (componente / classe / metodo ) : A8MIU.clsMensagem.
'' EnviarMensagem
Private Sub flEnviarMensagem()

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
    Dim objControleAcesso                   As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
    Dim objControleAcesso                   As A8MIU.clsControleAcessDado
#End If

Dim strMensagem                             As String
Dim plngCodigoEmpresa                       As Long
Dim strRetorno                              As String
Dim strAgendamento                          As String

Dim intTipoBackOffice                       As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Const strSiglaSistema                       As String = "A8"
Const lngCodigoLocalLiquidacao              As Long = 0
Const strCodigoVeiculoLegal                 As String = ""

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
        
    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
    
    plngCodigoEmpresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)

    strMensagem = flMontaMensagem("")
    
    If IsNull(dtpHoraAgendamento.value) Then
        strAgendamento = vbNullString
    Else
        strAgendamento = Mid$(fgDtHr_To_Xml(dtpHoraAgendamento.value), 9, 4)
    End If
    
    If Not flValidacaoFormalMensagem(strMensagem) Then Exit Sub
    
    fgCursor True
    
    Set objControleAcesso = fgCriarObjetoMIU("A8MIU.clsControleAcessDado")
    intTipoBackOffice = objControleAcesso.ObterTipoBackOfficeUsuario(vntCodErro, _
                                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcesso = Nothing
       
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    Call objMensagem.EnviarMensagem(strMensagem, _
                                    enumMensagemEntradaManual.NaoTratadaSLCC, _
                                    plngCodigoEmpresa, _
                                    intTipoBackOffice, _
                                    lngCodigoLocalLiquidacao, _
                                    strCodigoVeiculoLegal, _
                                    strSiglaSistema, _
                                    strAgendamento, _
                                    vbNullString, _
                                    vntCodErro, _
                                    vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    MsgBox "Mensagem gerada com sucesso.", vbInformation, "Entrada Manual"
    
    flFormatarIncluir
    
    Set objMensagem = Nothing
    
    fgCursor False
    
    Exit Sub
ErrorHandler:
        
    Set objControleAcesso = Nothing
    Set objMensagem = Nothing
    fgCursor False
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flEnviarMensagem", 0
    
End Sub

Private Function flValidarCampos() As String
    
Dim lngCont                                     As Long
Dim vntConteudo                                 As Variant
    
On Error GoTo ErrorHandler
    
    If cboEmpresa.ListIndex = -1 Then
        flValidarCampos = "Selecione uma Empresa."
        cboEmpresa.SetFocus
        Exit Function
    End If
    
    If treMensagem.Nodes.Count = 0 Then
        flValidarCampos = "Selecione uma mensagem."
        treMensagem.SetFocus
        Exit Function
    End If
    
    If treMensagem.SelectedItem Is Nothing Then
        flValidarCampos = "Selecione uma mensagem."
        treMensagem.SetFocus
        Exit Function
    End If
    
    For lngCont = 1 To sprMensagem.MaxRows
        
        sprMensagem.Row = lngCont
        sprMensagem.GetText intColNomeTag, lngCont, vntConteudo
        sprMensagem.Col = intColDescricaoTag
        
        If Mid(vntConteudo, 1, 6) <> "Repet_" And _
           Mid(vntConteudo, 1, 6) <> "Grupo_" And _
           Mid(vntConteudo, 1, 7) <> "/Repet_" And _
           Mid(vntConteudo, 1, 7) <> "/Grupo_" Then
        
            
            
            If sprMensagem.FontBold And Trim(vntConteudo) <> "" Then
                    
                vntConteudo = vbNullString
                    
                sprMensagem.GetText sprMensagem.MaxCols, lngCont, vntConteudo
                
                If Trim(vntConteudo) = vbNullString Or vntConteudo = 0 Then
                    sprMensagem.GetText intColDescricaoTag, lngCont, vntConteudo
                    
                    If Trim(vntConteudo) = "" Then
                        sprMensagem.GetText intColDescricaoTag + 1, lngCont, vntConteudo
                    End If
                    
                    If Trim(vntConteudo) = "" Then
                        sprMensagem.GetText intColDescricaoTag + 2, lngCont, vntConteudo
                    End If
                    
                    flValidarCampos = "Campo " & vntConteudo & " preenchimento obrigatório."
                    Exit Function
                End If
                            
            End If
        End If
    Next
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, "frmEntradaManualSPB", "flValidarCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

Private Function flCarregarComboEmpresa() As Boolean

Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    For Each xmlNode In xmlMapaNavegacao.selectSingleNode("frmEntradaManual/Grupo_Dados/Repeat_Empresa").childNodes
        cboEmpresa.AddItem Trim$(xmlNode.selectSingleNode("CO_EMPR").Text) & " - " & Trim$(xmlNode.selectSingleNode("NO_REDU_EMPR").Text)
        cboEmpresa.ItemData(cboEmpresa.NewIndex) = CLng(xmlNode.selectSingleNode("CO_EMPR").Text)
    Next
        
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboEmpresa", 0
    
End Function

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    flFormatarIncluir

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboGrupo_Click", Me.Caption

End Sub

Private Sub cboEvento_Click()

On Error GoTo ErrorHandler

    If cboEvento.ListIndex = -1 Then Exit Sub
    
    fgCursor True
    
    flCarregaTreeViewMensagem cboGrupo.ItemData(cboGrupo.ListIndex), _
                              cboServico.ItemData(cboServico.ListIndex), _
                              cboEvento.ItemData(cboEvento.ListIndex)

    fgCursor False
    
Exit Sub
ErrorHandler:
    
    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboEvento_Click", Me.Caption

End Sub

Private Sub cboGrupo_Click()
    
On Error GoTo ErrorHandler

    If cboGrupo.ListIndex = -1 Then Exit Sub
        
    fgCursor True
        
    flCarregarComboServico
    flCarregaTreeViewMensagem cboGrupo.ItemData(cboGrupo.ListIndex)
    
    fgCursor False

Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboGrupo_Click", Me.Caption
    
End Sub

Private Sub cboServico_Click()

On Error GoTo ErrorHandler

    If cboServico.ListIndex = -1 Then Exit Sub
    
    fgCursor True
    
    flCarregarComboEvento
    flCarregaTreeViewMensagem cboGrupo.ItemData(cboGrupo.ListIndex), _
                              cboServico.ItemData(cboServico.ListIndex)
    fgCursor False
    
Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboServico_Click", Me.Caption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Me.Icon = mdiLQS.Icon
    
    fgCenterMe Me
    
    Me.Show
    DoEvents
    
    strSepDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    strSepMilhar = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SThousand")
    dtpHoraAgendamento.value = Time
    dtpHoraAgendamento.value = Null
    
    flInicializar
    flFormatarIncluir
    
    flCarregarComboEmpresa
    flCarregarComboGrupoMensagem
        
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - Form_Load", Me.Caption

End Sub


Private Sub sprMensagem_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

Dim lngColWidthToTwips01                    As Long
Dim lngColWidthToTwips02                    As Long
Dim lngColWidthToTwips03                    As Long
Dim blnStatusReDraw                          As Boolean
    
On Error GoTo ErrorHandler
    
    Call sprMensagem.ColWidthToTwips(sprMensagem.ColWidth(6), lngColWidthToTwips01)
    Call sprMensagem.ColWidthToTwips(sprMensagem.ColWidth(9), lngColWidthToTwips02)
    
    If Col1 <> 6 And Col1 <> 9 Then
        Call sprMensagem.ColWidthToTwips(sprMensagem.ColWidth(9), lngColWidthToTwips03)
    End If
    
    DoEvents
    
    blnStatusReDraw = sprMensagem.ReDraw
    sprMensagem.ReDraw = True
    
    If (lngColWidthToTwips01 + lngColWidthToTwips02 + lngColWidthToTwips03) > sprMensagem.Width Then

        'If sprMensagem.ScrollBars = ScrollBarsVertical Then
        '    sprMensagem.ScrollBars = ScrollBarsBoth
        'Else
        '    sprMensagem.ScrollBars = ScrollBarsHorizontal
        'End If
        
    ElseIf sprMensagem.ScrollBars = ScrollBarsBoth Then
        'sprMensagem.ScrollBars = ScrollBarsVertical
    End If
    
    sprMensagem.Refresh
    sprMensagem.ReDraw = blnStatusReDraw
    
Exit Sub
ErrorHandler:

    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - " & "sprMensagem_ColWidthChange", Me.Caption

End Sub

Private Sub sprMensagem_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo ErrorHandler

    sprMensagem.Row = NewRow
    sprMensagem.Col = NewCol

    If sprMensagem.Col <> sprMensagem.MaxCols Then
        mintControle = Val(sprMensagem.Text)
    Else
        mintControle = 0
    End If

Exit Sub
ErrorHandler:
    
    mintControle = 0
    
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - sprMensagem_LeaveCell", Me.Caption

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Enviar"
            flEnviarMensagem
        Case gstrSair
            Unload Me
    End Select
        
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - tlbCadastro_ButtonClick", Me.Caption

End Sub

Private Sub treMensagem_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lngCodigoGrupo                        As Long
Dim lngCodigoServico                      As Long
Dim lngCodigoEvento                       As Long
Dim lngCodigoMensagem                     As Long
Dim strCodigoMensagem                     As String

On Error GoTo ErrorHandler
    
    fgCursor True
        
    Set xmlTagMensagem = Nothing
        
    DoEvents
    
    Select Case Left$(Node.Key, 1)
        
        Case "G"
            fgSearchItemCombo cboGrupo, Val(Mid(Node.Key, 3, 10))
        Case "S"
            fgSearchItemCombo cboServico, Val(Mid(Node.Key, 14, 10))
        Case "E"
            fgSearchItemCombo cboEvento, Val(Mid(Node.Key, 25, 10))
        Case "M"
            lngCodigoGrupo = Val(Mid(Node.Key, 3, 10))
            lngCodigoServico = Val(Mid(Node.Key, 14, 10))
            lngCodigoEvento = Val(Mid(Node.Key, 25, 10))
            lngCodigoMensagem = Val(Mid(Node.Key, 36, 10))
            txtDescrticaoMensagem.Tag = Val(Mid(Node.Key, 36, 10))
            strCodigoMensagem = Mid(Node.Text, 1, InStr(1, Node.Text, "-") - 1)
            Call flCarregaSpreadMensagem(strCodigoMensagem)
    End Select
    
    fgCursor False

Exit Sub
ErrorHandler:
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - treMensagem_NodeClick", Me.Caption

End Sub

'' Encaminhar a solicitação (Leitura de todas as mensagens SPB cadastradas, para o
'' preenchimento da planilha Spread) à camada controladora de caso de uso
'' (componente / classe / metodo ) : A8MIU.clsMensagem.LerMensagemO método
'' retornará uma String XML para a camada de interface.
Private Sub flCarregaSpreadMensagem(ByVal pstrCodigoMensagem As String)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim strMensagem                             As String
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim intTamanhoUdt                           As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set xmlTagMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strMensagem = objMensagem.LerMensagem(pstrCodigoMensagem, _
                                          0, _
                                          cboEmpresa.ItemData(cboEmpresa.ListIndex), _
                                          vntCodErro, _
                                          vntMensagemErro, _
                                          0)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    ReDim intTotalRepetNivel(0)
    sprMensagem.MaxRows = 0
    
    If strMensagem = "" Then
        MsgBox "Entrada Manual da mensagem " & pstrCodigoMensagem & " está indisponível.", vbInformation, Me.Caption
        Exit Sub
    End If
     
     xmlTagMensagem.loadXML strMensagem
 
    'Esconder sempre a coluna de Grupo,
    'só será visível se possuir repetição
    sprMensagem.Col = intPosicaoRelativaGrupo + 2
    sprMensagem.ColHidden = True
    sprMensagem.ColWidth(intPosicaoRelativaGrupo + 2) = 20
    
    sprMensagem.Col = intPosicaoRelativaGrupo + 3
    sprMensagem.ColHidden = True
    sprMensagem.ColWidth(intPosicaoRelativaGrupo + 3) = 20
    
    intTamanhoUdt = -1

    sprMensagem.ReDraw = False
    
    txtDescrticaoMensagem.Text = pstrCodigoMensagem & " - " & xmlTagMensagem.documentElement.selectSingleNode("//NO_MESG").Text
    
    For Each xmlNode In xmlTagMensagem.documentElement.childNodes
    
    With xmlNode

        sprMensagem.MaxRows = sprMensagem.MaxRows + 1
        
        sprMensagem.Row = sprMensagem.MaxRows
        sprMensagem.Col = intColDescricaoTag
        sprMensagem.FontBold = CLng(.selectSingleNode("IN_OBRI").Text) = 1
                
        sprMensagem.Col = intPosicaoRelativaGrupo + 3
        
        sprMensagem.SetText intColNivelTag, sprMensagem.MaxRows, .selectSingleNode("IN_NIVE_REPE").Text
        sprMensagem.SetText intColOrdemTag, sprMensagem.MaxRows, .selectSingleNode("NU_ORDE_TAG").Text
        sprMensagem.SetText intColNomeTag, sprMensagem.MaxRows, .selectSingleNode("NO_TAG").Text

        If .selectSingleNode("SQ_TIPO_TAG") Is Nothing Then
            sprMensagem.SetText intColTipoTag, sprMensagem.MaxRows, ""
        Else
            'sbcTagMensagem.numeroSequenciaTipoTag
            sprMensagem.SetText intColTipoTag, sprMensagem.MaxRows, .selectSingleNode("SQ_TIPO_TAG").Text
        End If

        If .selectSingleNode("TX_DEFA") Is Nothing Then
            sprMensagem.SetText sprMensagem.MaxCols, sprMensagem.MaxRows, ""
        Else
            'sbcTagMensagem.nomePadrao
            sprMensagem.SetText sprMensagem.MaxCols, sprMensagem.MaxRows, .selectSingleNode("TX_DEFA").Text
        End If

        If .selectSingleNode("IN_NIVE_REPE").Text > 1 Then
            sprMensagem.Col = .selectSingleNode("IN_NIVE_REPE").Text + intPosicaoRelativaGrupo
            sprMensagem.ColHidden = False
            sprMensagem.SetText sprMensagem.Col, sprMensagem.MaxRows, "001-" & .selectSingleNode("DE_TAG").Text
            sprMensagem.FontBold = CLng(.selectSingleNode("IN_OBRI").Text) = 1
        Else
            'Coluna 6 Descrição da Tag
            sprMensagem.SetText intColDescricaoTag, sprMensagem.MaxRows, .selectSingleNode("DE_TAG").Text
            sprMensagem.FontBold = CLng(.selectSingleNode("IN_OBRI").Text) = 1
        End If

        If Mid(.selectSingleNode("NO_TAG").Text, 1, 6) = "Repet_" Then
            sprMensagem.SetText 1, sprMensagem.MaxRows, "R"    '<-- Substituir por categoria
            sprMensagem.Col = .selectSingleNode("IN_NIVE_REPE").Text + intPosicaoRelativaGrupo + 1
            sprMensagem.ColHidden = False
            sprMensagem.Row = sprMensagem.MaxRows
            sprMensagem.Text = 1
            sprMensagem.CellType = CellTypeInteger
            sprMensagem.TypeHAlign = TypeHAlignRight
            sprMensagem.TypeVAlign = TypeVAlignCenter
            'Adilson 14/08/2003 - Permitir retirada do Repeat para as Msg. LDL0001 e BMA0002
            If .selectSingleNode("IN_OBRI").Text = 1 Then
                sprMensagem.TypeIntegerMin = 1
            Else
                sprMensagem.TypeIntegerMin = 0
            End If
            sprMensagem.TypeIntegerMax = 200
            sprMensagem.TypeSpin = True
            sprMensagem.TypeIntegerSpinInc = 1
        End If

        If intTamanhoUdt = -1 Then
           intTamanhoUdt = 0
           ReDim udtTag(0)
        Else
           intTamanhoUdt = UBound(udtTag) + 1
           ReDim Preserve udtTag(intTamanhoUdt)
        End If

        If .selectSingleNode("DE_TAG") Is Nothing Then
            udtTag(intTamanhoUdt).DescricaoTag = ""
        Else
            udtTag(intTamanhoUdt).DescricaoTag = .selectSingleNode("DE_TAG").Text
        End If
        
        udtTag(intTamanhoUdt).Tag = .selectSingleNode("NO_TAG").Text
        udtTag(intTamanhoUdt).NivelTag = .selectSingleNode("IN_NIVE_REPE").Text
        udtTag(intTamanhoUdt).OrdemTag = .selectSingleNode("NU_ORDE_TAG").Text
        
        If UBound(intTotalRepetNivel) < udtTag(intTamanhoUdt).NivelTag Then
            ReDim Preserve intTotalRepetNivel(udtTag(intTamanhoUdt).NivelTag)
        End If
        
        intTotalRepetNivel(udtTag(intTamanhoUdt).NivelTag) = intTotalRepetNivel(udtTag(intTamanhoUdt).NivelTag) + 1
        
        If .selectSingleNode("SQ_TIPO_TAG") Is Nothing Then
            udtTag(intTamanhoUdt).TipoTag = 0
        Else
            If .selectSingleNode("SQ_TIPO_TAG").Text <> vbNullString Then
                udtTag(intTamanhoUdt).TipoTag = .selectSingleNode("SQ_TIPO_TAG").Text
            End If
        End If
    End With
    
    Next
    
    intUltimaLinha = sprMensagem.MaxRows

    sprMensagem.ReDraw = True
        
    flFormatarCelulasSpread xmlTagMensagem
        
    Set objMensagem = Nothing
        
Exit Sub
ErrorHandler:
    
    fgCursor False
    Set objMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaSpreadMensagem", 0

End Sub

Private Sub flFormatarIncluir()

On Error GoTo ErrorHandler
    
    sprMensagem.MaxRows = 0
    
    txtDescrticaoMensagem.Text = ""
    txtDescrticaoMensagem.Tag = ""
        
    sprMensagem.Col = 7
    sprMensagem.ColHidden = True
    sprMensagem.Col = 8
    sprMensagem.ColHidden = True

    Erase udtTag()
    mintControle = 1

    Exit Sub
    
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flFormatarIncluir", 0
End Sub

Private Sub flFormatarCelulasSpread(ByRef pxmlMensagem As MSXML2.DOMDocument40, Optional ByVal blnCarregaDataHora As Boolean = True)
    
#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlNodeTag                              As MSXML2.IXMLDOMNode
Dim xmlNodeTipoTag                          As MSXML2.IXMLDOMNode
Dim intCont                                 As Integer
Dim strMascara                              As String
Dim lngTamanho                              As Integer
Dim lngCasasDec                             As Integer
Dim strNomeTipoTag                          As String
Dim strNomeTag                              As String
Dim intContCombo                            As Integer
Dim strDominio                              As String
Dim strISPB                                 As String
Dim vntMascaraData                          As Variant
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strNumCtrlIF = objMensagem.ObterNumeroControleIF(vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    strISPB = objMensagem.ObterISPBIF(cboEmpresa.ItemData(cboEmpresa.ListIndex), _
                                      vntCodErro, _
                                      vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    DoEvents
    
    For intCont = 1 To sprMensagem.MaxRows
        
        sprMensagem.Row = intCont
        sprMensagem.Col = intColTipoTag

        If Val(sprMensagem.Text) <> 0 Then
            
            Set xmlNodeTag = pxmlMensagem.documentElement.selectSingleNode("Grupo_Mensagem[SQ_TIPO_TAG='" & Val(sprMensagem.Text) & "']")
            
            Set xmlNodeTipoTag = xmlNodeTag.selectSingleNode("Grupo_TipoTag")
            
            strDominio = ""
            strDominio = xmlNodeTipoTag.selectSingleNode("DOMI").Text
            
            If Trim(strDominio) <> "" Then
                sprMensagem.Col = sprMensagem.MaxCols
                sprMensagem.CellType = CellTypeComboBox
                sprMensagem.TypeComboBoxEditable = False
                sprMensagem.TypeComboBoxList = strDominio
                sprMensagem.TypeMaxEditLen = 50
                sprMensagem.TypeHAlign = TypeHAlignRight
                sprMensagem.TypeVAlign = TypeVAlignCenter
            Else
                'Formatar a célula com o tipo de dado equivalente
                
                strNomeTipoTag = Trim(xmlNodeTipoTag.selectSingleNode("NO_TIPO_TAG").Text)

                sprMensagem.Col = intColNomeTag
                
                strNomeTag = sprMensagem.Text
                lngTamanho = CLng(xmlNodeTipoTag.selectSingleNode("NU_TAMA_TAG").Text)
                lngCasasDec = CLng(xmlNodeTipoTag.selectSingleNode("QT_CASA_DECI").Text)
                
                sprMensagem.Col = intColNomeTag
                
                If CLng(xmlNodeTipoTag.selectSingleNode("IN_TIPO_CTER").Text) = 0 Then
                    If Mid(UCase(strNomeTipoTag), 1, 9) = "DATA HORA" Then
                        
                        sprMensagem.Col = sprMensagem.MaxCols
                        sprMensagem.CellType = CellTypePic
                        sprMensagem.TypePicMask = "99-99-9999 99:99:99"
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        If blnCarregaDataHora = True Then
                            sprMensagem.SetText sprMensagem.MaxCols, intCont, Format(fgDataHoraServidor(DataHoraAux), "dd-mm-yyyy HH:mm:ss")
                        End If
                    
                    ElseIf Mid(UCase(strNomeTipoTag), 1, 4) = "DATA" Then

                        sprMensagem.Col = sprMensagem.MaxCols
                        sprMensagem.CellType = CellTypeDate
                        sprMensagem.TypeDateCentury = True
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        
                        sprMensagem.TypeDateSeparator = Asc("-")
                        
                        gstrMascaraDataDtp = UCase(gstrMascaraDataDtp)
                        vntMascaraData = Split(gstrMascaraDataDtp, gstrSeparadorData, , vbBinaryCompare)
                                            
                        If Left(vntMascaraData(0), 1) = "D" And Left(vntMascaraData(1), 1) = "M" Then
                            sprMensagem.TypeDateFormat = TypeDateFormatDDMMYY
                            If blnCarregaDataHora = True Then
                                sprMensagem.SetText sprMensagem.MaxCols, intCont, Format(fgDataHoraServidor(DataAux), "DD-MM-YYYY")
                            End If
                        ElseIf Left(vntMascaraData(1), 1) = "D" And Left(vntMascaraData(0), 1) = "M" Then
                            sprMensagem.TypeDateFormat = TypeDateFormatMMDDYY
                            If blnCarregaDataHora = True Then
                                sprMensagem.SetText sprMensagem.MaxCols, intCont, Format(fgDataHoraServidor(DataAux), "MM-DD-YYYY")
                            End If
                        ElseIf Left(vntMascaraData(0), 1) = "Y" Then
                            sprMensagem.TypeDateFormat = TypeDateFormatYYMMDD
                            If blnCarregaDataHora = True Then
                                sprMensagem.SetText sprMensagem.MaxCols, intCont, Format(fgDataHoraServidor(DataAux), "YYYY-MM-DD")
                            End If
                        Else
                            sprMensagem.TypeDateFormat = TypeDateFormatDDMMYY
                            If blnCarregaDataHora = True Then
                                sprMensagem.SetText sprMensagem.MaxCols, intCont, Format(fgDataHoraServidor(DataAux), "DD-MM-YYYY")
                            End If
                        End If
                    Else
                        If lngCasasDec <> 0 Then
                            strMascara = String(lngTamanho, "9") & strSepDecimal & String(lngCasasDec, "9")
                            sprMensagem.Col = sprMensagem.MaxCols
                            sprMensagem.CellType = CellTypeFloat
                            sprMensagem.TypeFloatMax = strMascara
                            sprMensagem.TypeFloatDecimalPlaces = lngCasasDec
                            sprMensagem.TypeHAlign = TypeHAlignRight
                            sprMensagem.TypeVAlign = TypeVAlignTop
                            sprMensagem.TypeFloatMoney = False
                            sprMensagem.TypeFloatSeparator = True
                            sprMensagem.TypeFloatDecimalChar = Asc(strSepDecimal)
                            sprMensagem.TypeFloatSepChar = Asc(strSepMilhar)
                        ElseIf lngCasasDec = 0 Then
                            strMascara = String(lngTamanho, "9")
                            sprMensagem.Col = sprMensagem.MaxCols
                            sprMensagem.CellType = CellTypeFloat
                            sprMensagem.TypeFloatMax = strMascara
                            sprMensagem.TypeFloatDecimalPlaces = 0
                            sprMensagem.TypeHAlign = TypeHAlignRight
                            sprMensagem.TypeVAlign = TypeVAlignTop
                            sprMensagem.TypeFloatMoney = False
                            sprMensagem.TypeFloatSeparator = True
                            sprMensagem.TypeFloatDecimalChar = Asc(strSepDecimal)
                            sprMensagem.TypeFloatSepChar = Asc(strSepMilhar)
                                                        
                            'sprMensagem.Col = sprMensagem.MaxCols
                            'sprMensagem.CellType = CellTypeFloat
                            'sprMensagem.TypeEditCharSet = TypeEditCharSetASCII
                            'sprMensagem.TypeEditCharCase = TypeEditCharCaseSetNone
                            'sprMensagem.TypeHAlign = TypeHAlignRight
                            'sprMensagem.TypeVAlign = TypeVAlignTop
                            'sprMensagem.TypeEditMultiLine = False
                            'sprMensagem.TypeEditPassword = False
                            'sprMensagem.TypeMaxEditLen = lngTamanho
                            'sprMensagem.Type
                        End If
                    End If
                Else
                   If Mid(UCase(strNomeTipoTag), 1, 4) = "HORA" Then
                        sprMensagem.Col = sprMensagem.MaxCols
                        sprMensagem.CellType = CellTypePic
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        sprMensagem.TypePicDefaultText = ""
                        sprMensagem.TypePicMask = String(lngTamanho, "9")
                        If blnCarregaDataHora = True Then
                            sprMensagem.SetText sprMensagem.MaxCols, intCont, Format(fgDataHoraServidor(DataHoraAux), "hhmmss")
                        End If
                   ElseIf InStr(1, strNomeTipoTag, "CodMsg") > 0 Then
                        sprMensagem.Col = sprMensagem.MaxCols
                        sprMensagem.BackColor = &HC0FFFF
                        sprMensagem.CellType = CellTypeStaticText
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        sprMensagem.CellBorderType = SS_BORDER_TYPE_OUTLINE
                        sprMensagem.CellBorderStyle = SS_BORDER_STYLE_SOLID
                        sprMensagem.CellBorderColor = RGB(0, 0, 0)
                        sprMensagem.Action = SS_ACTION_SET_CELL_BORDER
                       sprMensagem.SetText sprMensagem.MaxCols, intCont, Trim(Mid(txtDescrticaoMensagem.Text, 1, InStr(1, txtDescrticaoMensagem.Text, "-") - 1))
                   ElseIf InStr(1, strNomeTag, "NumCtrlIF") > 0 Or _
                          InStr(1, strNomeTag, "NumCtrlPart") > 0 Then
                        sprMensagem.Col = sprMensagem.MaxCols
                        sprMensagem.BackColor = &HC0FFFF
                        sprMensagem.CellType = CellTypeStaticText
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        sprMensagem.CellBorderType = SS_BORDER_TYPE_OUTLINE
                        sprMensagem.CellBorderStyle = SS_BORDER_STYLE_SOLID
                        sprMensagem.CellBorderColor = RGB(0, 0, 0)
                        sprMensagem.Action = SS_ACTION_SET_CELL_BORDER
                        sprMensagem.SetText sprMensagem.MaxCols, intCont, strNumCtrlIF
                   ElseIf Trim(strNomeTag) = "ISPBIF" Then
                        sprMensagem.Col = sprMensagem.MaxCols
                        sprMensagem.BackColor = &HC0FFFF
                        sprMensagem.CellType = CellTypeStaticText
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        sprMensagem.CellBorderType = SS_BORDER_TYPE_OUTLINE
                        sprMensagem.CellBorderStyle = SS_BORDER_STYLE_SOLID
                        sprMensagem.CellBorderColor = RGB(0, 0, 0)
                        sprMensagem.Action = SS_ACTION_SET_CELL_BORDER
                       sprMensagem.SetText sprMensagem.MaxCols, intCont, strISPB
                   Else
                       sprMensagem.Col = sprMensagem.MaxCols
                       sprMensagem.CellType = CellTypeEdit
                       sprMensagem.TypeEditCharSet = TypeEditCharSetASCII
                       sprMensagem.TypeEditCharCase = TypeEditCharCaseSetNone
                       sprMensagem.TypeHAlign = TypeHAlignRight
                       sprMensagem.TypeVAlign = TypeVAlignTop
                       sprMensagem.TypeEditMultiLine = False
                       sprMensagem.TypeEditPassword = False
                       sprMensagem.TypeMaxEditLen = lngTamanho
                   End If
                End If
            End If
        Else
            sprMensagem.Col = sprMensagem.MaxCols
            sprMensagem.BackColor = &HC0FFFF
            sprMensagem.CellType = CellTypeStaticText
            sprMensagem.CellBorderType = SS_BORDER_TYPE_OUTLINE
            sprMensagem.CellBorderStyle = SS_BORDER_STYLE_SOLID
            sprMensagem.CellBorderColor = RGB(0, 0, 0)
            sprMensagem.Action = SS_ACTION_SET_CELL_BORDER
        End If
    Next

    If sprMensagem.MaxRows > 26 Then
        'sprMensagem.ScrollBars = ScrollBarsVertical
        sprMensagem.ColWidth(9) = 20
    Else
        'sprMensagem.ScrollBars = ScrollBarsNone
        sprMensagem.ColWidth(9) = 20
    End If
    
     'sprMensagem.ScrollBars = ScrollBarsBoth
     
    Exit Sub
    
ErrorHandler:
    
    Set objMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flFormatarCelulasSpread", 0
    
End Sub

Private Sub sprMensagem_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo ErrorHandler

    fgCursor True
    
    If sprMensagem.TypeIntegerMin = 0 Then
        ControleNumeroRepeticoes Col, Row, False
    Else
        ControleNumeroRepeticoes Col, Row, True
    End If

    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManualSPB - sprMensagem_Change", Me.Caption
End Sub

Private Sub ControleNumeroRepeticoes(ByVal plngCol As Long, _
                                     ByVal plngRow As Long, _
                                     ByVal pblnRepeatObrigatorio As Boolean)

Dim intNumTags                              As Integer
Dim intNumRepet                             As Integer
Dim strAcao                                 As String
Dim intCont                                 As Integer
Dim intContRepet                            As Integer
Dim intAux01                                As Integer
Dim intAux02                                As Integer
Dim intNivelTagRepet                        As Integer
Dim intOrdemTag                             As Integer
Dim intLinhaAtual                           As Integer
Dim intUltimoControle                       As Integer
Dim intFechaTagRepet                        As Integer
Dim intColAtual                             As Integer
Dim vntOrdemTag                             As Variant
Dim intNumDaRepet                           As Integer
Dim intNumSeqRepet                          As Integer

On Error GoTo ErrorHandler

    sprMensagem.Col = plngCol
    sprMensagem.Row = plngRow

    'Coluna de Edição sempre igual a MaxCol
    If plngCol = sprMensagem.MaxCols Then Exit Sub

    'Adilson 13/08/2003 - Foram incluídas duas Mensagens no Catálogo
    '                     com Repetiçoões não obrigatórias
    'Número de repetições não pode ser ZERO (0)
    'Adilson 14/08/2003 - Permitir retirada do Repeat para as Msg. LDL0001 e BMA0002
    If pblnRepeatObrigatorio Then
        If sprMensagem.Text = 0 Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "O Grupo Repeat só pode ser eliminado se não for obrigatório"
            sprMensagem.SetText plngCol, plngRow, 1
            Exit Sub
        End If
    End If
    
    'Se o numero de repetiçoes for Maior ou = a 100 pede Confirmação
    If Val(sprMensagem.Text) >= 100 Then
        If MsgBox("Confirma quantidade de repetição ?", vbYesNo + vbDefaultButton2, "Repetição de Mensagem") = vbNo Then
            Exit Sub
        End If
    End If

    intLinhaAtual = plngRow

    'Verifica a ação a ser tomada (Inclusão ou Exclusão)
    'se o número informado para a repetição aumentar, então Inclusão
    If Val(sprMensagem.Text) > mintControle Then
        strAcao = "I"
    ElseIf Val(sprMensagem.Text) < mintControle Then 'se diminuir, então Exclusão
        strAcao = "E"
    Else 'se não, enviar brancos para abortar a ação
        strAcao = ""
    End If

    If strAcao = "" Then Exit Sub

    intUltimoControle = mintControle
    intNumRepet = Val(sprMensagem.Text) - mintControle
    mintControle = Val(sprMensagem.Text)

    intColAtual = sprMensagem.ActiveCol
    sprMensagem.Col = intColNivelTag
    intNivelTagRepet = Val(sprMensagem.Text)
    sprMensagem.Col = intColOrdemTag
    intOrdemTag = Val(sprMensagem.Text)

    'Pesquisa quantas linhas replicar
    For intCont = intOrdemTag To UBound(udtTag)
        If udtTag((intCont)).NivelTag > intNivelTagRepet Then
            intAux01 = intAux01 + 1
        Else
            Exit For
        End If
    Next

    intNumDaRepet = 0
    'Pesquisa qual o final do Repet para começar a nova Repetição
    For intCont = plngRow + 1 To sprMensagem.MaxRows
        sprMensagem.Row = intCont

        sprMensagem.Col = intColNivelTag
        If sprMensagem.Text = intNivelTagRepet Then

            sprMensagem.Row = sprMensagem.Row - 1
            sprMensagem.Col = intNivelTagRepet + 6

            If sprMensagem.Text <> "" Then
                intNumSeqRepet = Mid(sprMensagem.Text, 1, 3) + 1
            End If

            Exit For
        End If
    Next

    intFechaTagRepet = intCont

    sprMensagem.Col = intColAtual

    If strAcao = "I" Then
        intTotalRepetNivel(intNivelTagRepet + 1) = intTotalRepetNivel(intNivelTagRepet + 1) + (intAux01 * intNumRepet)
        
        sprMensagem.Row = intFechaTagRepet
        sprMensagem.MaxRows = sprMensagem.MaxRows + (intAux01 * intNumRepet)

        For intNumTags = 1 To (intAux01 * intNumRepet)
            sprMensagem.Action = ActionInsertRow
        Next

        For intNumTags = 1 To intNumRepet
            For intCont = intOrdemTag To UBound(udtTag)
                If udtTag((intCont)).NivelTag > intNivelTagRepet Then
                   If udtTag((intCont)).NivelTag > intNivelTagRepet + 1 Then
                        sprMensagem.SetText udtTag((intCont)).NivelTag + 5, sprMensagem.Row, "001-" & udtTag(intCont).DescricaoTag
                   Else
                        sprMensagem.SetText udtTag((intCont)).NivelTag + 5, sprMensagem.Row, Format(intNumSeqRepet + intNumTags - 1, "000") & "-" & udtTag(intCont).DescricaoTag
                   End If
                   If Mid(udtTag((intCont)).Tag, 1, 5) = "Repet" Then
                        sprMensagem.SetText 1, sprMensagem.Row, "R"     '<-- Substituir por categoria
                        sprMensagem.Col = sprMensagem.Col + 1
                        sprMensagem.Row = sprMensagem.Row
                        sprMensagem.Text = "1"
                        sprMensagem.CellType = CellTypeInteger
                        sprMensagem.TypeHAlign = TypeHAlignRight
                        sprMensagem.TypeVAlign = TypeVAlignCenter
                        'Adilson 14/08/2003 - Permitir retirada do Repeat para as Msg. LDL0001 e BMA0002
                        If pblnRepeatObrigatorio Then
                            sprMensagem.TypeIntegerMin = 1
                        Else
                            sprMensagem.TypeIntegerMin = 0
                        End If
                        sprMensagem.TypeIntegerMax = 500
                        sprMensagem.TypeSpin = True
                        sprMensagem.TypeIntegerSpinInc = 1
                        sprMensagem.Col = sprMensagem.Col - 1
                    End If
                    sprMensagem.SetText intColNivelTag, sprMensagem.Row, udtTag(intCont).NivelTag
                    sprMensagem.SetText intColOrdemTag, sprMensagem.Row, udtTag(intCont).OrdemTag
                    sprMensagem.SetText intColNomeTag, sprMensagem.Row, udtTag(intCont).Tag
                    sprMensagem.SetText intColTipoTag, sprMensagem.Row, udtTag(intCont).TipoTag
                    sprMensagem.Row = sprMensagem.Row + 1
                Else
                    Exit For
                End If
            Next
        Next
        intUltimaLinha = sprMensagem.ActiveRow
    Else ' Remover n (intNumRepet) repeticoes
        intTotalRepetNivel(intNivelTagRepet + 1) = intTotalRepetNivel(intNivelTagRepet + 1) + (intAux01 * intNumRepet)
        intAux02 = intAux01
        For intNumDaRepet = 1 To Abs(intNumRepet)
            sprMensagem.Row = intFechaTagRepet - 1
            
            'Carlos
            If mintControle = 0 And intNumDaRepet = Abs(intNumRepet) Then
                sprMensagem.Row = sprMensagem.Row + 1
                intAux01 = intAux02 + 1
            Else
                intAux01 = intAux02 - 1
            End If

            sprMensagem.GetText 3, sprMensagem.Row, vntOrdemTag

            intCont = 0
            DoEvents
            For intNumTags = sprMensagem.Row To (sprMensagem.Row - intAux01) Step -1
                sprMensagem.Col = 3
                sprMensagem.Row = intNumTags
                If intCont = 0 Then
                    sprMensagem.Col = 3
                    intCont = sprMensagem.Text
                    intFechaTagRepet = intFechaTagRepet - 1
                    sprMensagem.Action = ActionDeleteRow
                    sprMensagem.Row = sprMensagem.Row - 1
                    sprMensagem.MaxRows = sprMensagem.MaxRows - 1
                Else
                    sprMensagem.Col = 3
                    intCont = sprMensagem.Text
                    intFechaTagRepet = intFechaTagRepet - 1
                    sprMensagem.Action = ActionDeleteRow
                    sprMensagem.Row = sprMensagem.Row - 1
                    sprMensagem.MaxRows = sprMensagem.MaxRows - 1
                End If
            Next
        Next
    End If
    
    If intTotalRepetNivel(intNivelTagRepet + 1) = 0 Then
        If mintControle = 0 Then
            If intNivelTagRepet = 1 Then
                sprMensagem.Col = 7
                sprMensagem.ColHidden = True
                sprMensagem.ColWidth(7) = 30
            End If
            sprMensagem.Col = 8
            sprMensagem.ColHidden = True
            sprMensagem.ColWidth(8) = 30
        End If
    End If
    
    flFormatarCelulasSpread xmlTagMensagem, False
    sprMensagem.Row = intLinhaAtual
        
Exit Sub
ErrorHandler:

    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - " & "ControleNumeroRepeticoes", Me.Caption
        
End Sub

Private Function flMontaMensagem(ByVal pstrCodigoMensagem As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strNomeTagPrincipal                     As String
Dim strNomeTag                              As String
Dim vntConteudo                             As Variant
Dim lngCont                                 As Long

On Error GoTo ErrorHandler

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    strNomeTagPrincipal = Mid(treMensagem.SelectedItem.Key, 46)
    
    Call fgAppendNode(xmlMensagem, "", "SISMSG", "")

    strMensagem = strMensagem & "<" & Trim(strNomeTagPrincipal) & ">"

    For lngCont = 1 To sprMensagem.MaxRows
        
        sprMensagem.Row = lngCont

        sprMensagem.Col = intColNomeTag
        
        strMensagem = strMensagem & "<" & sprMensagem.Text & ">"
        
        sprMensagem.Col = sprMensagem.MaxCols

        'Alteração efetuada em 28/11/2001
        'tratamento do XML para excluir separador de MILHAR
        If sprMensagem.CellType = CellTypeFloat Then
            'Verifica o formato do campo:
            '
            'Se casas decimais forem informadas, então o campo é numérico...
            If sprMensagem.TypeFloatDecimalPlaces <> 0 Then
                strMensagem = strMensagem & Replace(UCase(Replace(sprMensagem.Text, strSepMilhar, vbNullString)), ".", ",")
            Else
                strMensagem = strMensagem & Replace(UCase(Replace(sprMensagem.Text, strSepMilhar, vbNullString)), ".", "")
            End If

        ElseIf sprMensagem.CellType = CellTypeDate Then
            
            If Trim(sprMensagem.Text) <> "" Then
                If IsDate(sprMensagem.Text) Then
                    strMensagem = strMensagem & fgDate_To_DtXML(sprMensagem.Text)
                End If
            End If

            '...se não, verifica se é HORA...
        ElseIf sprMensagem.CellType = CellTypePic Then
            
            If IsDate(sprMensagem.Text) Then
                strMensagem = strMensagem & fgDateHr_To_DtHrXML(sprMensagem.Text)
            ElseIf Len(sprMensagem.Text) = 6 Then
                strMensagem = strMensagem & sprMensagem.Text
            End If
        
        ElseIf sprMensagem.CellType = CellTypeComboBox Then
            If Trim(sprMensagem.Text) <> "" Then
                strMensagem = strMensagem & Trim(Mid(sprMensagem.Text, 1, InStr(1, sprMensagem.Text, "-") - 1))
            End If
        Else
            strMensagem = strMensagem & Replace(UCase(sprMensagem.Text), ".", ",")
        End If

        sprMensagem.Col = intColNomeTag

        If InStr(1, sprMensagem.Text, "Grupo_") = 0 And _
           InStr(1, sprMensagem.Text, "Repet_") = 0 Then
            strMensagem = strMensagem & "</" & sprMensagem.Text & ">"
        Else
            strMensagem = strMensagem
        End If

    Next lngCont

    strMensagem = strMensagem & "</" & Trim(strNomeTagPrincipal) & ">"

    Call fgAppendXML(xmlMensagem, "SISMSG", strMensagem)

    flMontaMensagem = xmlMensagem.xml

    Set xmlMensagem = Nothing

Exit Function
ErrorHandler:

    Set xmlMensagem = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontaMensagem", 0
    
End Function

Private Function flValidacaoFormalMensagem(ByRef pstrMensagem As String) As Boolean

Dim xmlMensagem                              As MSXML2.DOMDocument40
Dim xmlNodeMensagem                          As MSXML2.IXMLDOMNode
Dim xmlNodeTagMensagem                       As MSXML2.IXMLDOMNode
Dim blnTagObrigatoria                        As Boolean
Dim blnTagRepeatOuGrupoObrigatoria           As Boolean
Dim blnTagsPertencemRepeatOuGrupo            As Boolean
Dim strConteudoTag                           As String
Dim lngCount                                 As Long
Dim strNomeTag                               As String
Dim strNomeLogicoTag                         As String

On Error GoTo ErrorHandler

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    xmlMensagem.loadXML pstrMensagem

    For Each xmlNodeMensagem In xmlMensagem.documentElement.childNodes.Item(0).childNodes
                
        Set xmlNodeTagMensagem = xmlTagMensagem.documentElement.selectSingleNode("Grupo_Mensagem[NO_TAG='" & xmlNodeMensagem.nodeName & "']")
        
        If xmlNodeTagMensagem.selectSingleNode("IN_CATG_TAG").Text = 0 Then

            strNomeTag = Trim(xmlNodeTagMensagem.selectSingleNode("NO_TAG").Text)
            strNomeLogicoTag = Trim(xmlNodeTagMensagem.selectSingleNode("DE_TAG").Text)
            blnTagObrigatoria = IIf(xmlNodeTagMensagem.selectSingleNode("IN_OBRI").Text = 1, True, False)
            strConteudoTag = Trim(xmlNodeMensagem.Text)

            'Validar conteudo Tag
            If xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/IN_TIPO_CTER").Text = 0 Then
                
                'Versão 4.2.2- O conteudo da tag pode ser "" porque a tag não é obrigatória
                If IsNumeric(strConteudoTag) Then
                    'Versão 4.2.1 Validação de Dominio p/ numérico
                    If Len(strConteudoTag) > 28 Then
                        strConteudoTag = CVar(strConteudoTag)
                    Else
                        strConteudoTag = strConteudoTag
                    End If
                End If
                
                If strNomeTag = "DtMovto" Then
                    'Validar conteudo tag data , Erro Neg
                    If Not flValidaTagDataMovto(Trim(strConteudoTag), _
                                                xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NU_TAMA_TAG").Text, _
                                                blnTagObrigatoria) Then
                        flMostraMural "Campo " & strNomeLogicoTag & " conteúdo inválido."
                        Exit Function
                    End If
                Else

                    If xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NO_TIPO_TAG").Text = "Data" Then
                        
                        If Not flValidaTagData(Trim(strConteudoTag), _
                                               xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NU_TAMA_TAG").Text, _
                                               blnTagObrigatoria) Then
                            
                            flMostraMural "Campo " & strNomeLogicoTag & " conteúdo inválido."
                            Exit Function
                        End If
                    ElseIf xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NO_TIPO_TAG").Text = "Data Hora" Then
                        
                        If Not flValidaTagDataHora(Trim(strConteudoTag), _
                                                   xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NU_TAMA_TAG").Text, _
                                                   blnTagObrigatoria) Then
                            
                            flMostraMural "Campo " & strNomeLogicoTag & " conteúdo inválido."
                            Exit Function
                        End If
                    Else
                        'Validar conteudo tag numerico , Erro Neg
                        If Not flValidaTagNumerica(Trim(strConteudoTag), _
                                                   xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NU_TAMA_TAG").Text, _
                                                   xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/QT_CASA_DECI").Text, _
                                                   blnTagObrigatoria) Then
                             flMostraMural "Campo " & strNomeLogicoTag & " conteúdo inválido."
                             Exit Function
                        End If
                    End If
                End If
            Else
                If xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NO_TIPO_TAG").Text = "Hora" Then
                    
                    If Not flValidaTagHora(Trim(strConteudoTag), _
                                           xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NU_TAMA_TAG").Text, _
                                           blnTagObrigatoria) Then
                        flMostraMural "Campo " & strNomeLogicoTag & " conteúdo inválido."
                        Exit Function
                     End If
                Else
                    'Validar conteudo tag Alfa Numerico
                    If Not flValidaTagAlfaNumerica(Trim(strConteudoTag), _
                                                   xmlNodeTagMensagem.selectSingleNode("Grupo_TipoTag/NU_TAMA_TAG").Text, _
                                                   blnTagObrigatoria) Then
                        Exit Function
                End If
            End If
        End If
        'No Else são verificadas as situações onde a Tag de Repeat ou Grupo não são obrigatórias,
        'Neste caso as Tags internas não devem ser validadas, mesmo sendo obrigatórias
        'Adilson - Versão 4.3.4 - 25/08/2003
        'ElseIf (lrsTagMensagem.fields("IN_CATG_TAG") = 2) Then 'Tag Repeat
        '    If Left$(Trim$(lrsTagMensagem.fields("NO_TAG")), 1) = "/" Then
        '        lbTagsPertencemRepeatOuGrupo = False
        '        lbTagRepeatOuGrupoObrigatoria = False
        '    Else
        '        lbTagRepeatOuGrupoObrigatoria = IIf(lrsTagMensagem.fields("IN_OBRI") = 1, True, False)
        '        lbTagsPertencemRepeatOuGrupo = True
        '    End If
        'ElseIf (lrsTagMensagem.fields("IN_CATG_TAG") = 1) Then 'Tag Grupo
        '    If Left$(Trim$(lrsTagMensagem.fields("NO_TAG")), 1) = "/" Then
        '        lbTagsPertencemRepeatOuGrupo = False
        '        lbTagRepeatOuGrupoObrigatoria = False
        '    ElseIf lbTagsPertencemRepeatOuGrupo = False Then
        '        lbTagRepeatOuGrupoObrigatoria = IIf(lrsTagMensagem.fields("IN_OBRI") = 1, True, False)
        '        lbTagsPertencemRepeatOuGrupo = True
        '    End If
        End If
    Next
        
    flValidacaoFormalMensagem = True
    
    Set xmlMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set xmlMensagem = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "ValidacaoFormalMensagem", 0

End Function

Private Function flValidaTagNumerica(ByVal pstrConteudoTAG As String, _
                                     ByVal plngTamanhoInteiro As Long, _
                                     ByVal plngTamanhoDecimais As Long, _
                                     ByVal pblnTAGObrigatoria As Boolean) As Boolean

Dim strConteudoInteiro               As String
Dim strConteudoDecimal               As String
Dim vntArrayConteudo                 As Variant

On Error GoTo ErrorHandler
    
    If Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = False Then
        flValidaTagNumerica = True
        Exit Function
    End If
    
    If Not IsNumeric(pstrConteudoTAG) Then
        'FORMATO DO DADO INVALIDO
        flValidaTagNumerica = False
        Exit Function
    End If
        
    vntArrayConteudo = Split(pstrConteudoTAG, ",")
    
    If UBound(vntArrayConteudo) > 0 Then
        strConteudoInteiro = vntArrayConteudo(0)
        strConteudoDecimal = vntArrayConteudo(1)
    Else
        strConteudoInteiro = vntArrayConteudo(0)
        strConteudoDecimal = ""
    End If
    
    If Val(strConteudoInteiro) < 0 And Val(strConteudoDecimal) < 0 And pblnTAGObrigatoria Then
        'VALOR INVALIDO
        flValidaTagNumerica = False
        Exit Function
    ElseIf Len(strConteudoInteiro) > plngTamanhoInteiro Then
        'VALOR INVALIDO
        flValidaTagNumerica = False
        Exit Function
    End If
    
    If Len(strConteudoDecimal) > 0 Then
        If Len(strConteudoDecimal) > plngTamanhoDecimais Then
            'TAMANHO DO DADO INVALIDO
            flValidaTagNumerica = False
            Exit Function
        End If
    End If
        
    flValidaTagNumerica = True
        
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "ValidacaoFormalMensagem", 0
    
End Function

Private Function flValidaTagAlfaNumerica(ByVal pstrConteudoTAG As String, _
                                         ByVal plngTamanhoTag As Long, _
                                         ByVal pblnTAGObrigatoria As Boolean) As Boolean
                                     
On Error GoTo ErrorHandler
    
    If Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = False Then
        flValidaTagAlfaNumerica = True
        Exit Function
    ElseIf Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = True Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagAlfaNumerica = False
        Exit Function
    End If
    
    If Len(pstrConteudoTAG) > plngTamanhoTag Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagAlfaNumerica = False
        Exit Function
    End If
    
    flValidaTagAlfaNumerica = True
    
Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flValidaTagAlfaNumerica", 0
    
End Function

Private Function flValidaTagDataMovto(ByVal pstrConteudoTAG As String, _
                                      ByVal plngTamanhoTag As Long, _
                                      ByVal pblnTAGObrigatoria As Boolean) As Boolean
                                     
Dim Ano                                 As Integer
Dim Mes                                 As Integer
Dim Dia                                 As Integer
    
On Error GoTo ErrorHandler
    
    If Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = False Then
        flValidaTagDataMovto = True
        Exit Function
    ElseIf Val(pstrConteudoTAG) = 0 And pblnTAGObrigatoria = True Then
        'DATA DE MOVIMENTO NÃO INFORMADA
        flValidaTagDataMovto = False
        Exit Function
    End If
     
    If Not IsNumeric(pstrConteudoTAG) Then
        'DATA MOVIMENTO INVALIDA
        flValidaTagDataMovto = False
        Exit Function
    End If
    
    If Len(pstrConteudoTAG) <> plngTamanhoTag Then
        'DATA MOVIMENTO INVALIDA
        flValidaTagDataMovto = False
        Exit Function
    End If
    
    Ano = Mid(pstrConteudoTAG, 1, 4)
    Mes = Mid(pstrConteudoTAG, 5, 2)
    Dia = Mid(pstrConteudoTAG, 7, 2)
    
    If Not IsDate(Dia & "/" & Mes & "/" & Ano) Then
        'DATA MOVIMENTO INVALIDA
        flValidaTagDataMovto = False
        Exit Function
    End If
    
    flValidaTagDataMovto = True
    
Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flValidaTagDataMovto", 0

End Function

Private Function flValidaTagData(ByVal pstrConteudoTAG As String, _
                                 ByVal plngTamanhoTag As Long, _
                                 ByVal pblnTAGObrigatoria As Boolean) As Boolean
                                     
Dim Ano As Integer, Mes As Integer, Dia As Integer
    
On Error GoTo ErrorHandler
    
    If Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = False Then
        flValidaTagData = True
        Exit Function
    ElseIf Val(pstrConteudoTAG) = 0 And pblnTAGObrigatoria = True Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagData = False
        Exit Function
    End If
     
    If Not IsNumeric(pstrConteudoTAG) Then
        'FORMATO DO DADO INVALIDO
        flValidaTagData = False
        Exit Function
    End If
    
    If Len(pstrConteudoTAG) <> plngTamanhoTag Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagData = False
        Exit Function
    End If
    
    Ano = Mid(pstrConteudoTAG, 1, 4)
    Mes = Mid(pstrConteudoTAG, 5, 2)
    Dia = Mid(pstrConteudoTAG, 7, 2)
    
    If Not IsDate(Ano & "/" & Mes & "/" & Dia) Then
        'FORMATO DO DADO INVALIDO
        flValidaTagData = False
        Exit Function
    End If
    
    flValidaTagData = True
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flValidaTagData", 0
    
End Function

Private Function flValidaTagDataHora(ByVal pstrConteudoTAG As String, _
                                     ByVal plngTamanhoTag As Long, _
                                     ByVal pblnTAGObrigatoria As Boolean) As Boolean
                                     
Dim Ano                                     As Integer
Dim Mes                                     As Integer
Dim Dia                                     As Integer
Dim Hora                                    As Integer
Dim Min                                     As Integer
Dim Seg                                     As Integer
    
On Error GoTo ErrorHandler
    
    If Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = False Then
        flValidaTagDataHora = True
        Exit Function
    ElseIf Val(pstrConteudoTAG) = 0 And pblnTAGObrigatoria = True Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    End If
     
    If Not IsNumeric(pstrConteudoTAG) Then
        'FORMATO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    End If
    
    If Len(pstrConteudoTAG) <> plngTamanhoTag Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    End If
    
    Ano = Mid(pstrConteudoTAG, 1, 4)
    Mes = Mid(pstrConteudoTAG, 5, 2)
    Dia = Mid(pstrConteudoTAG, 7, 2)
    
    If Not IsDate(Dia & "/" & Mes & "/" & Ano) Then
        'FORMATO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    End If
    
    Hora = Mid(pstrConteudoTAG, 9, 2)
    Min = Mid(pstrConteudoTAG, 11, 2)
    Seg = Mid(pstrConteudoTAG, 13, 2)
    
    If Hora > 24 Then
        'FORMATO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    ElseIf Min > 60 Then
        'FORMATO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    ElseIf Seg > 60 Then
        'FORMATO DO DADO INVALIDO
        flValidaTagDataHora = False
        Exit Function
    End If
    
    flValidaTagDataHora = True
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flValidaTagDataHora", 0

End Function

Private Function flValidaTagHora(ByVal pstrConteudoTAG As String, _
                                 ByVal plngTamanhoTag As Long, _
                                 ByVal pblnTAGObrigatoria As Boolean) As Boolean
                                     
Dim Hora                                As Integer
Dim Min                                 As Integer
Dim Seg                                 As Integer
    
On Error GoTo ErrorHandler
    
    If Len(Trim(pstrConteudoTAG)) = 0 And pblnTAGObrigatoria = False Then
        flValidaTagHora = True
        Exit Function
    ElseIf Val(pstrConteudoTAG) = 0 And pblnTAGObrigatoria = True Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagHora = False
        Exit Function
    End If
     
    If Not IsNumeric(pstrConteudoTAG) Then
        'FORMATO DO DADO INVALIDO
        flValidaTagHora = False
        Exit Function
    End If
    
    If Len(pstrConteudoTAG) <> plngTamanhoTag Then
        'TAMANHO DO DADO INVALIDO
        flValidaTagHora = False
        Exit Function
    End If
    
    Hora = Mid(pstrConteudoTAG, 1, 2)
    Min = Mid(pstrConteudoTAG, 3, 2)
    Seg = Mid(pstrConteudoTAG, 5, 2)
    
    If Hora > 24 Then
        'FORMATO DO DADO INVALIDO
        flValidaTagHora = False
        Exit Function
    ElseIf Min > 60 Then
        'FORMATO DO DADO INVALIDO
        flValidaTagHora = False
        Exit Function
    ElseIf Seg > 60 Then
        'FORMATO DO DADO INVALIDO
        flValidaTagHora = False
        Exit Function
    End If
    
    flValidaTagHora = True
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flValidaTagHora", 0

End Function

Private Function flMostraMural(ByVal pstrMensagem As String)
    
On Error GoTo ErrorHandler

    frmMural.txtMural = pstrMensagem
    frmMural.Show

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMostraMural", 0
    
End Function

