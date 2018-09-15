Attribute VB_Name = "basA7"
Attribute VB_Description = "Empresa        : Regerbanc\r\nComponente     : BUS\r\nClasse         : basBUS\r\nData Criação   : 25/06/203\r\nObjetivo       : Modulo bas BUS\r\nAnalista       : Marcelo Kida\r\n\r\nProgramador    : Marcelo Kida\r\nData           : 25/06/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFB2E7F033C"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"
'Empresa        : Regerbanc
'Componente     : BUS
'Classe         : basBUS
'Data Criação   : 25/06/203
'Objetivo       : Modulo bas BUS
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 25/06/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

Public gstrSource                           As String
Public gstrAmbiente                         As String
Public gstrUsuario                          As String
Public gblnAcessoOnLine                     As Boolean
Public gblnRegistraTLB                      As Boolean
Public glngContaMinutosAlerta               As Long
Public gblnPerfilManutencao                 As Boolean
Public XpaMBS                               As Boolean

Public Const gstrOperLerTodos               As String = "LerTodos"
Public Const gstrOperAlterar                As String = "Alterar"
Public Const gstrOperIncluir                As String = "Incluir"
Public Const gstrOperExcluir                As String = "Excluir"

Public gstrURLWebService                    As String
Public glngTimeOut                          As Long

Public gstrHelpFile                         As String
Public gstrPrint                            As String

Public Const ERR_USUARIONAOLOGADO           As Long = 18
Public Const ERR_SEMACESSO                  As Long = 35

'------------------ API ----------------------------------------
'API'S para geração de Arquivo temporário no Windows \Temp
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Attribute GetTempPath.VB_Description = "------------------ API ----------------------------------------\r\nAPI'S para geração de Arquivo temporário no Windows \\Temp"
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long

'API's utilizadas para colocar o form em evidência, ou seja, colocar o form em foco
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'API para Obter ID usuario logado
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'---------------------------------------------------------------

'------------------ Constants ----------------------------------
'Constants BUS
Public Const gstrMascaraDataXml             As String = "YYYYMMDD"
'Váriaveis Globais
Public gstrMascaraDataDtp                   As String
Public gstrMascaraDataHoraDtp               As String

Public strHoraInicioVerificacao             As String
Public strHoraFimVerificacao                As String

'Caracteres especiais
Public Const CAR_APOSTROFE As Long = 39
Attribute CAR_APOSTROFE.VB_VarDescription = "Caracteres especiais"
Public Const CAR_ABRE_CHAVE As Long = 123
Public Const CAR_FECHA_CHAVE As Long = 125
Public Const CAR_ENTER As Long = 13
Public Const CAR_LINEFEED As Long = 10
Public Const CAR_SUBST As Long = 127
Public Const CAR_ASPAS As Long = 34
Public Const CAR_PORCENTO As Long = 37
Public Const CAR_INTERROGACAO As Long = 63
Public Const CAR_CASP1 As Long = 96
Public Const CAR_CASP2 As Long = 180
Public Const CAR_ASPAS1 As Long = 145
Public Const CAR_ASPAS2 As Long = 146
Public Const CAR_BARRA As Long = 47
Public Const CAR_PONTO As Long = 46
Public Const CAR_HIFEN As Long = 45
'---------------------------------------------------------------

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31

'------------------ Variaveis BUS ------------------------------
Private intNumeroSequencialErro As Integer
Attribute intNumeroSequencialErro.VB_VarDescription = "------------------ Variaveis BUS ------------------------------"
Private lngCodigoErroNegocio As Long
'---------------------------------------------------------------

Public Enum enumMenuRegraTraducao
    RegraTraducaoExcluir = 1
    RegraTraducaoNovo = 2
End Enum

'Constantes Controle de Acesso
Public Const PER_A7CADASTRO                 As String = "A7CADASTRO"
Public Const PER_A7PARAMSISTEMA             As String = "A7PARAMSISTEMA"
Public Const PER_A7ATRIBUTOMENSAGEM         As String = "A7ATRIBUTOMENSAGEM"
Public Const PER_A7TIPOMENSAGEM             As String = "A7TIPOMENSAGEM"
Public Const PER_A7REGRATRANSPORTE          As String = "A7REGRATRANSPORTE"
Public Const PER_A7PARAMPOSTAGEM            As String = "A7PARAMPOSTAGEM"
Public Const PER_A7PARAMNOTIFICACAO         As String = "A7PARAMNOTIFICACAO"

Public Const PER_A7MONITORACAO              As String = "PER_A7MONITORACAO"
Public Const PER_A7MONITORACAOMENSAGEM      As String = "PER_A7MONITORACAOMENSAGEM"
Public Const PER_A7MONITORACAOMENSAGEMREHEITADA      As String = "A7MONITORACAOMENSAGEMREHEITADA"

'----------
' *************************  SpreadSheet Settings *************************

' Action property settings
Public Const SS_ACTION_ACTIVE_CELL = 0
Public Const SS_ACTION_GOTO_CELL = 1
Public Const SS_ACTION_SELECT_BLOCK = 2
Public Const SS_ACTION_CLEAR = 3
Public Const SS_ACTION_DELETE_COL = 4
Public Const SS_ACTION_DELETE_ROW = 5
Public Const SS_ACTION_INSERT_COL = 6
Public Const SS_ACTION_INSERT_ROW = 7
Public Const SS_ACTION_RECALC = 11
Public Const SS_ACTION_CLEAR_TEXT = 12
Public Const SS_ACTION_PRINT = 13
Public Const SS_ACTION_DESELECT_BLOCK = 14
Public Const SS_ACTION_DSAVE = 15
Public Const SS_ACTION_SET_CELL_BORDER = 16
Public Const SS_ACTION_ADD_MULTISELBLOCK = 17
Public Const SS_ACTION_GET_MULTI_SELECTION = 18
Public Const SS_ACTION_COPY_RANGE = 19
Public Const SS_ACTION_MOVE_RANGE = 20
Public Const SS_ACTION_SWAP_RANGE = 21
Public Const SS_ACTION_CLIPBOARD_COPY = 22
Public Const SS_ACTION_CLIPBOARD_CUT = 23
Public Const SS_ACTION_CLIPBOARD_PASTE = 24
Public Const SS_ACTION_SORT = 25
Public Const SS_ACTION_COMBO_CLEAR = 26
Public Const SS_ACTION_COMBO_REMOVE = 27
Public Const SS_ACTION_RESET = 28
Public Const SS_ACTION_SEL_MODE_CLEAR = 29
Public Const SS_ACTION_VMODE_REFRESH = 30
Public Const SS_ACTION_SMARTPRINT = 32

' Appearance property settings
Public Const SS_APPEARANCE_FLAT = 0
Public Const SS_APPEARANCE_3D = 1
Public Const SS_APPEARANCE_3DWITHBORDER = 2

' BackColorStyle property settings
Public Const SS_BACKCOLORSTYLE_OVERGRID = 0
Public Const SS_BACKCOLORSTYLE_UNDERGRID = 1
Public Const SS_BACKCOLORSTYLE_OVERHORZGRIDONLY = 2
Public Const SS_BACKCOLORSTYLE_OVERVERTGRIDONLY = 3

' ButtonDrawMode property settings
Public Const SS_BDM_ALWAYS = 0
Public Const SS_BDM_CURRENT_CELL = 1
Public Const SS_BDM_CURRENT_COLUMN = 2
Public Const SS_BDM_CURRENT_ROW = 4
Public Const SS_BDM_ALWAYS_BUTTON = 8
Public Const SS_BDM_ALWAYS_COMBO = 16

' CellBorderStyle property settings
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_DASH = 2
Public Const SS_BORDER_STYLE_DOT = 3
Public Const SS_BORDER_STYLE_DASH_DOT = 4
Public Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Public Const SS_BORDER_STYLE_BLANK = 6
Public Const SS_BORDER_STYLE_FINE_SOLID = 11
Public Const SS_BORDER_STYLE_FINE_DASH = 12
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' CellBorderType property settings
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_OUTLINE = 16
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8

' CellType property settings
Public Const SS_CELL_TYPE_DATE = 0
Public Const SS_CELL_TYPE_EDIT = 1
Public Const SS_CELL_TYPE_FLOAT = 2
Public Const SS_CELL_TYPE_INTEGER = 3
Public Const SS_CELL_TYPE_PIC = 4
Public Const SS_CELL_TYPE_STATIC_TEXT = 5
Public Const SS_CELL_TYPE_TIME = 6
Public Const SS_CELL_TYPE_BUTTON = 7
Public Const SS_CELL_TYPE_COMBOBOX = 8
Public Const SS_CELL_TYPE_PICTURE = 9
Public Const SS_CELL_TYPE_CHECKBOX = 10
Public Const SS_CELL_TYPE_OWNER_DRAWN = 11

' ClipboardOptions property settings
Public Const SS_CLIP_NOHEADERS = 0
Public Const SS_CLIP_COPYROWHEADERS = 1
Public Const SS_CLIP_PASTEROWHEADERS = 2
Public Const SS_CLIP_COPYCOLHEADERS = 4
Public Const SS_CLIP_PASTECOLHEADERS = 8
Public Const SS_CLIP_COPYPASTEALLHEADERS = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Public Const SS_HEADER_BLANK = 0
Public Const SS_HEADER_NUMBERS = 1
Public Const SS_HEADER_LETTERS = 2

' CursorStyle property settings
Public Const SS_CURSOR_STYLE_USER_DEFINED = 0
Public Const SS_CURSOR_STYLE_DEFAULT = 1
Public Const SS_CURSOR_STYLE_ARROW = 2
Public Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Public Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Public Const SS_CURSOR_TYPE_DEFAULT = 0
Public Const SS_CURSOR_TYPE_COLRESIZE = 1
Public Const SS_CURSOR_TYPE_ROWRESIZE = 2
Public Const SS_CURSOR_TYPE_BUTTON = 3
Public Const SS_CURSOR_TYPE_GRAYAREA = 4
Public Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Public Const SS_CURSOR_TYPE_COLHEADER = 6
Public Const SS_CURSOR_TYPE_ROWHEADER = 7
Public Const SS_CURSOR_TYPE_DRAGDROPAREA = 8
Public Const SS_CURSOR_TYPE_DRAGDROP = 9

' DAutoSize property settings
Public Const SS_AUTOSIZE_NO = 0
Public Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Public Const SS_AUTOSIZE_BEST_GUESS = 2

' EditEnterAction property settings
Public Const SS_CELL_EDITMODE_EXIT_NONE = 0
Public Const SS_CELL_EDITMODE_EXIT_UP = 1
Public Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Public Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Public Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Public Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Public Const SS_CELL_EDITMODE_EXIT_SAME = 7
Public Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' OperationMode property settings
Public Const SS_OP_MODE_NORMAL = 0
Public Const SS_OP_MODE_READONLY = 1
Public Const SS_OP_MODE_ROWMODE = 2
Public Const SS_OP_MODE_SINGLE_SELECT = 3
Public Const SS_OP_MODE_MULTI_SELECT = 4
Public Const SS_OP_MODE_EXT_SELECT = 5

' Position property settings
Public Const SS_POSITION_UPPER_LEFT = 0
Public Const SS_POSITION_UPPER_CENTER = 1
Public Const SS_POSITION_UPPER_RIGHT = 2
Public Const SS_POSITION_CENTER_LEFT = 3
Public Const SS_POSITION_CENTER_CENTER = 4
Public Const SS_POSITION_CENTER_RIGHT = 5
Public Const SS_POSITION_BOTTOM_LEFT = 6
Public Const SS_POSITION_BOTTOM_CENTER = 7
Public Const SS_POSITION_BOTTOM_RIGHT = 8

' PrintOrientation property settings
Public Const SS_PRINTORIENT_DEFAULT = 0
Public Const SS_PRINTORIENT_PORTRAIT = 1
Public Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Public Const SS_PRINT_ALL = 0
Public Const SS_PRINT_CELL_RANGE = 1
Public Const SS_PRINT_CURRENT_PAGE = 2
Public Const SS_PRINT_PAGE_RANGE = 3

' ScrollBars property settings
Public Const SS_SCROLLBAR_NONE = 0
Public Const SS_SCROLLBAR_H_ONLY = 1
Public Const SS_SCROLLBAR_V_ONLY = 2
Public Const SS_SCROLLBAR_BOTH = 3

' ScrollBarTrack property settings
Public Const SS_SCROLLBARTRACK_OFF = 0
Public Const SS_SCROLLBARTRACK_VERTICAL = 1
Public Const SS_SCROLLBARTRACK_HORIZONTAL = 2
Public Const SS_SCROLLBARTRACK_BOTH = 3

' SelBackColor property settings
Public Const SPREAD_COLOR_NONE = &H8000000B

' SelectBlockOptions property settings
Public Const SS_SELBLOCKOPT_COLS = 1
Public Const SS_SELBLOCKOPT_ROWS = 2
Public Const SS_SELBLOCKOPT_BLOCKS = 4
Public Const SS_SELBLOCKOPT_ALL = 8

' SortBy property settings
Public Const SS_SORT_BY_ROW = 0
Public Const SS_SORT_BY_COL = 1

' SortKeyOrder property settings
Public Const SS_SORT_ORDER_NONE = 0
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_SORT_ORDER_DESCENDING = 2

' TextTip property settings
Public Const SS_TEXTTIP_OFF = 0
Public Const SS_TEXTTIP_FIXED = 1
Public Const SS_TEXTTIP_FLOATING = 2
Public Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Public Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

' TypeButtonAlign property settings
Public Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Public Const SS_CELL_BUTTON_ALIGN_TOP = 1
Public Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Public Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' TypeButtonType property settings
Public Const SS_CELL_BUTTON_NORMAL = 0
Public Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeCheckTextAlign property settings
Public Const SS_CHECKBOX_TEXT_LEFT = 0
Public Const SS_CHECKBOX_TEXT_RIGHT = 1

' TypeCheckType property settings
Public Const SS_CHECKBOX_NORMAL = 0
Public Const SS_CHECKBOX_THREE_STATE = 1

' TypeDateFormat property settings
Public Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Public Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Public Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Public Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Public Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Public Const SS_CELL_EDIT_CASE_NO_CASE = 1
Public Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Public Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Public Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeHAlign property settings
Public Const SS_CELL_H_ALIGN_LEFT = 0
Public Const SS_CELL_H_ALIGN_RIGHT = 1
Public Const SS_CELL_H_ALIGN_CENTER = 2

' TypeTextAlignVert property settings
Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Public Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Public Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Public Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Public Const SS_CELL_TIME_24_HOUR_CLOCK = 1

' TypeVAlign property settings
Public Const SS_CELL_V_ALIGN_TOP = 0
Public Const SS_CELL_V_ALIGN_BOTTOM = 1
Public Const SS_CELL_V_ALIGN_VCENTER = 2

' UnitType property settings
Public Const SS_CELL_UNIT_NORMAL = 0
Public Const SS_CELL_UNIT_VGA = 1
Public Const SS_CELL_UNIT_TWIPS = 2

' UserResize property settings
Public Const SS_USER_RESIZE_COL = 1
Public Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Public Const SS_USER_RESIZE_DEFAULT = 0
Public Const SS_USER_RESIZE_ON = 1
Public Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Public Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' ActionKey method settings
Public Const SS_KBA_CLEAR = 0
Public Const SS_KBA_CURRENT = 1
Public Const SS_KBA_POPUP = 2

' AddCustomFunctionExt method Flags parameter settings
Public Const SS_CUSTFUNC_WANTCELLREF = 1
Public Const SS_CUSTFUNC_WANTRANGEREF = 2

' CFGetParamInfo method Type parameter settings
Public Const SS_VALUE_TYPE_LONG = 0
Public Const SS_VALUE_TYPE_DOUBLE = 1
Public Const SS_VALUE_TYPE_STR = 2
Public Const SS_VALUE_TYPE_CELL = 3
Public Const SS_VALUE_TYPE_RANGE = 4

' CFGetParamInfo method Status parameter settings
Public Const SS_VALUE_STATUS_OK = 0
Public Const SS_VALUE_STATUS_ERROR = 1
Public Const SS_VALUE_STATUS_EMPTY = 2

' GetRefStyle/SetRefStyle methods return values/parameter settings
Public Const SS_REFSTYLE_DEFAULT = 0
Public Const SS_REFSTYLE_A1 = 1
Public Const SS_REFSTYLE_R1C1 = 2

' PrintOptions method PageOrder parameter settings
Public Const SS_PAGEORDER_AUTO = 0
Public Const SS_PAGEORDER_DOWNTHENOVER = 1
Public Const SS_PAGEORDER_OVERTHENDOWN = 2

' TextTipFetch method MultiLine parameter settings
Public Const SS_TT_MULTILINE_SINGLE = 0
Public Const SS_TT_MULTILINE_MULTI = 1
Public Const SS_TT_MULTILINE_AUTO = 2

' *************************  PrintPreview Settings *************************

' GrayAreaMarginType property values
Public Const SPV_GRAYAREAMARGINTYPE_SCALED = 0
Public Const SPV_GRAYAREAMARGINTYPE_ACTUAL = 1

' MousePointer property values
Public Const SPV_MOUSEPOINTER_DEFAULT = 0
Public Const SPV_MOUSEPOINTER_ARROW = 1
Public Const SPV_MOUSEPOINTER_CROSS = 2
Public Const SPV_MOUSEPOINTER_I_BEAM = 3
Public Const SPV_MOUSEPOINTER_ICON = 4
Public Const SPV_MOUSEPOINTER_SIZE = 5
Public Const SPV_MOUSEPOINTER_SIZE_NE_SW = 6
Public Const SPV_MOUSEPOINTER_SIZE_N_S = 7
Public Const SPV_MOUSEPOINTER_SIZE_NW_SE = 8
Public Const SPV_MOUSEPOINTER_SIZE_W_E = 9
Public Const SPV_MOUSEPOINTER_UP_ARROW = 10
Public Const SPV_MOUSEPOINTER_HOURGLASS = 11
Public Const SPV_MOUSEPOINTER_NO_DROP = 12

' PageViewType property values
Public Const SPV_PAGEVIEWTYPE_WHOLE_PAGE = 0
Public Const SPV_PAGEVIEWTYPE_NORMAL_SIZE = 1
Public Const SPV_PAGEVIEWTYPE_PERCENTAGE = 2
Public Const SPV_PAGEVIEWTYPE_PAGE_WIDTH = 3
Public Const SPV_PAGEVIEWTYPE_PAGE_HEIGHT = 4
Public Const SPV_PAGEVIEWTYPE_MULTIPLE_PAGES = 5

' ScrollBarH property values
Public Const SPV_SCROLLBARH_SHOW = 0
Public Const SPV_SCROLLBARH_AUTO = 1
Public Const SPV_SCROLLBARH_HIDE = 2

' ScrollBarV property values
Public Const SPV_SCROLLBARV_SHOW = 0
Public Const SPV_SCROLLBARV_AUTO = 1
Public Const SPV_SCROLLBARV_HIDE = 2

' ZoomState property values
Public Const SPV_ZOOMSTATE_INDETERMINATE = 0
Public Const SPV_ZOOMSTATE_IN = 1
Public Const SPV_ZOOMSTATE_OUT = 2
Public Const SPV_ZOOMSTATE_SWITCH = 3

Public gintIndexWorksheets                  As Integer

Public gstrVersao                           As String

'------------
Public gxmlEmpresa                          As MSXML2.DOMDocument40
Public gxmlSistema                          As MSXML2.DOMDocument40
Public gxmlTipoMensagem                     As MSXML2.DOMDocument40
Public gxmlOcorrencia                       As MSXML2.DOMDocument40

Public Function fgObterDetalhesVersoes() As String

#If EnableSoap = 1 Then
    Dim objVersao           As MSSOAPLib30.SoapClient30
#Else
    Dim objVersao           As A7Miu.clsVersao
#End If

Dim xmlVersoes              As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If gstrVersao <> vbNullString Then
        fgObterDetalhesVersoes = gstrVersao
        Exit Function
    End If

    Set xmlVersoes = CreateObject("MSXML2.DOMDocument.4.0")
    fgAppendNode xmlVersoes, "", "Componentes", ""
    flAdicionaDadosVersao xmlVersoes

    Set objVersao = fgCriarObjetoMIU("A7MIU.clsVersao")
    xmlVersoes.loadXML objVersao.ObterVersoesComponentes(xmlVersoes.xml, _
                                                         vntCodErro, _
                                                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    gstrVersao = xmlVersoes.xml
    fgObterDetalhesVersoes = gstrVersao

    Set objVersao = Nothing
    Set xmlVersoes = Nothing

Exit Function
ErrorHandler:

    Set objVersao = Nothing
    Set xmlVersoes = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    fgRaiseError App.EXEName, "basA8LQS", "fgObterDetalhesVersoes", 0
End Function

Private Sub flAdicionaDadosVersao(ByRef xmlVersao As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodePropriedade                   As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    Set objDomNode = xmlVersao.createElement("Componente")
    
    Set objDomNodePropriedade = xmlVersao.createElement("Title")
    objDomNodePropriedade.Text = App.Title
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Tipo")
    objDomNodePropriedade.Text = fgObterTipoComponente
    objDomNode.appendChild objDomNodePropriedade

    Set objDomNodePropriedade = xmlVersao.createElement("Major")
    objDomNodePropriedade.Text = App.Major
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Minor")
    objDomNodePropriedade.Text = App.Minor
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Revision")
    objDomNodePropriedade.Text = App.Revision
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("FileDescription")
    objDomNodePropriedade.Text = App.FileDescription
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Date")
    objDomNodePropriedade.Text = flDataComponente
    objDomNode.appendChild objDomNodePropriedade
    
    xmlVersao.documentElement.appendChild objDomNode
    
    Set objDomNode = Nothing
    Set objDomNodePropriedade = Nothing
    
Exit Sub
ErrorHandler:

    Set objDomNode = Nothing
    Set objDomNodePropriedade = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "flAdicionaDadosVersao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Function flDataComponente() As String

    On Error GoTo ErrorHandler
    
    flDataComponente = fgDtHr_To_Xml(FileDateTime(App.Path & "\" & App.EXEName & ".exe"))

Exit Function
ErrorHandler:
    
    flDataComponente = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    
End Function

Public Function fgIconeApp() As StdPicture
    Set fgIconeApp = mdiBUS.Icon
End Function

Public Sub Main()

Dim strCommandLine()                        As String

On Error GoTo ErrorHandler
    
    If App.PrevInstance Then End
    
    App.OleRequestPendingTimeout = 20000
    App.OleRequestPendingMsgTitle = "A7 - BUS de Interface"
    App.OleRequestPendingMsgText = "Servidor processando. Aguarde."
    
    'Adilson - 10/11/2003 - Alteração realizada para respeitar a configuração de Data nas Opções Regionais
    gstrMascaraDataDtp = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate")
    gstrMascaraDataHoraDtp = gstrMascaraDataDtp & " " & _
                             "HH:mm:ss"
    'Adilson - 10/11/2003 - Fim
    
    strCommandLine = Split(Command(), ";")
    
    If strCommandLine(0) <> "Desenv" Then
        
        flObterConfiguracaoCommandLine
        
        If gblnRegistraTLB Then
            fgRegistraComponentes
        End If
        
        fgControlarAcesso
    Else
        gblnPerfilManutencao = True
        gstrURLWebService = strCommandLine(5)
        glngTimeOut = strCommandLine(6)
        gstrPrint = strCommandLine(7)
        gstrHelpFile = strCommandLine(8)
    End If
    
    mdiBUS.Show
    
    DoEvents
    
    Exit Sub
ErrorHandler:
    
    MsgBox "Erro ->" & Err.Description
    
    End
    
    
End Sub

Public Sub fgCenterMe(NameFrm As Form)

Dim intTop                                    As Integer

On Error Resume Next
        
    NameFrm.Left = (mdiBUS.ScaleWidth - NameFrm.Width) / 2   ' Center form horizontally.
    
    If NameFrm.MDIChild Then
       intTop = (mdiBUS.ScaleHeight - NameFrm.Height) / 2 - 640
    Else
       intTop = (mdiBUS.ScaleHeight - NameFrm.Height) / 2 + 200
    End If
    
    If intTop < 0 Then intTop = 0
    
    NameFrm.Top = intTop ' Center form vertically.
    
End Sub

Public Sub fgCarregarCombos(ByRef cboComboBox As ComboBox, _
                            ByRef xmlMapaNavegacao As MSXML2.DOMDocument40, _
                            ByRef strTagName As String, _
                            ByRef strCodigoBaseName As String, _
                            ByRef strDescricaoBaseName As String)

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    cboComboBox.Clear
    For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_" & strTagName & "/*")
        cboComboBox.AddItem objDomNode.selectSingleNode(strCodigoBaseName).Text & " - " & _
                            objDomNode.selectSingleNode(strDescricaoBaseName).Text
    Next
    
    cboComboBox.ListIndex = -1
    Set objDomNode = Nothing
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    Set objDomNode = Nothing
    Call fgCursor(False)
    
    Call fgRaiseError(App.EXEName, "basA7BusClient", "fgCarregarCombos", 0)
    
End Sub

Public Sub fgClassificarListview(List As ListView, Coluna As Integer)

Dim blnTestDate                          As Boolean
Dim blnTestNumber                        As Boolean
Dim vntConteudo                          As Variant
Dim objItemAux                           As ListItem

    blnTestDate = True
    
    For Each objItemAux In List.ListItems
        If Coluna = 1 Then
            vntConteudo = objItemAux.Text
        Else
            vntConteudo = objItemAux.SubItems(Coluna - 1)
        End If
        
        If Not IsDate(vntConteudo) Then
            blnTestDate = False
            Exit For
        End If
    Next
    
    blnTestNumber = True
    
    For Each objItemAux In List.ListItems
        If Coluna = 1 Then
            vntConteudo = objItemAux.Text
        Else
            vntConteudo = objItemAux.SubItems(Coluna - 1)
        End If
        
        If Not IsNumeric(vntConteudo) Then
            blnTestNumber = False
            Exit For
        End If
    Next
    
    With List
        If blnTestDate Or blnTestNumber Then
            .ColumnHeaders.Add , "AUX", "DATA", 0
            For Each objItemAux In List.ListItems
                If Coluna = 1 Then
                    vntConteudo = objItemAux.Text
                Else
                    vntConteudo = objItemAux.SubItems(Coluna - 1)
                End If
                
                If blnTestDate Then
                    vntConteudo = Format$(vntConteudo, "yyyymmdd hh:mm:ss")
                Else
                    vntConteudo = Format$(vntConteudo, "000000000.000000000")
                End If
                
                objItemAux.SubItems(.ColumnHeaders.Count - 1) = vntConteudo
            Next
            Coluna = .ColumnHeaders.Count
        End If
            
        If Not .Sorted Then
            .Sorted = True
            .SortKey = Coluna - 1
            .SortOrder = lvwAscending
        Else
            .SortKey = Coluna - 1
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        End If
        
        If blnTestDate Or blnTestNumber Then
            Call .ColumnHeaders.Remove(.ColumnHeaders.Item("AUX").Index)
        End If
    End With
End Sub

Public Sub fgCursor(Optional pblnStatus As Boolean = False)
    
    If pblnStatus Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Public Sub fgSearchItemCombo(ByRef pobjCombo As ComboBox, _
                             ByVal pintItem As Integer, _
                    Optional ByVal pstrText As String)

Dim lngItem                                   As Integer

On Error GoTo ErrorHandler
    
    pobjCombo.ListIndex = -1
    
    For lngItem = 0 To pobjCombo.ListCount - 1
        If Trim(pstrText) <> "" Then
            If Trim(Left(pobjCombo.List(lngItem), Len(Trim(pstrText)))) = Trim(pstrText) Then
                pobjCombo.ListIndex = lngItem
                Exit For
            End If
        Else
            If pobjCombo.ItemData(lngItem) = pintItem Then
                pobjCombo.ListIndex = lngItem
                Exit For
            End If
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, "fgSearchItemCombo"

End Sub

Public Function fgLimpaCaracterInvalido(ByVal pstrTexto As String) As Variant

Dim strRetorno                                As String

On Error GoTo ErrorHandler

    strRetorno = ""
    strRetorno = Replace(pstrTexto, Chr(CAR_APOSTROFE), "")
    strRetorno = Replace(strRetorno, Chr(CAR_ABRE_CHAVE), "")
    strRetorno = Replace(strRetorno, Chr(CAR_FECHA_CHAVE), "")
    strRetorno = Replace(strRetorno, Chr(CAR_SUBST), "")
    strRetorno = Replace(strRetorno, Chr(CAR_ASPAS), "")
    strRetorno = Replace(strRetorno, Chr(CAR_CASP1), "")
    strRetorno = Replace(strRetorno, Chr(CAR_CASP2), "")
    strRetorno = Replace(strRetorno, Chr(CAR_ASPAS1), "")
    strRetorno = Replace(strRetorno, Chr(CAR_ASPAS2), "")
    
    fgLimpaCaracterInvalido = strRetorno

    Exit Function

ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, "fgLimpaCaracterInvalido"
    
End Function

Public Function fgObterCodigoCombo(ByVal pobjCombo As ComboBox) As String

Dim intPosSeparator                          As Integer

    intPosSeparator = InStr(1, pobjCombo.Text, "-")
    If intPosSeparator = 0 Then
        fgObterCodigoCombo = "ERRO"
    Else
        fgObterCodigoCombo = Trim$(Left$(pobjCombo.Text, intPosSeparator - 1))
    End If
    
End Function

Public Function fgLockWindow(Optional ByVal plWHnd As Long)
    
    LockWindowUpdate plWHnd

End Function
Public Sub fgRaiseError(ByVal strComponente As String, _
                        ByVal strClasse As String, _
                        ByVal strMetodo As String, _
                        ByRef lngCodigoErroNegocio As Long, _
               Optional ByRef intNumeroSequencialErro As Integer = 0, _
               Optional ByVal strComplemento As String = "", _
               Optional ByRef blnGravarErro As Boolean = False)

Dim strTexto                                As String
Dim ErrNumber                               As Long
Dim ErrDescription                          As String
Dim ErrSource                               As String
Dim ErrLastDllError                         As Long
Dim ErrHelpContext                          As Long
Dim ErrHelpFile                             As String

Dim objDOMErro                              As MSXML2.DOMDocument40
Dim objElement                              As IXMLDOMElement


    If lngCodigoErroNegocio <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        ErrNumber = vbObjectError + 513 + lngCodigoErroNegocio
        ErrSource = strComponente
        ErrDescription = "Obter descrição de erro de negócio"
    Else
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        ErrSource = Err.Source
        ErrLastDllError = Err.LastDllError
        ErrHelpContext = Err.HelpContext
        ErrHelpFile = Err.HelpFile
        On Error GoTo ErrHandler
    End If
    
    Set objDOMErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDOMErro.loadXML(ErrDescription) Then
        fgAppendNode objDOMErro, "", "Erro", ""
        fgAppendNode objDOMErro, "Erro", "Grupo_ErrorInfo", ""
        fgAppendNode objDOMErro, "Grupo_ErrorInfo", "Number", ErrNumber
        fgAppendNode objDOMErro, "Grupo_ErrorInfo", "Description", ErrDescription
        fgAppendNode objDOMErro, "Erro", "Repet_Origem", ""
    End If
        
    fgAppendNode objDOMErro, "Repet_Origem", "Grupo_Origem", ""
            
    Set objElement = objDOMErro.createElement("Origem")
    objElement.Text = strComponente & " - " & strClasse & " - " & strMetodo
    
    objDOMErro.selectSingleNode("//Repet_Origem/Grupo_Origem[position()=last()]").appendChild objElement
        
    Set objElement = Nothing
    
    Set objElement = objDOMErro.createElement("Complemento")
    objElement.Text = strComplemento
    objDOMErro.selectSingleNode("//Repet_Origem/Grupo_Origem[position()=last()]").appendChild objElement
    Set objElement = Nothing
        
    strTexto = objDOMErro.xml
    
    Set objDOMErro = Nothing
    
    Err.Raise ErrNumber, strComponente & " - " & strClasse & " - " & strMetodo, strTexto

ErrHandler:
    Err.Raise Err.Number, strComponente & " - " & strClasse & " - " & strMetodo, Err.Description, ErrHelpFile, ErrHelpContext
End Sub

Public Function fgObterEstacaoTrabalho() As String

Dim strEstacao                              As String
Dim lngLen                                  As Long

On Error GoTo ErrorHandler
    
    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    strEstacao = String(lngLen, "X")
    
    GetComputerName strEstacao, lngLen
    strEstacao = Left(strEstacao, lngLen)
    
    fgObterEstacaoTrabalho = UCase(strEstacao)

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusClient", "fgObterEstacaoTrabalho", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Sub fgRegistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
        
    strCLIREG32 = App.Path & "\CLIREG32.EXE"
        
    If gblnRegistraTLB Then
        'Registra os novos componentes
        strArquivo = App.Path & "\A7MIU"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -d -nologo -q -s " & gstrSource & " -l"
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA7BusClient", "fgRegistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Public Sub fgDesregistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
    
    strCLIREG32 = App.Path & "\CliReg32.Exe"
    
    'Caso tenha registrado os componentes, fuma tudo
    If gblnRegistraTLB Then
        
        'Desregistra os componentes
        strArquivo = App.Path & "\A7MIU"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -u -d -nologo -q -l"
            
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA7BusClient", "fgDesregistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
        
End Sub

Public Function fgCriarObjetoMIU(ByVal pstrNomeClasse As String) As Object

Dim strWSDL                                 As String
Dim strServico                              As String
Dim strPorta                                As String

Dim objSoapClient                           As MSSOAPLib30.SoapClient30

On Error GoTo ErrorHandler

    #If EnableSoap = 1 Then
        strServico = UCase$(Split(pstrNomeClasse, ".")(0))
        strWSDL = gstrURLWebService & "/" & strServico & ".WSDL"
        strPorta = Split(pstrNomeClasse, ".")(1) & "SoapPort"
        
        Set objSoapClient = CreateObject("MSSOAP.SoapClient30")
        Call objSoapClient.MSSoapInit(strWSDL, strServico, strPorta)
        
        objSoapClient.ConnectorProperty("Timeout") = glngTimeOut * 1000
        
        Set fgCriarObjetoMIU = objSoapClient
    #Else
        Set fgCriarObjetoMIU = CreateObject(pstrNomeClasse)
    #End If

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, "basA8LQS", "fgCriarObjetoMIU", 0
   
End Function

Private Sub flObterConfiguracaoCommandLine()

Dim strCommandLine                           As String
Dim vntParametros                            As Variant

On Error GoTo ErrorHandler
    
    strCommandLine = Command()
    
    vntParametros = Split(strCommandLine, ";")
    
    gstrAmbiente = vntParametros(LBound(vntParametros))
    gstrSource = vntParametros(LBound(vntParametros) + 1)
    gstrUsuario = vntParametros(LBound(vntParametros) + 2)
    gblnAcessoOnLine = (UCase(vntParametros(LBound(vntParametros) + 3)) = "ON")
    #If EnableSoap = 1 Then
        gblnRegistraTLB = False
    #Else
        gblnRegistraTLB = vntParametros(LBound(vntParametros) + 4)
    #End If
    gstrURLWebService = vntParametros(LBound(vntParametros) + 5)
    glngTimeOut = vntParametros(LBound(vntParametros) + 6)
    gstrPrint = vntParametros(LBound(vntParametros) + 7)
    gstrHelpFile = vntParametros(LBound(vntParametros) + 8)
    
    Exit Sub
ErrorHandler:
    
    Err.Raise vbObjectError + 266, "strParametros", "Parâmetros Inválidos- Command Line"
    
End Sub

Public Sub fgObterInformacaoAlerta()

#If EnableSoap = 1 Then
    Dim objAlerta                           As MSSOAPLib30.SoapClient30
#Else
    Dim objAlerta                           As A7Miu.clsAlerta
#End If

Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String
Dim objListItem                             As MSComctlLib.ListItem
Dim strAlerta                               As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    Set objAlerta = fgCriarObjetoMIU("A7Miu.clsAlerta")

    strxmlAlerta = objAlerta.ObterInformacaoAlerta(vntCodErro, vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objAlerta = Nothing

    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then Exit Sub

    If xmlAlerta.loadXML(strxmlAlerta) Then
        frmAlerta.strxmlInfoAlerta = strxmlAlerta

        For Each xmlNode In xmlAlerta.selectNodes("Alerta/Repet_Alerta/*")
            Set objListItem = frmAlerta.lstMensagens.ListItems.Add(, , "")

            objListItem.SubItems(1) = xmlNode.selectSingleNode("DE_OCOR").Text
            objListItem.SubItems(2) = xmlNode.selectSingleNode("DH_OCOR").Text
            objListItem.SubItems(3) = xmlNode.selectSingleNode("DE_FONT_ERRO").Text
            objListItem.SubItems(4) = xmlNode.selectSingleNode("DE_ERRO").Text
            objListItem.Tag = xmlNode.xml

            strAlerta = strAlerta & ""
        Next

        Set xmlAlerta = Nothing
        Set objAlerta = Nothing

        If Trim(strAlerta) = "" Then Exit Sub

        frmAlerta.Show
        frmAlerta.SetFocus
    End If

    fgCursor

    Exit Sub

ErrorHandler:
    fgCursor

    Set xmlAlerta = Nothing
    Set objAlerta = Nothing

    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

End Sub

Public Function fgMaiorData(ByVal pdatData1 As Date, ByVal pdatData2 As Date) As Date
    fgMaiorData = IIf(pdatData1 > pdatData2, pdatData1, pdatData2)
End Function

Public Sub fgExportaExcel(ByVal pForm As Form, _
                 Optional ByVal pvControle As Variant)

Dim pControle                               As Control
Dim objExcel                                As Excel.Application
Dim blnPrimeiroGrid                         As Boolean

    On Error GoTo ErrorHandler
    
    fgCursor True
    
    Set objExcel = CreateObject("Excel.Application")
    gintIndexWorksheets = 1
    objExcel.Workbooks.Add
    blnPrimeiroGrid = True
    If Not IsMissing(pvControle) Then
        objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
        If TypeOf pvControle Is MSFlexGrid Then
            If pControle.Rows > pControle.FixedRows Then
                Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
            End If
        ElseIf TypeOf pvControle Is ListView Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        ElseIf TypeOf pvControle Is vaSpread Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        End If
        
        blnPrimeiroGrid = False
    Else
        For Each pControle In pForm.Controls
            If TypeOf pControle Is MSFlexGrid Then
                If pControle.Rows > pControle.FixedRows Then
                    If objExcel.Worksheets.Count < gintIndexWorksheets Then
                       objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                    End If
                    objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
                    Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                    gintIndexWorksheets = gintIndexWorksheets + 1
                    blnPrimeiroGrid = False
                End If
            ElseIf TypeOf pControle Is ListView Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            ElseIf TypeOf pControle Is vaSpread Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            End If
        Next
    End If
    
    If blnPrimeiroGrid = True Then
        MsgBox "Não existem dados à serem exportados para o Excel.", vbInformation, "Atenção"
    Else
        objExcel.Visible = True
    End If
    
    Set objExcel = Nothing

    fgCursor

    Exit Sub

ErrorHandler:
    fgCursor
    Set objExcel = Nothing
    mdiBUS.uctLogErros.MostrarErros Err, "basA7"
End Sub

Public Sub fgExportaPDF(ByVal pForm As Form, _
                 Optional ByVal pvControle As Variant)

Dim pControle                               As Control
Dim objExcel                                As Excel.Application
Dim blnPrimeiroGrid                         As Boolean

    On Error GoTo ErrorHandler
    
    fgCursor True
    
    Set objExcel = CreateObject("Excel.Application")
    gintIndexWorksheets = 1
    objExcel.Workbooks.Add
    blnPrimeiroGrid = True
    If Not IsMissing(pvControle) Then
        objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
        If TypeOf pvControle Is MSFlexGrid Then
            If pControle.Rows > pControle.FixedRows Then
                Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
            End If
        ElseIf TypeOf pvControle Is ListView Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        ElseIf TypeOf pvControle Is vaSpread Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        End If
        
        blnPrimeiroGrid = False
    Else
        For Each pControle In pForm.Controls
            If TypeOf pControle Is MSFlexGrid Then
                If pControle.Rows > pControle.FixedRows Then
                    If objExcel.Worksheets.Count < gintIndexWorksheets Then
                       objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                    End If
                    objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
                    Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                    
                    'Para cada Sheet gera um arquivo PDF
                    objExcel.Worksheets(gintIndexWorksheets).Select
                    With objExcel.Worksheets(gintIndexWorksheets).PageSetup
                        .Orientation = xlLandscape
                        .Zoom = 70
                        .LeftMargin = Application.InchesToPoints(0.393700787401575)
                        .RightMargin = Application.InchesToPoints(0.393700787401575)
                        .TopMargin = Application.InchesToPoints(0.393700787401575)
                        .BottomMargin = Application.InchesToPoints(0.393700787401575)
                        .HeaderMargin = Application.InchesToPoints(0.47244094488189)
                        .FooterMargin = Application.InchesToPoints(0.47244094488189)
                    End With
                    
                    objExcel.Worksheets(gintIndexWorksheets).Cells.Select
                    objExcel.Worksheets(gintIndexWorksheets).Cells.EntireColumn.AutoFit
                    
                    objExcel.Worksheets(gintIndexWorksheets).PrintOut ActivePrinter:=gstrPrint, Collate:=True
                    
                    gintIndexWorksheets = gintIndexWorksheets + 1
                    blnPrimeiroGrid = False
                End If
            ElseIf TypeOf pControle Is ListView Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                
                    'Para cada Sheet gera um arquivo PDF
                    objExcel.Worksheets(gintIndexWorksheets).Select
                    With objExcel.Worksheets(gintIndexWorksheets).PageSetup
                        .Orientation = xlLandscape
                        .Zoom = 70
                        .LeftMargin = Application.InchesToPoints(0.393700787401575)
                        .RightMargin = Application.InchesToPoints(0.393700787401575)
                        .TopMargin = Application.InchesToPoints(0.393700787401575)
                        .BottomMargin = Application.InchesToPoints(0.393700787401575)
                        .HeaderMargin = Application.InchesToPoints(0.47244094488189)
                        .FooterMargin = Application.InchesToPoints(0.47244094488189)
                    End With
                    
                    objExcel.Worksheets(gintIndexWorksheets).Cells.Select
                    objExcel.Worksheets(gintIndexWorksheets).Cells.EntireColumn.AutoFit
                    
                    objExcel.Worksheets(gintIndexWorksheets).PrintOut ActivePrinter:=gstrPrint, Collate:=True
                
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            ElseIf TypeOf pControle Is vaSpread Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets    '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                
                    'Para cada Sheet gera um arquivo PDF
                    objExcel.Worksheets(gintIndexWorksheets).Select
                    With objExcel.Worksheets(gintIndexWorksheets).PageSetup
                        .Orientation = xlLandscape
                        .Zoom = 70
                        .LeftMargin = Application.InchesToPoints(0.393700787401575)
                        .RightMargin = Application.InchesToPoints(0.393700787401575)
                        .TopMargin = Application.InchesToPoints(0.393700787401575)
                        .BottomMargin = Application.InchesToPoints(0.393700787401575)
                        .HeaderMargin = Application.InchesToPoints(0.47244094488189)
                        .FooterMargin = Application.InchesToPoints(0.47244094488189)
                    End With
                    
                    objExcel.Worksheets(gintIndexWorksheets).Cells.Select
                    objExcel.Worksheets(gintIndexWorksheets).Cells.EntireColumn.AutoFit
                    
                    objExcel.Worksheets(gintIndexWorksheets).PrintOut ActivePrinter:=gstrPrint, Collate:=True
                
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            End If
        Next
    End If
    
    If blnPrimeiroGrid = True Then
        MsgBox "Não existem dados à serem exportados para o PDF.", vbInformation, "Atenção"
    End If
    
    Set objExcel = Nothing

    fgCursor

    Exit Sub

ErrorHandler:
    fgCursor
    Set objExcel = Nothing
    mdiBUS.uctLogErros.MostrarErros Err, "basA7"
End Sub

Public Sub flGeraDadosExcel(ByRef pobjExcel As Excel.Application, _
                            ByVal pControle As Control, _
                            ByVal blnPrimeiroGrid As Boolean)

Dim llCol                                   As Long
Static llRow                                As Long

Dim llMaxLen                                As Long
Dim llTotalLinhas                           As Long
Dim lsRange                                 As String
Dim lsSeparadorDecimal                      As String
Dim ListItem                                As MSComctlLib.ListItem
Dim strAux                                  As String
Dim llMaxLenHeader                          As Long

On Error GoTo ErrorHandler
    
    lsSeparadorDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    
    With pobjExcel.Worksheets(gintIndexWorksheets)
        
        If TypeOf pControle Is MSFlexGrid Then
        
            pControle.ReDraw = False
        
            If blnPrimeiroGrid Then
                llTotalLinhas = 3
            Else
                llTotalLinhas = llRow + 4
            End If

            For llCol = 0 To pControle.Cols - 1
                
                For llRow = 0 To pControle.Rows - 1
                        
                        If pControle.ColWidth(llCol) <> 0 Then
                            If IsNumeric(pControle.TextMatrix(llRow, llCol)) Then
                                If Len(Trim(strAux)) <= 28 Then
                                    If InStr(1, pControle.TextMatrix(llRow, llCol), lsSeparadorDecimal) > 0 Then
                                        .Cells(llRow + llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(pControle.TextMatrix(llRow, llCol))
                                    Else
                                        .Cells(llRow + llTotalLinhas, llCol + 1) = pControle.TextMatrix(llRow, llCol)
                                    End If
                                    
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                    
                                    If InStr(1, pControle.TextMatrix(llRow, llCol), lsSeparadorDecimal) > 0 Then
                                        .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                                    Else
                                        .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0_);[Red](#,##0)"
                                    End If
                                Else
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "@"
                                    .Cells(llTotalLinhas, llCol + 1) = CVar(pControle.TextMatrix(llRow, llCol))
                                End If
                                
                            ElseIf IsDate(pControle.TextMatrix(llRow, llCol)) Then
                                If IsTime(pControle.TextMatrix(llRow, llCol)) Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "HH:MM"
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = Format(pControle.TextMatrix(llRow, llCol), "HH:MM")
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                ElseIf Hour(pControle.TextMatrix(llRow, llCol)) <> 0 Or Minute(pControle.TextMatrix(llRow, llCol)) <> 0 Or Second(pControle.TextMatrix(llRow, llCol)) <> 0 Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.TextMatrix(llRow, llCol)
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                Else
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy"
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.TextMatrix(llRow, llCol)
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                End If
                            Else
                                .Cells(llRow + llTotalLinhas, llCol + 1) = Trim(pControle.TextMatrix(llRow, llCol))
                            End If
            
                            If blnPrimeiroGrid Then
                            
                                pControle.Col = llCol
                                pControle.Row = llRow
                                
                                .Cells(llRow + llTotalLinhas, llCol + 1).Font.Bold = pControle.CellFontBold
                                .Cells(llRow + llTotalLinhas, llCol + 1).Font.Color = pControle.CellFontBold
                                .Cells(llRow + llTotalLinhas, llCol + 1).VerticalAlignment = xlBottom
                                
                                Select Case pControle.CellAlignment
                                    Case flexAlignCenterCenter
                                        .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlCenter
                                    Case flexAlignLeftCenter
                                        .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlLeft
                                    Case flexAlignRightCenter
                                        .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlRight
                                End Select
                                
                                If llMaxLen < Len(Trim(pControle.TextMatrix(llRow, llCol))) Then
                                    llMaxLen = Len(Trim(pControle.TextMatrix(llRow, llCol)))
                                End If
                            End If
                                           
                            If llRow <= pControle.FixedRows - 1 Then
                                .Cells(llRow + llTotalLinhas, llCol + 1).Font.Bold = True
                            End If
                               
                        End If
                Next
                
                If blnPrimeiroGrid Then
                    .Cells(llTotalLinhas, llCol + 1).ColumnWidth = llMaxLen
                    llMaxLen = 0
                End If
            
            Next
            
            If .Cells(1, 1).ColumnWidth > 0 Then
                .Cells(1, 1) = "" 'flObterNomeEmpresa
                .Cells(1, 1).Font.Bold = True
            Else
                .Cells(1, 2) = "" 'flObterNomeEmpresa
                .Cells(1, 2).Font.Bold = True
            End If

            
            pControle.ReDraw = True
        
        ElseIf TypeOf pControle Is ListView Then
        
            For llCol = 0 To pControle.ColumnHeaders.Count - 1
                
                'Se tamanho da coluna do listview = 0, o tamanho da coluna do excel deve ser 0
                If Val(pControle.ColumnHeaders(llCol + 1).Width) = 0 Then
                    .Cells(llTotalLinhas, llCol + 1).ColumnWidth = 0
                Else
                    
                    llTotalLinhas = 3
                    llMaxLenHeader = Len(pControle.ColumnHeaders(llCol + 1).Text)
                    
                    .Cells(3, llCol + 1) = pControle.ColumnHeaders(llCol + 1).Text
                    .Cells(3, llCol + 1).Font.Bold = True
                    .Cells(3, llCol + 1).EntireColumn.AutoFit
                    
                    Select Case pControle.ColumnHeaders(llCol + 1).Alignment
                        Case lvwColumnCenter
                            .Cells(1, llCol + 1).HorizontalAlignment = xlCenter
                        Case lvwColumnLeft
                            .Cells(1, llCol + 1).HorizontalAlignment = xlLeft
                        Case lvwColumnRight
                            .Cells(1, llCol + 1).HorizontalAlignment = xlRight
                    End Select
                    
                    For Each ListItem In pControle.ListItems
                        
                        llTotalLinhas = llTotalLinhas + 1
                        
                        If llCol = 0 Then
                            strAux = ListItem.Text
                        Else
                            strAux = ListItem.SubItems(llCol)
                        End If
                        
                        If IsNumeric(strAux) Then
                        
                            If Len(Trim(strAux)) <= 28 Then
                                If InStr(1, strAux, lsSeparadorDecimal) > 0 Then
                                    .Cells(llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(strAux)
                                Else
                                    .Cells(llTotalLinhas, llCol + 1) = strAux
                                End If
                                                            
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                                
                                If InStr(1, strAux, lsSeparadorDecimal) > 0 Then
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                                Else
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "#,##0_);[Red](#,##0)"
                                End If
                            
                            Else
                                .Cells(llTotalLinhas, llCol + 1).NumberFormat = "@"
                                .Cells(llTotalLinhas, llCol + 1) = CVar(strAux)
                            End If
                            
                        ElseIf IsDate(strAux) Then
                            
                            If Len(strAux) < 6 Then
                                .Cells(llTotalLinhas, llCol + 1) = strAux
                            ElseIf Hour(strAux) <> 0 Or Minute(strAux) <> 0 Or Second(strAux) <> 0 Then
                                If Len(strAux) > 9 Then
                                    .Cells(llTotalLinhas, llCol + 1) = " " & Format(strAux, "dd/mm/yyyy hh:mm:ss")
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                Else
                                    .Cells(llTotalLinhas, llCol + 1) = " " & Format(strAux, "hh:mm:ss")
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "hh:mm:ss"
                                End If
                            Else
                                .Cells(llTotalLinhas, llCol + 1) = " " & Format(strAux, "dd/mm/yyyy")
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                                .Cells(llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy"
                            End If
                            
                        Else
                            
                            .Cells(llTotalLinhas, llCol + 1) = Trim(strAux)
                        
                        End If
    
                        Select Case pControle.ColumnHeaders(llCol + 1).Alignment
                            Case lvwColumnCenter
                                .Cells(llTotalLinhas, llCol + 1).HorizontalAlignment = xlCenter
                            Case lvwColumnLeft
                                .Cells(llTotalLinhas, llCol + 1).HorizontalAlignment = xlLeft
                            Case lvwColumnRight
                                .Cells(llTotalLinhas, llCol + 1).HorizontalAlignment = xlRight
                        End Select
                            
                        If llMaxLen < Len(strAux) Then
                            llMaxLen = Len(strAux)
                        End If
                                       
                    Next ListItem
                    
                    If llMaxLen > 0 And llMaxLen > llMaxLenHeader Then
                        .Cells(llTotalLinhas, llCol + 1).ColumnWidth = llMaxLen
                        llMaxLen = 0
                    End If
                    
                End If
                
            Next llCol
    
            .Cells(1, 1) = "" 'flObterNomeEmpresa
            .Cells(1, 1).Font.Bold = True

        ElseIf TypeOf pControle Is vaSpread Then
        
            If blnPrimeiroGrid Then
                llTotalLinhas = 3
            Else
                llTotalLinhas = llRow + 4
            End If
            
            pControle.ReDraw = False

            For llCol = 1 To pControle.MaxCols
                
                For llRow = 0 To pControle.MaxRows
                
                    pControle.Row = llRow
                    pControle.Col = llCol
                        
                    If pControle.ColWidth(llCol) <> 0 Then
                        
                        If IsNumeric(pControle.Text) Then
                            
                            If Len(Trim(strAux)) <= 28 Then
                                If InStr(1, pControle.Text, lsSeparadorDecimal) > 0 Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(pControle.Text)
                                Else
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = pControle.Text
                                End If
                                
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                                
                                If InStr(1, pControle.Text, lsSeparadorDecimal) > 0 Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                                Else
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0_);[Red](#,##0)"
                                End If
                            Else
                                .Cells(llTotalLinhas, llCol + 1).NumberFormat = "@"
                                .Cells(llTotalLinhas, llCol + 1) = CVar(pControle.Text)
                            End If
                            
                        ElseIf IsDate(pControle.Text) Then
                            If IsTime(pControle.Text) Then
                                .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "HH:MM"
                                .Cells(llRow + llTotalLinhas, llCol + 1) = Format(pControle.Text, "HH:MM")
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                            ElseIf Hour(pControle.Text) <> 0 Or Minute(pControle.Text) <> 0 Or Second(pControle.Text) <> 0 Then
                                .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.Text
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                            Else
                                .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy"
                                .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.Text
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                            End If
                        Else
                            .Cells(llRow + llTotalLinhas, llCol + 1) = Trim(pControle.Text)
                        End If
        
                        If blnPrimeiroGrid Then
                            .Cells(llRow + llTotalLinhas, llCol + 1).Font.Bold = pControle.FontBold
                            .Cells(llRow + llTotalLinhas, llCol + 1).Font.Color = vbAutomatic
                            .Cells(llRow + llTotalLinhas, llCol + 1).VerticalAlignment = xlBottom
                            
                            Select Case pControle.TypeHAlign
                                Case TypeHAlignCenter
                                    .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlCenter
                                Case TypeHAlignLeft
                                    .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlLeft
                                Case TypeHAlignRight
                                    .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlRight
                            End Select
                            
                            If llMaxLen < Len(Trim(pControle.Text)) Then
                                llMaxLen = Len(Trim(pControle.Text))
                            End If
                        End If
                           
                    End If
                Next
                
                If blnPrimeiroGrid Then
                    .Cells(llTotalLinhas, llCol + 1).ColumnWidth = llMaxLen
                    llMaxLen = 0
                End If
            
            Next
            
            If .Cells(1, 1).ColumnWidth > 0 Then
                .Cells(1, 1) = "" 'flObterNomeEmpresa
                .Cells(1, 1).Font.Bold = True
            Else
                .Cells(1, 2) = "" 'flObterNomeEmpresa
                .Cells(1, 2).Font.Bold = True
            End If
            
            pControle.ReDraw = True

        End If
        
    End With

Exit Sub
ErrorHandler:
    pControle.ReDraw = True
    mdiBUS.uctLogErros.MostrarErros Err, "basA7"
End Sub

Public Function fgControlarAcesso()

#If EnableSoap = 1 Then
    Dim objPerfil                           As MSSOAPLib30.SoapClient30
#Else
    Dim objPerfil                           As A6A7A8Miu.clsPerfil
#End If

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String
Dim lngCont                                 As Long
Dim objControl                              As Control
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    For Each objControl In mdiBUS.Controls
        If TypeName(objControl) = "Menu" Then
            If objControl.Caption <> "-" Then
                objControl.Enabled = False
            End If
        End If
    Next

    Set objPerfil = fgCriarObjetoMIU("A6A7A8Miu.clsPerfil")
    strControleAcesso = objPerfil.ObterControleAcesso("A7", _
                                                      vntCodErro, _
                                                      vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objPerfil = Nothing
   
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        fgErroLoadXML xmlControleAcesso, App.EXEName, "basA7", "fgControlarAcesso"
    End If
    
    For Each objControl In mdiBUS.Controls
        If TypeName(objControl) = "Menu" Then
            If objControl.Caption <> "-" Then
                Set xmlNode = xmlControleAcesso.selectSingleNode("//Grupo_Acesso[Perfil='" & UCase(objControl.Name) & "']/Perfil")
                If Not xmlNode Is Nothing Then
                    objControl.Enabled = True
                End If
            End If
        End If
    Next
    
    'Verifica se o usuário está associado ao GRUPO MANUTENÇÃO,
    'através da função << MANUT >>
    Set xmlNode = xmlControleAcesso.selectSingleNode("//Grupo_Acesso[Perfil='MANUT']/Perfil")
    
    If Not xmlNode Is Nothing Then
        gblnPerfilManutencao = True
    End If
        
    If XpaMBS Then
        
        mdiBUS.mnuCadastro.Enabled = True
        mdiBUS.mnuCadastroAtributo.Enabled = True
        mdiBUS.mnuCadastroParamComunicacaoSistema.Enabled = True
        mdiBUS.mnuCadastroParamNotificacao.Enabled = True
        mdiBUS.mnuCadastroProcOperAtiv.Enabled = True
        mdiBUS.mnuCadastroRegraTransporte.Enabled = True
        mdiBUS.mnuCadastroSep.Enabled = True
        mdiBUS.mnuCadastroSep1.Enabled = True
        mdiBUS.mnuCadastroTipoMensagem.Enabled = True
        mdiBUS.mnuCadParamGerais.Enabled = True
        mdiBUS.mnuControleAcessoUsuario.Enabled = True
        
        mdiBUS.mnuFerramentas.Enabled = True
        mdiBUS.mnuFerrImportacaoArquivoOperacoes.Enabled = True
        mdiBUS.mnuMonitoracao.Enabled = True
        mdiBUS.mnuMonitoracaoLogMensagensRejeitadas.Enabled = True
        mdiBUS.mnuMonitoracaoMensagens.Enabled = True
        mdiBUS.mnuReprocessaMensagem.Enabled = True
        mdiBUS.mnuTesteConectividade.Enabled = True
        mdiBUS.mnuLogExecucaoBatch.Enabled = True
        mdiBUS.mnuExportarPDF.Enabled = True
         
        gblnPerfilManutencao = True
    End If
    
    mdiBUS.mnuAjuda.Enabled = True
    mdiBUS.mnuAjudaManual.Enabled = True
    mdiBUS.mnuAjudaSobre.Enabled = True
    
    Set xmlControleAcesso = Nothing
    
Exit Function
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    Set objPerfil = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "basA7", "fgControlarAcesso", 0
    
End Function

Public Sub fgAcertarTamanhoColuna(ByRef plstListView As ListView)
                                
    Dim lngTotalColunas As Long
    Dim lngL As Long
    Const LVM_FIRST As Long = &H1000
    Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
    Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
    Dim pblnUltima As Boolean
    
    pblnUltima = True
    
    If plstListView.hwnd = 0& Or plstListView.View <> lvwReport Then
        Exit Sub
    End If
    
    lngTotalColunas = plstListView.ColumnHeaders.Count

    'Loop and correct all other column widths
    If lngTotalColunas <= 0 Then Exit Sub

    For lngL = 0 To lngTotalColunas - 1
        If (lngL = lngTotalColunas - 1) And Not pblnUltima Then plstListView.ColumnHeaders.Add , , vbNullString, 0
        SendMessageLong plstListView.hwnd, LVM_SETCOLUMNWIDTH, lngL, LVSCW_AUTOSIZE_USEHEADER
        If (lngL = lngTotalColunas - 1) And Not pblnUltima Then plstListView.ColumnHeaders.Remove (lngTotalColunas + 1)
    Next lngL

End Sub

Public Sub fgSizeOfListCol(ByRef plstListView As ListView)
Dim objColumnHeader                         As ColumnHeader

    For Each objColumnHeader In plstListView.ColumnHeaders
        Debug.Print objColumnHeader.Index & " - " & _
                    objColumnHeader.Width
    Next
    

End Sub

Public Sub fgObterIntervaloVerificacao()

#If EnableSoap = 1 Then
    Dim objMonitoracao      As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao      As A7Miu.clsMonitoracao
#End If

Dim xmlIntervalo            As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set xmlIntervalo = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    
    If xmlIntervalo.loadXML(objMonitoracao.ObterIntervaloVerificaServer(vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        If xmlIntervalo.selectSingleNode("@HoraInicio") Is Nothing Then
            strHoraInicioVerificacao = Trim$(xmlIntervalo.selectSingleNode("//@HoraInicio").Text)
        End If
        
        If xmlIntervalo.selectSingleNode("@HoraFim") Is Nothing Then
            strHoraFimVerificacao = Trim$(xmlIntervalo.selectSingleNode("//@HoraFim").Text)
        End If

    End If

    Set xmlIntervalo = Nothing
    Set objMonitoracao = Nothing

Exit Sub
ErrorHandler:
    
    Set xmlIntervalo = Nothing
    Set objMonitoracao = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "basA7", "fgObterIntervaloVerificacao", 0

End Sub
Public Function fgVerificaJanelaVerificacao() As Boolean

Dim dtmDataHoraAtual                        As Date
Dim dtmDataHoraInicio                       As Date
Dim dtmDataHoraFim                          As Date

On Error GoTo ErrorHandler

    dtmDataHoraAtual = Now
    dtmDataHoraInicio = fgDtHrStr_To_DateTime(Format$(Date, "YYYYMMDD") & fgCompletaString(strHoraInicioVerificacao, "0", 4, True) & "00")
    dtmDataHoraFim = fgDtHrStr_To_DateTime(Format$(Date, "YYYYMMDD") & fgCompletaString(strHoraFimVerificacao, "0", 4, True) & "00")
    
    If dtmDataHoraAtual >= dtmDataHoraInicio And dtmDataHoraAtual <= dtmDataHoraFim Then
        fgVerificaJanelaVerificacao = True
    Else
        fgVerificaJanelaVerificacao = False
    End If
    
    Exit Function

ErrorHandler:
         
    fgRaiseError App.EXEName, "basA7", "fgVerificaJanelaVerificacao", 0

End Function

'Marcar/Desmarcar todos os registros
Public Sub fgMarcarDesmarcarTodas(ByVal lstListView As ListView, _
                                  ByVal plngTipoSelecao As enumTipoSelecao)

Dim lngLinha                                As Long

    On Error GoTo ErrorHandler

    For lngLinha = 1 To lstListView.ListItems.Count
        lstListView.ListItems(lngLinha).Checked = (plngTipoSelecao = enumTipoSelecao.MarcarTodas)
    Next

    Exit Sub

ErrorHandler:
   fgRaiseError App.EXEName, "basA7", "fgMarcarDesmarcarTodas", 0

End Sub

'Executar o Método Genérico 'Executar' do objeto A7MIU.clsMIU
Public Function fgMIUExecutarGenerico(ByVal pstrOperacao As String, _
                                      ByVal pstrObjeto As String, _
                                      ByVal pXMLFiltro As MSXML2.DOMDocument40) As String

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A7Miu.clsMIU
#End If

Dim xmlLeitura              As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

    On Error GoTo ErrorHandler
    
    If pXMLFiltro.xml = vbNullString Then
        Call fgAppendNode(pXMLFiltro, vbNullString, "Repeat_Filtro", vbNullString)
    End If
    
    Set objMIU = fgCriarObjetoMIU("A7MIU.clsMIU")
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_Filtros", "")
    Call fgAppendAttribute(xmlLeitura, "Grupo_Filtros", "Operacao", pstrOperacao)
    Call fgAppendAttribute(xmlLeitura, "Grupo_Filtros", "Objeto", pstrObjeto)
    Call fgAppendXML(xmlLeitura, "Grupo_Filtros", pXMLFiltro.xml)
                    
    fgMIUExecutarGenerico = objMIU.Executar(xmlLeitura.xml, _
                                            vntCodErro, _
                                            vntMensagemErro)
    
    If Not IsNull(vntCodErro) Then
        If Val(Trim$(vntCodErro)) <> 0 Then
            GoTo ErrorHandler
        End If
    End If
    
    Set objMIU = Nothing
    Set xmlLeitura = Nothing
    
    Exit Function

ErrorHandler:
    Set objMIU = Nothing
    Set xmlLeitura = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call fgRaiseError(App.EXEName, "basA7", "fgMIUExecutarGenerico", 0)
    
End Function
