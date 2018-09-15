VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubReservaD0 
   Caption         =   "Sub-reserva -  D-Zero"
   ClientHeight    =   7920
   ClientLeft      =   810
   ClientTop       =   2775
   ClientWidth     =   14040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   14040
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDetalheVeicLega 
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3795
   End
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4005
      _extentx        =   7064
      _extenty        =   661
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7560
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   635
      ButtonWidth     =   2752
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar Filtro"
            Key             =   "aplicarfiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
            Key             =   "showfiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Árvore"
            Key             =   "showtreeview"
            Object.ToolTipText     =   "Mostrar TreeView"
            ImageIndex      =   3
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Lista"
            Key             =   "showlist"
            Object.ToolTipText     =   "Mostrar Lista"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Detalhe"
            Key             =   "showdetail"
            Object.ToolTipText     =   "Mostrar Detalhe"
            ImageIndex      =   5
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvGeral 
      Height          =   3555
      Left            =   30
      TabIndex        =   4
      Top             =   390
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   6271
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgOutrosIcones"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread vasLista 
      Height          =   3645
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Movimentação em D0"
      Top             =   360
      Width           =   9375
      _Version        =   196608
      _ExtentX        =   16536
      _ExtentY        =   6429
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridSolid       =   0   'False
      MaxCols         =   2
      MaxRows         =   1
      OperationMode   =   1
      RowHeaderDisplay=   0
      SpreadDesigner  =   "frmSubReservaD0.frx":0000
      UnitType        =   2
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   30
      Top             =   3960
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
            Picture         =   "frmSubReservaD0.frx":01E8
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":02FA
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":040C
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":075E
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":0AB0
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":0E02
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":1154
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":15A6
            Key             =   "posterior"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetalhe 
      Height          =   1185
      Left            =   60
      TabIndex        =   5
      Tag             =   "Detalhamento da Movimentação do Item de Caixa"
      Top             =   6330
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   2090
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   -2147483638
      BackColorBkg    =   16777215
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
   End
   Begin MSComctlLib.ImageList imgOutrosIcones 
      Left            =   600
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":19F8
            Key             =   "itemgrupo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":1D4A
            Key             =   "selectednode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":209C
            Key             =   "itemelementar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":23EE
            Key             =   "treeminus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":2710
            Key             =   "treeplus"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaD0.frx":2A32
            Key             =   "leaf"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   -90
      MousePointer    =   7  'Size N S
      Top             =   6150
      Width           =   13320
   End
   Begin VB.Image imgDummyV 
      Height          =   3465
      Left            =   3855
      MousePointer    =   9  'Size W E
      Top             =   405
      Width           =   90
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6675
      TabIndex        =   0
      Top             =   30
      Width           =   300
   End
   Begin VB.Label lblBarra 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmSubReservaD0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a consulta da posição do caixa em D0.

Option Explicit

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1
Private WithEvents objBuscaNo               As frmBuscaNo
Attribute objBuscaNo.VB_VarHelpID = -1
Private strDocFiltros                       As String

Private blnEventByPass                      As Boolean
Private xmlItensCaixa                       As MSXML2.DOMDocument40
Private strGruposVeiculosLegais             As String
Private fblnDummyV                          As Boolean
Private fblnDummyH                          As Boolean
Private strCarregaTreeView                  As String
Private strItemCaixaGrupoVeiculoLegal       As String

Private lngLinhaFinalSpread                 As Long
Private intTipoBackOfficeUsuario            As Integer

Private Const strFuncionalidade             As String = "frmItemCaixa"
Private Const strTableComboInicial          As String = "Empresa"
Private Const intLinhaInicioMovimentacao    As Integer = 4

Private lngAlturaTableCombo                 As Long

'Linhas e Colunas do SPREAD
Private Const ROW_ABERTURA                  As Integer = 3
Private Const ROW_MOVIMENTACAO              As Integer = 4

Private Const TOT_COLUNAS_SPREAD            As Integer = 15

Private Const COL_PRINCIPAL                 As Integer = 1
Private Const COL_NIVEL_1                   As Integer = 2
Private Const COL_NIVEL_2                   As Integer = 3
Private Const COL_NIVEL_3                   As Integer = 4
Private Const COL_NIVEL_4                   As Integer = 5
Private Const COL_NIVEL_5                   As Integer = 6
Private Const COL_DESCRICAO                 As Integer = 7
Private Const COL_SEPARADOR                 As Integer = 8
Private Const COL_PREVISTO                  As Integer = 9
Private Const COL_TOTAL                     As Integer = 10
Private Const COL_REALIZADO_OU_SOLICITADO   As Integer = 11
Private Const COL_CONFIRMADO                As Integer = 12
Private Const COL_VARIACAO_OU_REAL_PREV     As Integer = 13
Private Const COL_DATA_POSICAO_OCULTA       As Integer = 14
Private Const COL_CODIGO_OCULTO             As Integer = 15

Private Const MAX_FIXED_ROWS                As Integer = 100

'Colunas do FLEXGRID
Private Const COL_SISTEMA                   As Integer = 0
Private Const COL_EMPRESA                   As Integer = 1
Private Const COL_LOCAL_LIQUIDACAO          As Integer = 2
Private Const COL_TIPO_LIQUIDACAO           As Integer = 3
Private Const COL_DESCRICAO_ATIVO           As Integer = 4
Private Const COL_CNPJ_CONTRAPARTE          As Integer = 5
Private Const COL_NOME_CONTRAPARTE          As Integer = 6
Private Const COL_ENTRADA                   As Integer = 7
Private Const COL_SAIDA                     As Integer = 8
Private Const COL_SITUACAO_MOVIMENTO        As Integer = 9
Private Const COL_DATA_MOVIMENTO            As Integer = 10
Private Const COL_DATA_RETORNO              As Integer = 11

Private Sub ctlTableCombo_AplicarFiltro(xmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    If Trim(strDocFiltros) = vbNullString Then
        ctlTableCombo.Height = lngAlturaTableCombo
        DoEvents
        
        'Se o filtro ainda não tiver sido acionado, então força o acionamento...
        Call tlbButtons_ButtonClick(tlbButtons.Buttons("showfiltro"))
        
        Call fgCursor(False)
        
        Set xmlDomFiltro = Nothing
        Set xmlDomLeitura = Nothing
        
        Exit Sub
    Else
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltro.loadXML(strDocFiltros)
        
        If Not xmlDomFiltro.selectSingleNode("//Grupo_BancoLiquidante") Is Nothing Then
            Call fgRemoveNode(xmlDomFiltro, "Grupo_BancoLiquidante")
        End If
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", _
                                        xmlDomLeitura.selectSingleNode("//BancoLiquidante").Text)
                                        
        Call flCarregarTrvGeral(xmlDomFiltro.xml)
        strDocFiltros = xmlDomFiltro.xml
    End If
    
    Call flInicializarFlxDetalhe
    Call fgCursor(False)
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - ctlTableCombo_AplicarFiltro"

End Sub

Private Sub ctlTableCombo_DropDown()

On Error GoTo ErrorHandler

    lngAlturaTableCombo = ctlTableCombo.Height
    Call ctlTableCombo.fgCarregarCombo
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - ctlTableCombo_DropDown"

End Sub

Private Sub ctlTableCombo_MouseMove()

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButton3D

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - ctlTableCombo_MouseMove"
   
End Sub

Private Sub flxDetalhe_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - flxDetalhe_MouseMove"
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        If Trim(strDocFiltros) = vbNullString Then Exit Sub
        
        Call tlbButtons_ButtonClick(tlbButtons.Buttons("refresh"))
    End If
    
    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    Me.Icon = mdiSBR.Icon
    
    Call fgCenterMe(Me)
    Call fgCursor(True)
        
    intTipoBackOfficeUsuario = fgObterTipoBackOfficeUsuario
    
    Call flInicializarVasLista
    vasLista.CursorStyle = CursorStyleArrow
    Call flInicializarFlxDetalhe
    
    ctlTableCombo.TituloCombo = strTableComboInicial
    DoEvents
    
    '>>>>> Inicialização Formulário de Busca
    Set objBuscaNo = New frmBuscaNo
    Load objBuscaNo
    objBuscaNo.Criterio = "Veículo Legal"
    Set objBuscaNo.objTreeView = trvGeral
    
    '>>>>> Inicialização Formulário de Filtro
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaD0
    Load objFiltro
    
    Call flCarregarListaItensCaixa
    Call objFiltro.fgCarregarPesquisaAnterior
    
    Me.Show
    DoEvents
    
    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next

    With Me
        .ctlTableCombo.Left = 0
        .ctlTableCombo.Top = 0
        
        .trvGeral.Left = 0
        .trvGeral.Top = .ctlTableCombo.Height
        .trvGeral.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .trvGeral.Top, .tlbButtons.Top - .trvGeral.Top) - .txtDetalheVeicLega.Height
        .trvGeral.Width = IIf(.imgDummyV.Visible, .imgDummyV.Left, .ScaleWidth)
        
        .txtDetalheVeicLega.Left = 0
        .txtDetalheVeicLega.Top = .trvGeral.Top + .trvGeral.Height
        .txtDetalheVeicLega.Width = .trvGeral.Width
        
        .tlbButtons.Top = .ScaleHeight - .tlbButtons.Height
        
        .imgDummyH.Left = 0
        .imgDummyH.Width = .ScaleWidth
        
        .flxDetalhe.Left = 0
        .flxDetalhe.Top = IIf(.imgDummyH.Visible, .imgDummyH.Top + .imgDummyH.Height, .trvGeral.Top)
        .flxDetalhe.Height = IIf(.imgDummyH.Visible, .tlbButtons.Top - .imgDummyH.Top - .imgDummyH.Height, .tlbButtons.Top - .flxDetalhe.Top)
        .flxDetalhe.Width = .imgDummyH.Width
        
        .lblBarra.Left = 0
        .lblBarra.Top = 0
        .lblBarra.Width = .ScaleWidth
        
        .lblData.Left = .ScaleWidth - .lblData.Width
        .lblData.Top = 0
        
        .imgDummyV.Top = 0
        .imgDummyV.Height = .trvGeral.Height
        
        .vasLista.Left = IIf(.imgDummyV.Visible, .imgDummyV.Left + .imgDummyV.Width, 0)
        .vasLista.Top = .trvGeral.Top
        .vasLista.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .trvGeral.Top, .tlbButtons.Top - .trvGeral.Top)
        .vasLista.Width = .ScaleWidth - .vasLista.Left
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo ErrorHandler
    
    Set frmSubReservaD0 = Nothing
    Unload objBuscaNo
    Set objBuscaNo = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - Form_Unload"
    
End Sub

Private Sub imgDummyV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyV = False
End Sub

Private Sub lblBarra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - lblBarra_MouseMove"
   
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, lsTituloTableCombo As String)
 
On Error GoTo ErrorHandler

    strDocFiltros = xmlDocFiltros
    
    Call fgCursor(True)
    
    Call flCarregarTrvGeral(strDocFiltros)
    Call flInicializarFlxDetalhe
    Call flRefreshPosicaoAtualTela(strDocFiltros)
    
    Call fgCursor(False)
    
    ctlTableCombo.TituloCombo = IIf(lsTituloTableCombo = vbNullString, strTableComboInicial, lsTituloTableCombo)
    
    tlbButtons.Buttons("aplicarfiltro").Value = tbrPressed
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - objFiltro_AplicarFiltro"
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim lsJanelas                               As String
Dim strSituacaoGrid                         As String

On Error GoTo ErrorHandler

    If tlbButtons.Buttons("showtreeview").Value = tbrPressed Then
        lsJanelas = lsJanelas & "1"
    End If
    
    If tlbButtons.Buttons("showlist").Value = tbrPressed Then
        lsJanelas = lsJanelas & "2"
    End If
    
    If tlbButtons.Buttons("showdetail").Value = tbrPressed Then
        lsJanelas = lsJanelas & "3"
    End If
    
    Call flArranjarJanelasExibicao(lsJanelas)
    
    Select Case Button.Key
        Case "showfiltro"
            Set objFiltro = New frmFiltro
            Set objFiltro.FormOwner = Me
            objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaD0
            objFiltro.Show vbModal
            
        Case "refresh"
            If Trim(strDocFiltros) = vbNullString Then Exit Sub
            
            fgCursor True
            
            'Verifica se o GRID está populado...
            If lngLinhaFinalSpread > 0 Then
                '...se sim, verifica se os dados do GRID são para o Veículo Legal...
                If Not trvGeral.SelectedItem Is Nothing Then
                    If InStr(2, trvGeral.SelectedItem.Key, "k_") > 0 Then
                        '...se sim, captura a situação dos GRIDs
                        strSituacaoGrid = flCapturaSituacaoGrid
                    End If
                End If
            End If
            
            Call flRefreshPosicaoAtualTela(strDocFiltros)
            
            'Verifica se os GRIDs devem ser reconstruídos...
            If Trim(strSituacaoGrid) <> vbNullString Then
                '...se sim, remonta o GRID principal e o DETALHE (se for o caso)
                Call flRemontaGrids(strSituacaoGrid)
            End If
            
            fgCursor
            
        Case Else
            'Não faz nada.
            
    End Select

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - tlbButtons_ButtonClick"
        
End Sub

' Arranja as janelas de exibição segundo seleção pelo usuário.

Private Sub flArranjarJanelasExibicao(ByVal pstrJanelas As String)

On Error GoTo ErrorHandler

    Select Case pstrJanelas
    Case ""
        imgDummyH.Visible = False
        imgDummyV.Visible = False
        trvGeral.Visible = False
        vasLista.Visible = False
        flxDetalhe.Visible = False
    
    Case "1"
        imgDummyH.Visible = False
        imgDummyV.Visible = False
        trvGeral.Visible = True
        vasLista.Visible = False
        flxDetalhe.Visible = False
    
    Case "2"
        imgDummyH.Visible = False
        imgDummyV.Visible = False
        trvGeral.Visible = False
        vasLista.Visible = True
        flxDetalhe.Visible = False
    
    Case "3"
        imgDummyH.Visible = False
        imgDummyV.Visible = False
        trvGeral.Visible = False
        vasLista.Visible = False
        flxDetalhe.Visible = True
    
    Case "12"
        imgDummyH.Visible = False
        imgDummyV.Visible = True
        trvGeral.Visible = True
        vasLista.Visible = True
        flxDetalhe.Visible = False
    
    Case "13"
        imgDummyH.Visible = True
        imgDummyV.Visible = False
        trvGeral.Visible = True
        vasLista.Visible = False
        flxDetalhe.Visible = True

    Case "23"
        imgDummyH.Visible = True
        imgDummyV.Visible = False
        trvGeral.Visible = False
        vasLista.Visible = True
        flxDetalhe.Visible = True

    Case "123"
        imgDummyH.Visible = True
        imgDummyV.Visible = True
        trvGeral.Visible = True
        vasLista.Visible = True
        flxDetalhe.Visible = True

    End Select
    
    txtDetalheVeicLega.Visible = trvGeral.Visible
    Call Form_Resize

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flArranjarJanelasExibicao", 0
    
End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    fblnDummyH = False
    
End Sub

Private Sub imgDummyV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    fblnDummyV = True
    
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not fblnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyH.Top = y + imgDummyH.Top

    On Error Resume Next
    
    With Me
        If .imgDummyH.Top < 1926 Then
            .imgDummyH.Top = 1926
        End If
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If
        
        .flxDetalhe.Top = .imgDummyH.Top + .imgDummyH.Height
        .flxDetalhe.Height = .tlbButtons.Top - .imgDummyH.Top - .imgDummyH.Height
        .trvGeral.Height = .imgDummyH.Top - .trvGeral.Top - .txtDetalheVeicLega.Height
        .txtDetalheVeicLega.Top = .imgDummyH.Top - .txtDetalheVeicLega.Height
        .vasLista.Height = .trvGeral.Height + .txtDetalheVeicLega.Height
    End With
    
    On Error GoTo 0
    
End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    fblnDummyH = True
    
End Sub

Private Sub imgDummyV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not fblnDummyV Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyV.Left = x + imgDummyV.Left

    On Error Resume Next
    
    With Me
        If .imgDummyV.Left < 1926 Then
            .imgDummyV.Left = 1926
        End If
        If .imgDummyV.Left > (.Width - 500) And (.Width - 500) > 0 Then
            .imgDummyV.Left = .Width - 500
        End If
        
        .trvGeral.Width = .imgDummyV.Left
        .txtDetalheVeicLega.Width = .imgDummyV.Left
        .vasLista.Left = .imgDummyV.Left + .imgDummyV.Width
        .vasLista.Width = .Width - (.imgDummyV.Left + 180)
    End With
    
    On Error GoTo 0
    
End Sub

' Carrega treeview.

Private Sub flCarregarTrvGeral(ByRef xmlDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objItemCaixa    As MSSOAPLib30.SoapClient30
#Else
    Dim objItemCaixa    As A6MIU.clsItemCaixa
#End If

Dim xmlRetorno          As MSXML2.DOMDocument40
Dim objNode             As MSComctlLib.Node
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    lblData.Caption = vbNullString
    
    Call flInicializarVasLista
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Set objItemCaixa = fgCriarObjetoMIU("A6MIU.clsItemCaixa")
    
    Call xmlRetorno.loadXML(objItemCaixa.ObterRelacaoItensCaixaGrupoVeicLegal(xmlDocFiltros, _
                                                                              False, _
                                                                              intTipoBackOfficeUsuario, _
                                                                              vntCodErro, _
                                                                              vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If Not xmlRetorno.documentElement.selectSingleNode("Repeat_ItensCaixa") Is Nothing Then
        strItemCaixaGrupoVeiculoLegal = xmlRetorno.documentElement.selectSingleNode("Repeat_ItensCaixa").xml
    Else
        strItemCaixaGrupoVeiculoLegal = vbNullString
    End If
    
    If Not xmlRetorno.documentElement.selectSingleNode("Repeat_VeiculoLegal") Is Nothing Then
        strCarregaTreeView = xmlRetorno.documentElement.selectSingleNode("Repeat_VeiculoLegal").xml
    Else
        strCarregaTreeView = vbNullString
    End If
    
    Call fgCarregarTreeViewFluxoCaixa("VeiculoLegal", _
                                      Me.trvGeral, _
                                      "CO_GRUP_VEIC_LEGA;CO_VEIC_LEGA", _
                                      "NO_GRUP_VEIC_LEGA;NO_VEIC_LEGA", _
                                      intTipoBackOfficeUsuario, _
                                      strCarregaTreeView)
                                        
    Set objItemCaixa = Nothing
    Exit Sub
    
ErrorHandler:
    Set objItemCaixa = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaD0 - flCarregarTrvGeral", 0
End Sub

' Inicializa grid para consulta de movimentação de itens de caixa em D0.

Private Sub flInicializarVasLista()

Dim lngColunas                              As Long

On Error GoTo ErrorHandler

    With Me.vasLista
        .Redraw = False

        .MaxRows = 0
        .MaxRows = 100
        .RowsFrozen = 2
        
        .MaxCols = TOT_COLUNAS_SPREAD
        .ColWidth(1) = 200
        .ColWidth(2) = 200
        .ColWidth(3) = 200
        .ColWidth(4) = 200
        .ColWidth(5) = 200
        .ColWidth(6) = 200
        .ColWidth(7) = 2000
        .ColWidth(8) = 15
        .ColsFrozen = 8
        
        'Coluna para armazenar Data da Posição e código (Grupo / Veículo Legal)
        .Col = COL_DATA_POSICAO_OCULTA
        .ColHidden = True
        .Col = COL_CODIGO_OCULTO
        .ColHidden = True
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = 8
        .Row2 = .MaxRows
        .AllowCellOverflow = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = 7
        .Row2 = .MaxRows
        .BackColorStyle = BackColorStyleOverVertGridOnly
        .BackColor = &H8000000E
        .BlockMode = False

        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .BackColor = &H8000000F
        .ForeColor = vbBlack
        .RowHeight(1) = 300
        .FontSize = 10
        .FontBold = True
        .BlockMode = False

        '>>> Seleção Principal (Veículo Legal / Item Caixa)
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = 8
        .Row2 = 2
        .CellBorderType = 16    'Contorno da seleção de células informada
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionSetCellBorder
        .BlockMode = False
        
        '>>> Previsto
        .BlockMode = True
        .Col = 9
        .Row = 1
        .Col2 = 9
        .Row2 = 2
        .CellBorderType = 16    'Contorno da seleção de células informada
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionSetCellBorder
        .BlockMode = False
        
        '>>> Realizado
        .BlockMode = True
        .Col = 10
        .Row = 1
        .Col2 = 12
        .Row2 = 1
        .CellBorderType = 16    'Contorno da seleção de células informada
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionSetCellBorder
        .BlockMode = False
        
        '>>> Variação
        .BlockMode = True
        .Col = 13
        .Row = 1
        .Col2 = 13
        .Row2 = 1
        .CellBorderType = 16    'Contorno da seleção de células informada
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionSetCellBorder
        .BlockMode = False
        
        '>>> Sub Categorias de Realizado e Variação
        For lngColunas = 10 To .MaxCols
            .BlockMode = True
            .Col = lngColunas
            .Row = 2
            .Col2 = lngColunas
            .Row2 = 2
            .CellBorderType = 16    'Contorno da seleção de células informada
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderColor = RGB(0, 0, 0)
            .Action = ActionSetCellBorder
            .BlockMode = False
        Next
        
        .BlockMode = True
        .Col = 1
        .Row = 2
        .Col2 = .MaxCols
        .Row2 = 2
        .BackColor = &H8000000F
        .BlockMode = False
        
        .BlockMode = True
        .Col = 9
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 2
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        For lngColunas = 9 To .MaxCols
            .ColWidth(lngColunas) = 1700
        Next
        
        '>>> Formata o alinhamento a direita para os valores do GRID
        .BlockMode = True
        .Col = 9
        .Row = 3
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .TypeHAlign = TypeHAlignRight
        .BlockMode = True
        
        .SetText COL_PRINCIPAL, 1, "Grupo / Veículo Legal"
        .SetText COL_PREVISTO, 1, "Previsto"
        .SetText COL_REALIZADO_OU_SOLICITADO, 1, "Realizado"
        .SetText COL_VARIACAO_OU_REAL_PREV, 1, "Variação"
        .SetText COL_TOTAL, 2, "Total"
        .SetText COL_REALIZADO_OU_SOLICITADO, 2, "Solicitado"
        .SetText COL_CONFIRMADO, 2, "Confirmado"
        .SetText COL_VARIACAO_OU_REAL_PREV, 2, "Real. - Prev."
        
        .Redraw = True
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInicializarVasLista", 0

End Sub

Private Sub trvGeral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler
    
    If Button = vbRightButton Then
        Call objBuscaNo.Move(mdiSBR.Width - mdiSBR.ScaleWidth + Me.Left + trvGeral.Left + x, mdiSBR.Height - mdiSBR.ScaleHeight + Me.Top + trvGeral.Top + y)
        objBuscaNo.Show
        blnEventByPass = True
    End If
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - trvGeral_MouseDown"
    
End Sub

Private Sub trvGeral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim objNode                                 As MSComctlLib.Node
    
On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButtonFlat
    
    Set objNode = trvGeral.HitTest(x, y)

    If Not objNode Is Nothing Then
        If objNode.children = 0 Then
            txtDetalheVeicLega.Text = Split(objNode.Key, "k_")(2) & " - " & objNode.Text
        Else
            txtDetalheVeicLega.Text = vbNullString
        End If
    Else
        txtDetalheVeicLega.Text = vbNullString
    End If

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - trvGeral_MouseMove"
   
End Sub

Private Sub trvGeral_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnEventByPass = False
    
End Sub

Private Sub trvGeral_NodeClick(ByVal Node As MSComctlLib.Node)

Dim objNode                                 As MSComctlLib.Node
    
If blnEventByPass Then Exit Sub

On Error GoTo ErrorHandler
    
    objBuscaNo.Hide
    DoEvents
    
    For Each objNode In Me.trvGeral.Nodes
        objNode.Image = "itemgrupo"
    Next
    Node.Image = "selectednode"
    
    fgCursor True
    
    If InStr(2, Node.Key, "k_") = 0 Then
        Call flCarregarListaPorGrupoVeiculoLegal(Node)
    Else
        Call flCarregarListaPorVeiculoLegal(Node)
    End If
    
    fgCursor
    Exit Sub
    
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - trvGeral_NodeClick"
    
End Sub

' Carrega lista de itens de caixa.

Private Sub flCarregarListaItensCaixa()

#If EnableSoap = 1 Then
    Dim objMIU      As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU      As A6MIU.clsMIU
#End If

Dim xmlDomNode      As IXMLDOMNode
Dim strLerTodos     As String
Dim vntCodErro      As Variant
Dim vntMensagemErro As Variant

On Error GoTo ErrorHandler

    Set xmlItensCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    
    If Not xmlItensCaixa.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.SBR, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlItensCaixa, App.EXEName, Me.Name, "flCarregarListaItensCaixa")
    End If
    
    Set xmlDomNode = xmlItensCaixa.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa")
    xmlDomNode.selectSingleNode("@Operacao").Text = "LerTodos"
    xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaSubReserva

    strLerTodos = objMIU.Executar(xmlDomNode.xml, _
                                  vntCodErro, _
                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strLerTodos <> vbNullString Then
        Call xmlItensCaixa.loadXML(strLerTodos)
    End If
    
    Set objMIU = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlItensCaixa = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarListaItensCaixa", 0

End Sub

' Carrega movimentação em D0 por grupo de veículo legal.

Private Sub flCarregarListaPorGrupoVeiculoLegal(ByVal objNodeSel As MSComctlLib.Node)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaSubReserva  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaSubReserva  As A6MIU.clsMonitoracaoSubReserva
#End If

Dim objDomNode                              As IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomAux                               As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim arrPrimKey()                            As String
Dim strNodeKey                              As String
Dim strDataAbertura                         As String
Dim intSituacaoCaixa                        As Integer

Dim lngLinhaGrid                            As Integer
Dim lngColunaGrid                           As Integer

Dim strGrupVeicLega                         As String

Dim dblTotalVeiculo                         As Double

Dim dblValorAbertura                        As Double
Dim dblValorTotal                           As Double

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    lblData.Caption = vbNullString
    
    With Me.vasLista
        .Redraw = False
        
        Call flInicializarVasLista
        Call flInicializarFlxDetalhe
        
        .SetText COL_PRINCIPAL, 1, objNodeSel.Text
        
        '>>> Captura código do Grupo Veículo Legal... -------------------------------------------------
        arrPrimKey = Split(objNodeSel.Key, "k_")
        strNodeKey = arrPrimKey(1)
        
        '... e formata XML Filtro padrão
        Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                         "GrupoVeiculoLegal", strNodeKey)
        
        If strDocFiltros <> vbNullString Then
            Set xmlDomAux = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlDomAux.loadXML(strDocFiltros)
            
            If Not xmlDomAux.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal") Is Nothing Then
                Call fgAppendXML(xmlDomFiltros, "Repeat_Filtros", xmlDomAux.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal").xml)
            End If
            
            If Not xmlDomAux.selectSingleNode("Repeat_Filtros/Grupo_BancoLiquidante") Is Nothing Then
                Call fgAppendXML(xmlDomFiltros, "Repeat_Filtros", xmlDomAux.selectSingleNode("Repeat_Filtros/Grupo_BancoLiquidante").xml)
            End If
            
            If Not xmlDomAux.selectSingleNode("Repeat_Filtros/Grupo_BackOfficePerfilGeral") Is Nothing Then
                Call fgAppendXML(xmlDomFiltros, "Repeat_Filtros", xmlDomAux.selectSingleNode("Repeat_Filtros/Grupo_BackOfficePerfilGeral").xml)
            End If
        End If
        
        Set xmlDomAux = Nothing
        '>>> -------------------------------------------------------------------------------------------
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMonitoracaoFluxoCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsMonitoracaoSubReserva")
        
        strRetLeitura = objMonitoracaoFluxoCaixaSubReserva.ObterResumoCaixaSubReserva(xmlDomFiltros.xml, _
                                                                                      vntCodErro, _
                                                                                      vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strRetLeitura <> vbNullString Then
            If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                '100 - Documento XML Inválido.
                lngCodigoErroNegocio = 100
                GoTo ErrorHandler
            End If
            
            lngLinhaGrid = ROW_ABERTURA
            For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                strDataAbertura = objDomNode.selectSingleNode("DT_CAIX_DISP").Text
                intSituacaoCaixa = Val(objDomNode.selectSingleNode("CO_SITU_CAIX").Text)
                dblValorAbertura = fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_UTLZ_ABER_CAIX").Text)
                
                .SetText COL_NIVEL_1, lngLinhaGrid, objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SetText COL_PREVISTO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_PREV").Text)

                dblValorTotal = fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_TOT_MOV").Text)
                
                .SetText COL_TOTAL, lngLinhaGrid, _
                            fgVlrXml_To_Interface(dblValorAbertura + dblValorTotal)

                .SetText COL_REALIZADO_OU_SOLICITADO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_SOLI").Text)
                .SetText COL_CONFIRMADO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_CONF").Text)
                .SetText COL_VARIACAO_OU_REAL_PREV, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_VAR_MOV").Text)
                .SetText COL_DATA_POSICAO_OCULTA, lngLinhaGrid, strDataAbertura
                .SetText COL_CODIGO_OCULTO, lngLinhaGrid, objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & "|" & _
                                                          objDomNode.selectSingleNode("SG_SIST").Text & "|" & _
                                                          intSituacaoCaixa
                                                          
                lngLinhaGrid = lngLinhaGrid + 1
            Next
            
            lngLinhaFinalSpread = lngLinhaGrid - 1
            
            Call flColorirValoresNegativos
        End If
        
        .Redraw = True
    End With
            
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaD0 - flCarregarListaPorGrupoVeiculoLegal", 0
End Sub

' Carrega movimentação em D0 por veículo legal selecionado.

Private Sub flCarregarListaPorVeiculoLegal(ByVal objNodeSel As MSComctlLib.Node)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaSubReserva  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaSubReserva  As A6MIU.clsMonitoracaoSubReserva
#End If

Dim objDomNode                              As IXMLDOMNode
Dim objDOMNodeAux                           As IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomFiltroTela                        As MSXML2.DOMDocument40
Dim xmlItemCaixaGrupoVeiculoLegal           As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim arrPrimKey()                            As String

Dim strDataPosicaoCaixaSubReserva           As String

Dim blnAchouItem                            As Boolean
Dim blnItemGenerico                         As Boolean

Dim intCodGrupoVeicLegal                    As Integer
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim intNivelItemCaixa                       As Integer

Dim lngLinhaGrid                            As Long

Dim varItemCaixaAux                         As Variant

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    lblData.Caption = vbNullString
    
    With Me.vasLista
        .Redraw = False
        
        Call flInicializarVasLista
        Call flInicializarFlxDetalhe
        
        .SetText COL_PRINCIPAL, 1, objNodeSel.Text
        
        '>>> Captura código do Grupo Veículo Legal... -------------------------------------------------
        arrPrimKey = Split(objNodeSel.Key, "k_")
        intCodGrupoVeicLegal = Val(arrPrimKey(1))
        strCodigoVeiculoLegal = arrPrimKey(2)
        strSiglaSistema = arrPrimKey(3)
        
        '... e formata XML Filtro padrão
        Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                         "VeiculoLegal", strCodigoVeiculoLegal)
        Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                         "Sistema", strSiglaSistema)
        Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                         "NivelAgrupamento", 1)
        
        If strDocFiltros <> vbNullString Then
            Set xmlDomFiltroTela = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlDomFiltroTela.loadXML(strDocFiltros)
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Data/DataIni") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "DataCaixa", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Data/DataIni").Text)
            End If
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Sistema/Sistema") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "Sistema", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Sistema/Sistema").Text)
            End If
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BancoLiquidante/BancoLiquidante") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "BancoLiquidante", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BancoLiquidante/BancoLiquidante").Text)
            End If
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BackOfficePerfilGeral/BackOfficePerfilGeral") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "BackOfficePerfilGeral", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BackOfficePerfilGeral/BackOfficePerfilGeral").Text)
            End If
            
            Set xmlDomFiltroTela = Nothing
        End If
        '>>> -------------------------------------------------------------------------------------------
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMonitoracaoFluxoCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsMonitoracaoSubReserva")
        
        strRetLeitura = objMonitoracaoFluxoCaixaSubReserva.ObterMovimentoSubReserva(xmlDomFiltros.xml, _
                                                                                    vntCodErro, _
                                                                                    vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strRetLeitura <> vbNullString Then
            If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                '100 - Documento XML Inválido.
                lngCodigoErroNegocio = 100
                GoTo ErrorHandler
            End If
            
            .SetText COL_NIVEL_1, ROW_ABERTURA, "Abertura"
            
            '>>> O valor de ABERTURA é utilizado nas colunas (PREVISTO e TOTAL REALIZADO) ---------------------------------------------------------------------------------------------
            If Not xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva") Is Nothing Then
                .SetText COL_PREVISTO, ROW_ABERTURA, _
                            fgVlrXml_To_Interface(xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva/Grupo_PosicaoCaixaSubReserva/VA_UTLZ_ABER_CAIX_SUB_RESE").Text)
    
                .SetText COL_TOTAL, ROW_ABERTURA, _
                            fgVlrXml_To_Interface(xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva/Grupo_PosicaoCaixaSubReserva/VA_UTLZ_ABER_CAIX_SUB_RESE").Text)
    
                lblData.Caption = _
                    fgDescricaoEstadoCaixa(Val(xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva/Grupo_PosicaoCaixaSubReserva/CO_SITU_CAIX_SUB_RESE_ATUAL").Text)) & _
                    " - " & _
                    fgDtXML_To_Date(xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva/Grupo_PosicaoCaixaSubReserva/DT_CAIX_SUB_RESE_ATUAL").Text)
            Else
                If Not xmlDomFiltros.selectSingleNode("//DataCaixa") Is Nothing Then
                    lblData.Caption = "Inexistente - " & fgDtXML_To_Date(Mid$(xmlDomFiltros.selectSingleNode("//DataCaixa").Text, 10, 8))
                End If
            End If
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            .SetText COL_NIVEL_1, ROW_MOVIMENTACAO, "Movimentação"
            
            lngLinhaGrid = ROW_MOVIMENTACAO + 1
            
            If strItemCaixaGrupoVeiculoLegal <> vbNullString Then
                Set xmlItemCaixaGrupoVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
                Call xmlItemCaixaGrupoVeiculoLegal.loadXML(strItemCaixaGrupoVeiculoLegal)
            
                For Each objDomNode In xmlItemCaixaGrupoVeiculoLegal.selectNodes("/Repeat_ItensCaixa/*")
                    
                    If Not blnItemGenerico Then
                        blnItemGenerico = True
                    
                        .BlockMode = False
                        .Col = COL_NIVEL_1
                        .Row = lngLinhaGrid
                        .CellType = CellTypePicture
                        .TypePictCenter = True
                        .TypeHAlign = TypeHAlignCenter
                        
                        .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
                        .SetText COL_NIVEL_2, lngLinhaGrid, gstrItemGenerico
                        .SetText COL_CODIGO_OCULTO, lngLinhaGrid, xmlItensCaixa.selectSingleNode("Repeat_ItemCaixa/Grupo_ItemCaixa[DE_ITEM_CAIX='" & gstrItemGenerico & "']/CO_ITEM_CAIX").Text
                        
                        lngLinhaGrid = lngLinhaGrid + 1
                    End If
                    
                    For Each objDOMNodeAux In xmlItensCaixa.selectNodes("Repeat_ItemCaixa/Grupo_ItemCaixa[CO_ITEM_CAIX_NIVE_01='" & objDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "']")
                        If intCodGrupoVeicLegal = Val(objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                            
                            intNivelItemCaixa = fgObterNivelItemCaixa(objDOMNodeAux.selectSingleNode("CO_ITEM_CAIX").Text)
                            
                            If intNivelItemCaixa = 1 Then
                                .BlockMode = False
                                .Col = COL_NIVEL_1
                                .Row = lngLinhaGrid
                                .CellType = CellTypePicture
                                .TypePictCenter = True
                                .TypeHAlign = TypeHAlignCenter
                                
                                If objDOMNodeAux.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar Then
                                    .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
                                Else
                                    .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture
                                End If
                                
                                .SetText COL_NIVEL_2, lngLinhaGrid, objDOMNodeAux.selectSingleNode("DE_ITEM_CAIX").Text
                                .SetText COL_CODIGO_OCULTO, lngLinhaGrid, objDOMNodeAux.selectSingleNode("CO_ITEM_CAIX").Text
                                
                                lngLinhaGrid = lngLinhaGrid + 1
                            End If
                            
                        End If
                    Next
                
                    If intCodGrupoVeicLegal < Val(objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                        Exit For
                    End If
                            
                Next
            
            Else
                
                .BlockMode = False
                .Col = COL_NIVEL_1
                .Row = lngLinhaGrid
                .CellType = CellTypePicture
                .TypePictCenter = True
                .TypeHAlign = TypeHAlignCenter
                
                .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
                .SetText COL_NIVEL_2, lngLinhaGrid, gstrItemGenerico
                .SetText COL_CODIGO_OCULTO, lngLinhaGrid, xmlItensCaixa.selectSingleNode("Repeat_ItemCaixa/Grupo_ItemCaixa[DE_ITEM_CAIX='" & gstrItemGenerico & "']/CO_ITEM_CAIX").Text
                
                lngLinhaGrid = lngLinhaGrid + 1
    
            End If
            
            .SetText COL_NIVEL_1, lngLinhaGrid, "Fechamento"
            
            lngLinhaFinalSpread = lngLinhaGrid
            
            If Not xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva") Is Nothing Then
                strDataPosicaoCaixaSubReserva = xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva/Grupo_PosicaoCaixaSubReserva/DT_CAIX_SUB_RESE").Text
            End If
            
            For Each objDomNode In xmlDomLeitura.selectNodes("PosicaoMovimentoSubReserva/Repeat_SubReservaD0/*")
                blnAchouItem = False
                
                For lngLinhaGrid = ROW_MOVIMENTACAO + 1 To lngLinhaFinalSpread - 1
                    varItemCaixaAux = vbNullString
                    
                    .GetText COL_CODIGO_OCULTO, lngLinhaGrid, varItemCaixaAux
                    varItemCaixaAux = Left$(varItemCaixaAux, 4)
                    
                    If varItemCaixaAux = enumTipoCaixa.CaixaSubReserva & Replace(objDomNode.selectSingleNode("CO_ITEM_CAIX").Text, " ", vbNullString) Or _
                       varItemCaixaAux = vbNullString Then
                        blnAchouItem = True
                        
                        Exit For
                    End If
                Next
                
                If blnAchouItem Then
                    .SetText COL_PREVISTO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_PREV").Text)
                    .SetText COL_TOTAL, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_TOT_MOV").Text)
                    .SetText COL_REALIZADO_OU_SOLICITADO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_SOLI").Text)
                    .SetText COL_CONFIRMADO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_CONF").Text)
                    .SetText COL_VARIACAO_OU_REAL_PREV, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_VAR_MOV").Text)
                    .SetText COL_DATA_POSICAO_OCULTA, lngLinhaGrid, strDataPosicaoCaixaSubReserva
                End If
            Next

            Call flTotalizarMovimento(ROW_ABERTURA, ROW_MOVIMENTACAO, lngLinhaFinalSpread)
        End If
        
        .Redraw = True
    End With
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    Set xmlDomFiltroTela = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaD0 - flCarregarListaPorVeiculoLegal", 0

End Sub

' Identifica e colore valores negativos.

Private Sub flColorirValoresNegativos()

Dim lngLinhaGrid                            As Long
Dim lngColunaGrid                           As Long
Dim varConteudoCelula                       As Variant

On Error GoTo ErrorHandler

    With Me.vasLista
        For lngColunaGrid = COL_PREVISTO To .MaxCols
            .Col = lngColunaGrid
            
            For lngLinhaGrid = ROW_ABERTURA To lngLinhaFinalSpread
                .Row = lngLinhaGrid
            
                .GetText lngColunaGrid, lngLinhaGrid, varConteudoCelula
                If Trim(varConteudoCelula) = vbNullString And .BackColor <> &H8000000F Then
                    .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface(0)
                ElseIf Left$(Trim(varConteudoCelula), 1) = "-" Then
                    .BlockMode = False
                    .Col = lngColunaGrid
                    .Row = lngLinhaGrid
                    .ForeColor = vbRed
                End If
            Next
        Next
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flColorirValoresNegativos", 0
   
End Sub

Private Sub vasLista_Click(ByVal Col As Long, ByVal Row As Long)

Dim lngLinhaGrid                            As Long
Dim varConteudoCelula                       As Variant
Dim varRadicalItemPai                       As Variant
Dim intNivelItemPai                         As Integer
Dim intNivelItemCaixa                       As Integer

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Call flInicializarFlxDetalhe
    
    With vasLista
        .BlockMode = False
        .Col = Col
        .Row = Row

        If Not .TypePictPicture Is Nothing Then
            .GetText .MaxCols, .Row, varConteudoCelula
            If varConteudoCelula = vbNullString Then
                intNivelItemPai = 1
                varRadicalItemPai = "2"
            Else
                intNivelItemPai = fgObterNivelItemCaixa(varConteudoCelula)
                varRadicalItemPai = Left$(varConteudoCelula, ((intNivelItemPai) * 3) + 1)
            End If
            
            lngLinhaGrid = .Row + 1
            
            If .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture Then
                .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture

                Do
                    .GetText .MaxCols, lngLinhaGrid, varConteudoCelula
                    
                    intNivelItemCaixa = fgObterNivelItemCaixa(varConteudoCelula)
                    If intNivelItemCaixa > intNivelItemPai And InStr(1, varConteudoCelula, varRadicalItemPai) = 1 Then
                        .RowHeight(lngLinhaGrid) = 0
                    End If

                    lngLinhaGrid = lngLinhaGrid + 1
                    If lngLinhaGrid >= lngLinhaFinalSpread Then Exit Do
                Loop

            ElseIf .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture Then
                .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture

                If .RowHeight(.Row + 1) <> 0 Then
                    If Not Me.trvGeral.SelectedItem Is Nothing Then
                        'Especificamente para esta função, o 1º argumento deve ser passado como STRING
                        '<< trvGeral.SelectedItem.Key >>
                        Call flCarregarListaPorNiveisItemCaixa(Me.trvGeral.SelectedItem.Key, .Col, .Row)
                    End If
                Else
                    Do
                        .GetText .MaxCols, lngLinhaGrid, varConteudoCelula
                        intNivelItemCaixa = fgObterNivelItemCaixa(varConteudoCelula)
                        
                        If intNivelItemCaixa = intNivelItemPai + 1 And InStr(1, varConteudoCelula, varRadicalItemPai) = 1 Then
                            .BlockMode = False
                            .Col = intNivelItemCaixa + 1
                            .Row = lngLinhaGrid
                            
                            If .TypePictPicture <> imgOutrosIcones.ListImages("leaf").Picture Then
                                .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture
                            End If
                            
                            .RowHeight(lngLinhaGrid) = 225
                        End If
                        
                        lngLinhaGrid = lngLinhaGrid + 1
                        If lngLinhaGrid >= lngLinhaFinalSpread Then Exit Do
                    Loop
                End If

            End If
        
        Else
            Call flSelecionarCelulaSpread(Col, Row)
            
            If Not trvGeral.SelectedItem Is Nothing Then
                'Verifica se a lista refere-se a um Grupo de Veículo Legal...
                If InStr(2, trvGeral.SelectedItem.Key, "k_") = 0 Then
                    '...se sim, verifica se algum item foi selecionado...
                    If Row >= ROW_ABERTURA And Row <= lngLinhaFinalSpread Then
                        '...se sim, apresenta a Situação e Data do Caixa DISPONÍVEL
                        Call vasLista.GetText(COL_CODIGO_OCULTO, Row, varConteudoCelula)
                        If varConteudoCelula <> vbNullString Then
                            lblData.Caption = fgDescricaoEstadoCaixa(Val(Split(varConteudoCelula, "|")(2))) & " - "
                            
                            Call vasLista.GetText(COL_DATA_POSICAO_OCULTA, Row, varConteudoCelula)
                            lblData.Caption = lblData.Caption & fgDtXML_To_Date(varConteudoCelula)
                        End If
                    End If
                End If
            End If
        End If

    End With
    
    fgCursor
    
    Exit Sub
    
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - vasLista_Click"

End Sub

Private Sub vasLista_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim varTextoSpread                          As Variant
Dim objNodeBusca                            As MSComctlLib.Node
Dim intColunas                              As Integer
Dim blnMostraDetalhe                        As Boolean

On Error GoTo ErrorHandler

    If Not Me.trvGeral.SelectedItem Is Nothing Then
        'Verifica se os dados do GRID são referentes ao Grupo de Veículo Legal...
        If InStr(2, trvGeral.SelectedItem.Key, "k_") = 0 Then
            Call vasLista.GetText(COL_NIVEL_1, Row, varTextoSpread)
            If varTextoSpread <> vbNullString And Row >= ROW_ABERTURA Then
                Set objNodeBusca = objBuscaNo.ProcuraNoProx(varTextoSpread)
                If Not objNodeBusca Is Nothing Then
                    objNodeBusca.EnsureVisible
                    objNodeBusca.Selected = True
                    Call trvGeral_NodeClick(objNodeBusca)
                End If
            End If
        
        '...se não, então são do Veículo Legal
        Else
            With vasLista
                .Row = Row
                
                For intColunas = COL_PRINCIPAL To COL_DESCRICAO
                    .Col = intColunas
                    
                    If Not .TypePictPicture Is Nothing Then
                        If .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture Then
                            'Verifica se existe data de abertura
                            .Col = COL_DATA_POSICAO_OCULTA
                            If Len(.Text) = 8 Then
                                blnMostraDetalhe = True
                            Else
                                Call flInicializarFlxDetalhe
                            End If
                            
                            Exit For
                        Else
                            Call flInicializarFlxDetalhe
                            
                            Exit Sub
                        End If
                    End If
                Next
            End With
            
            If blnMostraDetalhe Then
                fgCursor True
                Call flCarregarDetalheMovimento(trvGeral.SelectedItem.Key, Row)
            End If
        End If
    End If
    
    fgCursor
    
    Exit Sub
    
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - vasLista_DblClick"
    
End Sub

Private Sub vasLista_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - vasLista_MouseMove"
End Sub
                                 
' Controla refresh da pesquisa atual da tela.

Private Sub flRefreshPosicaoAtualTela(Optional ByVal pstrPosicaoAtualPesquisa As String = vbNullString)

Dim xmlPosPesquisa                          As MSXML2.DOMDocument40

    Set xmlPosPesquisa = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlPosPesquisa.loadXML(pstrPosicaoAtualPesquisa)
    
    If trvGeral.SelectedItem Is Nothing Then
        With xmlPosPesquisa.documentElement
            If Not .selectSingleNode("Grupo_VeiculoLegal") Is Nothing And trvGeral.Nodes.Count > 1 Then
                trvGeral.Nodes(2).Selected = True
                trvGeral.Nodes(2).EnsureVisible
                Call trvGeral_NodeClick(trvGeral.Nodes(2))
                Exit Sub
            End If
        
            If Not .selectSingleNode("Grupo_GrupoVeiculoLegal") Is Nothing And trvGeral.Nodes.Count > 0 Then
                trvGeral.Nodes(1).Selected = True
                trvGeral.Nodes(1).EnsureVisible
                Call trvGeral_NodeClick(trvGeral.Nodes(1))
                Exit Sub
            End If
            
            If Not .selectSingleNode("Grupo_BancoLiquidante") Is Nothing And lngLinhaFinalSpread = 0 Then
                Call flCarregarTrvGeral(.selectSingleNode("//BancoLiquidante").Text)
                Exit Sub
            End If
        End With
        
        Set xmlPosPesquisa = Nothing
    Else
        Call trvGeral_NodeClick(trvGeral.SelectedItem)
    End If
    
End Sub

' Seleciona uma determinada célula na lista de movimentação.

Private Sub flSelecionarCelulaSpread(ByVal lngCol As Long, ByVal lngRow As Long)

Dim varConteudoCelula                       As Variant
Dim lngColunaGrid                           As Long

On Error GoTo ErrorHandler

    'Força o início da seleção para o início do detalhe,
    'apenas se a seleção de linha ou coluna estiver fora do esperado
    If lngRow <= 2 Or lngRow > lngLinhaFinalSpread Then lngRow = 3
    If lngCol <= 8 Then lngCol = 9

    With Me.vasLista
        'Verifica se o grid está preenchido, se não, sai da rotina
        .GetText COL_NIVEL_1, 3, varConteudoCelula
        If varConteudoCelula = vbNullString Then Exit Sub
    
        '>>> Retorna o ForeColor do Cabeçalho para a cor preta e
        '            o BackColor para a cor cinza
        .BlockMode = True
        .Row = 1
        .Col = 1
        .Row2 = 2
        .BackColor = &H8000000F
        .ForeColor = vbBlack
        .Col2 = .MaxCols
        .BlockMode = False
        
        '>>> Retorna as formatações do cabeçalho PREVISTO para o contorno original
        .BlockMode = True
        .Col = 9
        .Row = 1
        .Col2 = 9
        .Row2 = 2
        .CellBorderType = 0     'Sem contorno da seleção de células informada
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionSetCellBorder
        .BlockMode = False
        
        .BlockMode = True
        .Col = 9
        .Row = 1
        .Col2 = 9
        .Row2 = 2
        .CellBorderType = 16    'Contorno da seleção de células informada
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionSetCellBorder
        .BlockMode = False
        
        '>>> Retorna o BackColor das Colunas Principais para a cor branca
        .BlockMode = True
        .Col = 1
        .Row = 3
        .Col2 = 7
        .Row2 = .MaxRows
        .BackColor = &H8000000E
    
        '>>> Retorna o BackColor das Colunas de Valores para a cor branca
        .Col = 9
        .Row = 3
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .BackColor = vbWhite
        
        lngColunaGrid = 1
        Do
            .GetText lngColunaGrid, lngRow, varConteudoCelula
            If varConteudoCelula <> vbNullString Then Exit Do
            lngColunaGrid = lngColunaGrid + 1
        Loop
        
        '>>> Formata o BackColor das Colunas Principais para a cor amarela clara
        .Col = lngColunaGrid
        .Row = lngRow
        .Col2 = 7
        .Row2 = lngRow
        .BackColor = &HC0FFFF
        
        '>>> Formata o BackColor das Colunas de Valores para a cor amarela clara
        .Col = 9
        .Row = lngRow
        .Col2 = .MaxCols
        .Row2 = lngRow
        .BackColor = &HC0FFFF
        
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flSelecionarCelulaSpread", 0
    
End Sub

' Carrega movimentação por sub-níveis de item de caixa.

Private Sub flCarregarListaPorNiveisItemCaixa(ByVal strNodeSel As String, _
                                              ByVal intNivelItemDesejado As Integer, _
                                              ByVal lngLinhaItemPai As Integer)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaSubReserva  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaSubReserva  As A6MIU.clsMonitoracaoSubReserva
#End If

Dim objDomNode                              As IXMLDOMNode
Dim objDOMNodeAux                           As IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomFiltroTela                        As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlItemCaixaGrupoVeiculoLegal           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim arrPrimKey()                            As String

Dim intCodGrupoVeicLegal                    As Integer
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim intNivelItemCaixa                       As Integer

Dim lngLinhaGrid                            As Long
Dim intLinhasAdicionadas                    As Integer

Dim varItemCaixaAux                         As Variant
Dim varDataAux                              As Variant

Dim strDataPosicaoCaixaSubReserva           As String

Dim blnAchouItem                            As Boolean
Dim blnTemFilhos                            As Boolean

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    With Me.vasLista
        .Redraw = False

        '>>> Captura código do Grupo Veículo Legal... -------------------------------------------------
        arrPrimKey = Split(strNodeSel, "k_")
        intCodGrupoVeicLegal = Val(arrPrimKey(1))
        strCodigoVeiculoLegal = arrPrimKey(2)
        strSiglaSistema = arrPrimKey(3)
        
        '... e formata XML Filtro padrão
        Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                         "VeiculoLegal", strCodigoVeiculoLegal)
        Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                         "Sistema", strSiglaSistema)
        Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                         "NivelAgrupamento", intNivelItemDesejado)
                                         
        If strDocFiltros <> vbNullString Then
            Set xmlDomFiltroTela = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlDomFiltroTela.loadXML(strDocFiltros)
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Data/DataIni") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "DataCaixa", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Data/DataIni").Text)
            End If
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BancoLiquidante/BancoLiquidante") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "BancoLiquidante", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BancoLiquidante/BancoLiquidante").Text)
            End If
            
            If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BackOfficePerfilGeral/BackOfficePerfilGeral") Is Nothing Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                        "BackOfficePerfilGeral", _
                                        xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BackOfficePerfilGeral/BackOfficePerfilGeral").Text)
            End If
            
            Set xmlDomFiltroTela = Nothing
        End If
        '>>> -------------------------------------------------------------------------------------------
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMonitoracaoFluxoCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsMonitoracaoSubReserva")
        
        strRetLeitura = objMonitoracaoFluxoCaixaSubReserva.ObterMovimentoSubReserva(xmlDomFiltros.xml, _
                                                                                    vntCodErro, _
                                                                                    vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strRetLeitura <> vbNullString Then
            If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                '100 - Documento XML Inválido.
                lngCodigoErroNegocio = 100
                GoTo ErrorHandler
            End If
            
            Set xmlItemCaixaGrupoVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
            
            Call xmlItemCaixaGrupoVeiculoLegal.loadXML(strItemCaixaGrupoVeiculoLegal)
            
            lngLinhaGrid = lngLinhaItemPai + 1
            intLinhasAdicionadas = 0
           
            For Each objDomNode In xmlItemCaixaGrupoVeiculoLegal.selectNodes("/Repeat_ItensCaixa/*")
                For Each objDOMNodeAux In xmlItensCaixa.selectNodes("Repeat_ItemCaixa/Grupo_ItemCaixa[CO_ITEM_CAIX_NIVE_01='" & objDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "']")
                    
                    If intCodGrupoVeicLegal = Val(objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                        
                        intNivelItemCaixa = fgObterNivelItemCaixa(objDOMNodeAux.selectSingleNode("CO_ITEM_CAIX").Text)
                        
                        .GetText .MaxCols, lngLinhaItemPai, varItemCaixaAux
                        varItemCaixaAux = Left$(varItemCaixaAux, ((intNivelItemDesejado - 1) * 3) + 1)
                        
                        If intNivelItemCaixa = intNivelItemDesejado And _
                           varItemCaixaAux = Left$(objDOMNodeAux.selectSingleNode("CO_ITEM_CAIX").Text, ((intNivelItemDesejado - 1) * 3) + 1) Then
                           
                            blnTemFilhos = IIf(objDOMNodeAux.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar, False, True)
                            Call flInserirLinhaSpread(intNivelItemDesejado + 1, lngLinhaGrid, blnTemFilhos)
                            
                            .SetText intNivelItemDesejado + 2, lngLinhaGrid, objDOMNodeAux.selectSingleNode("DE_ITEM_CAIX").Text
                            .SetText .MaxCols, lngLinhaGrid, objDOMNodeAux.selectSingleNode("CO_ITEM_CAIX").Text
                            
                            lngLinhaGrid = lngLinhaGrid + 1
                            intLinhasAdicionadas = intLinhasAdicionadas + 1
                            
                        End If
                        
                    End If
                Next
                
                If intCodGrupoVeicLegal < Val(objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                    Exit For
                End If
                    
            Next
            
            lngLinhaFinalSpread = lngLinhaFinalSpread + intLinhasAdicionadas
            
            If Not xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva") Is Nothing Then
                strDataPosicaoCaixaSubReserva = xmlDomLeitura.selectSingleNode("PosicaoMovimentoSubReserva/Repeat_PosicaoCaixaSubReserva/Grupo_PosicaoCaixaSubReserva/DT_CAIX_SUB_RESE").Text
            End If
            
            For Each objDomNode In xmlDomLeitura.selectNodes("PosicaoMovimentoSubReserva/Repeat_SubReservaD0/*")
                blnAchouItem = False
                
                For lngLinhaGrid = lngLinhaItemPai + 1 To lngLinhaItemPai + intLinhasAdicionadas
                    .GetText .MaxCols, lngLinhaGrid, varItemCaixaAux
                    varItemCaixaAux = Left$(varItemCaixaAux, ((intNivelItemDesejado) * 3) + 1)
                    If varItemCaixaAux = enumTipoCaixa.CaixaSubReserva & Replace$(objDomNode.selectSingleNode("CO_ITEM_CAIX").Text, " ", vbNullString) Then
                        blnAchouItem = True
                        
                        Exit For
                    End If
                Next
                
                If blnAchouItem Then
                    .SetText COL_PREVISTO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_PREV").Text)
                    .SetText COL_TOTAL, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_TOT_MOV").Text)
                    .SetText COL_REALIZADO_OU_SOLICITADO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_SOLI").Text)
                    .SetText COL_CONFIRMADO, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_CONF").Text)
                    .SetText COL_VARIACAO_OU_REAL_PREV, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_VAR_MOV").Text)
                    .SetText COL_DATA_POSICAO_OCULTA, lngLinhaGrid, strDataPosicaoCaixaSubReserva
                End If
            Next
            
            Call flColorirValoresNegativos
        End If
        
        .Redraw = True
    End With
            
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaD0 - flCarregarListaPorNiveisItemCaixa", 0
    
End Sub

' Insere linhas ao grid já montado.

Private Sub flInserirLinhaSpread(ByVal lngColunaPicture As Long, _
                                 ByVal lngLinhaInserir As Long, _
                                 ByVal blnTemFilhos As Boolean)

On Error GoTo ErrorHandler

    With Me.vasLista
        .BlockMode = False
        .Col = lngColunaPicture
        .Row = lngLinhaInserir
        .Action = ActionInsertRow
        
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignCenter
        
        If blnTemFilhos Then
            .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture
        Else
            .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
        End If
        
        .BlockMode = True
        .Col = 1
        .Row = lngLinhaInserir
        .Col2 = COL_SEPARADOR - 1
        .Row2 = lngLinhaInserir
        .BackColorStyle = BackColorStyleOverVertGridOnly
        .BackColor = &H8000000E
        .BlockMode = False
    
        .BlockMode = True
        .Col = COL_SEPARADOR + 1
        .Row = lngLinhaInserir
        .Col2 = .MaxCols
        .Row2 = lngLinhaInserir
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInserirLinhaSpread", 0

End Sub

' Efetua totalização do movimento.

Private Sub flTotalizarMovimento(ByVal lngLinhaAbertura As Long, _
                                 ByVal lngLinhaMovimentacao As Long, _
                                 ByVal lngLinhaFechamento As Long)
                                 
Dim lngLinhaGrid                            As Long
Dim lngColunaGrid                           As Long

Dim varValorItemCaixa                       As Variant
Dim varValorMovimento                       As Variant
Dim varValorAbertura                        As Variant

Dim dblTotal                                As Double

On Error GoTo ErrorHandler

    With Me.vasLista
        For lngColunaGrid = COL_PREVISTO To .MaxCols - 1
            For lngLinhaGrid = ROW_ABERTURA To lngLinhaFinalSpread
                'Verifica se a linha é de ABERTURA ou de FECHAMENTO e,
                'se a coluna é de REALIZADO SOLICITADO ou REALIZADO CONFIRMADO ou VARIAÇÃO REAL-PREVISTO
                If Not ((lngLinhaGrid = ROW_ABERTURA Or _
                         lngLinhaGrid = lngLinhaFinalSpread) And _
                        (lngColunaGrid = COL_REALIZADO_OU_SOLICITADO Or _
                         lngColunaGrid = COL_CONFIRMADO Or _
                         lngColunaGrid = COL_VARIACAO_OU_REAL_PREV)) Then
                    
                    '...se não, atualiza com ZEROS, se o valor estiver em branco
                    .GetText lngColunaGrid, lngLinhaGrid, varValorItemCaixa
                    If varValorItemCaixa = vbNullString Then
                        .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface(0)
                    End If
                End If
            Next
        Next
        
        For lngColunaGrid = COL_PREVISTO To .MaxCols - 1
            dblTotal = 0
            For lngLinhaGrid = ROW_MOVIMENTACAO + 1 To lngLinhaFinalSpread - 1
                .GetText lngColunaGrid, lngLinhaGrid, varValorItemCaixa
                dblTotal = dblTotal + fgVlrXml_To_Decimal(varValorItemCaixa)
            Next
            .SetText lngColunaGrid, ROW_MOVIMENTACAO, fgVlrXml_To_Interface(dblTotal)
        Next
        
        For lngColunaGrid = COL_PREVISTO To COL_TOTAL
            .GetText lngColunaGrid, ROW_ABERTURA, varValorAbertura
            .GetText lngColunaGrid, ROW_MOVIMENTACAO, varValorMovimento
            dblTotal = fgVlrXml_To_Decimal(varValorAbertura) + fgVlrXml_To_Decimal(varValorMovimento)
            .SetText lngColunaGrid, lngLinhaFinalSpread, fgVlrXml_To_Interface(dblTotal)
        Next
        
        Call flColorirValoresNegativos
    End With

    Exit Sub
       
ErrorHandler:

    fgRaiseError App.EXEName, Me.Name, "frmSubReservaD0 - flTotalizarMovimento", 0

End Sub

' Formatação inicial do grid de detalhe do movimento.

Private Sub flInicializarFlxDetalhe()

Dim lngColunas                              As Long

On Error GoTo ErrorHandler

    With Me.flxDetalhe
        .Redraw = False

        .Rows = 0
        .Rows = MAX_FIXED_ROWS
        .FixedRows = 1
        .FixedCols = 0
        
        .Cols = 12
        .ColWidth(0) = 800
        .ColWidth(1) = 1500
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 2150
        .ColWidth(7) = 1500
        .ColWidth(8) = 1500
        .ColWidth(9) = 1800
        .ColWidth(10) = 1800
        .ColWidth(11) = 1500
        
        .MergeCells = flexMergeNever
        
        .TextMatrix(0, 0) = "Sistema"
        .TextMatrix(0, 1) = "Empresa"
        .TextMatrix(0, 2) = "Local Liquidação"
        .TextMatrix(0, 3) = "Tipo Liquidação"
        .TextMatrix(0, 4) = "Descrição Ativo"
        .TextMatrix(0, 5) = "CNPJ Contraparte"
        .TextMatrix(0, 6) = "Nome Contraparte"
        .TextMatrix(0, 7) = "Entrada"
        .TextMatrix(0, 8) = "Saída"
        .TextMatrix(0, 9) = "Situação Movimento"
        .TextMatrix(0, 10) = "Data Movimento"
        .TextMatrix(0, 11) = "Data Retorno"
        
        .ColAlignment(10) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
        .ColAlignment(11) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
        
        .Redraw = True
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInicializarFlxDetalhe", 0
            
End Sub

' Carrega detalhe do movimento selecionado.

Private Sub flCarregarDetalheMovimento(ByVal strNodeSel As String, _
                                       ByVal lngLinhaItem As Long)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaSubReserva  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaSubReserva  As A6MIU.clsMonitoracaoSubReserva
#End If

Dim objDomNode                              As IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomFiltroTela                        As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim arrPrimKey()                            As String

Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim varItemCaixa                            As Variant
Dim varData                                 As Variant
Dim blnDataCarregada                        As Boolean

Dim lngRow                                  As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    '>>> Captura código do Grupo Veículo Legal... -------------------------------------------------
    arrPrimKey = Split(strNodeSel, "k_")
    strCodigoVeiculoLegal = arrPrimKey(2)
    strSiglaSistema = arrPrimKey(3)
    
    vasLista.GetText COL_CODIGO_OCULTO, lngLinhaItem, varItemCaixa
    vasLista.GetText COL_DATA_POSICAO_OCULTA, lngLinhaItem, varData
    
    '... e formata XML Filtro padrão
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DetalheMovimento", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_DetalheMovimento", _
                                     "VeiculoLegal", strCodigoVeiculoLegal)
    Call fgAppendNode(xmlDomFiltros, "Grupo_DetalheMovimento", _
                                     "Sistema", strSiglaSistema)
    Call fgAppendNode(xmlDomFiltros, "Grupo_DetalheMovimento", _
                                     "ItemCaixa", varItemCaixa)
    
    If strDocFiltros <> vbNullString Then
        Set xmlDomFiltroTela = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltroTela.loadXML(strDocFiltros)
        
        If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Data/DataIni") Is Nothing Then
            Call fgAppendNode(xmlDomFiltros, "Grupo_DetalheMovimento", _
                                    "DataMovimento", _
                                    Mid$(xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_Data/DataIni").Text, 10, 8))
            blnDataCarregada = True
        End If
        
        If Not xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BancoLiquidante/BancoLiquidante") Is Nothing Then
            Call fgAppendNode(xmlDomFiltros, "Grupo_DetalheMovimento", _
                                    "BancoLiquidante", _
                                    xmlDomFiltroTela.documentElement.selectSingleNode("Grupo_BancoLiquidante/BancoLiquidante").Text)
        End If
        
        Set xmlDomFiltroTela = Nothing
    End If
                                     
    If Not blnDataCarregada Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_DetalheMovimento", _
                                         "DataMovimento", varData)
    End If
    '>>> -------------------------------------------------------------------------------------------
        
    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMonitoracaoFluxoCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsMonitoracaoSubReserva")
    
    Call flInicializarFlxDetalhe
    
    strRetLeitura = objMonitoracaoFluxoCaixaSubReserva.ObterDetalheMovimento(xmlDomFiltros.xml, _
                                                                             vntCodErro, _
                                                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strRetLeitura <> vbNullString Then
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            '100 - Documento XML Inválido.
            lngCodigoErroNegocio = 100
            GoTo ErrorHandler
        End If
    
        With Me.flxDetalhe
            .Redraw = False

            lngRow = 1
            For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                
                If lngRow >= MAX_FIXED_ROWS Then
                    .Rows = lngRow + 1
                End If
                
                .TextMatrix(lngRow, COL_SISTEMA) = objDomNode.selectSingleNode("SG_SIST").Text
                .TextMatrix(lngRow, COL_EMPRESA) = objDomNode.selectSingleNode("NO_REDU_EMPR").Text
                .TextMatrix(lngRow, COL_LOCAL_LIQUIDACAO) = objDomNode.selectSingleNode("SG_LOCA_LIQU").Text
                .TextMatrix(lngRow, COL_TIPO_LIQUIDACAO) = objDomNode.selectSingleNode("DE_TIPO_LIQU").Text
                .TextMatrix(lngRow, COL_ENTRADA) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("ENTRADA").Text)
                .TextMatrix(lngRow, COL_SAIDA) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("SAIDA").Text)
                .TextMatrix(lngRow, COL_DESCRICAO_ATIVO) = objDomNode.selectSingleNode("DE_ATIV").Text
                .TextMatrix(lngRow, COL_CNPJ_CONTRAPARTE) = objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text
                .TextMatrix(lngRow, COL_NOME_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                .TextMatrix(lngRow, COL_DATA_MOVIMENTO) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_MOVI_CAIX_SUB_RESE").Text)
                
                If objDomNode.selectSingleNode("DT_RETN_OPER").Text <> gstrDataVazia Then
                    .TextMatrix(lngRow, COL_DATA_RETORNO) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_RETN_OPER").Text)
                End If
                
                If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("ENTRADA").Text) = 0 Then
                    .TextMatrix(lngRow, COL_ENTRADA) = vbNullString
                End If
                
                If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("SAIDA").Text) = 0 Then
                    .TextMatrix(lngRow, COL_SAIDA) = vbNullString
                Else
                    .FillStyle = flexFillSingle
                    .Row = lngRow
                    .Col = COL_SAIDA
                    .RowSel = lngRow
                    .ColSel = COL_SAIDA
                    .CellForeColor = vbRed
                End If
                
                Select Case objDomNode.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text
                Case enumTipoMovimento.EstornoPrevisto
                    .TextMatrix(lngRow, COL_SITUACAO_MOVIMENTO) = "Estorno Previsão"
                Case enumTipoMovimento.EstornoRealizadoConfirmado
                    .TextMatrix(lngRow, COL_SITUACAO_MOVIMENTO) = "Estorno Realizado Confirmado"
                Case enumTipoMovimento.EstornoRealizadoSolicitado
                    .TextMatrix(lngRow, COL_SITUACAO_MOVIMENTO) = "Estorno Realizado Solicitado"
                Case enumTipoMovimento.Previsto
                    .TextMatrix(lngRow, COL_SITUACAO_MOVIMENTO) = "Previsão"
                Case enumTipoMovimento.RealizadoConfirmado
                    .TextMatrix(lngRow, COL_SITUACAO_MOVIMENTO) = "Realizado Confirmado"
                Case enumTipoMovimento.RealizadoSolicitado
                    .TextMatrix(lngRow, COL_SITUACAO_MOVIMENTO) = "Realizado Solicitado"
                End Select
                
                lngRow = lngRow + 1
            Next
        
            .Redraw = True
        End With
    End If
            
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomFiltroTela = Nothing
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaSubReserva = Nothing
        
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaD0 - flCarregarDetalheMovimento", 0

End Sub

' Captura situação atual da pesquisa.

Private Function flCapturaSituacaoGrid() As String

Dim intCont                                 As Integer
Dim intContAux                              As Integer
Dim vntConteudoCelula                       As Variant
Dim strRetorno                              As String
Dim strRetornoDetalhe                       As String

    With vasLista
        For intCont = ROW_MOVIMENTACAO + 1 To lngLinhaFinalSpread - 1
            .Row = intCont
            For intContAux = COL_NIVEL_1 To COL_NIVEL_5
                .Col = intContAux
                
                'Verifica se para a linha selecionada...
                If .BackColor = &HC0FFFF Then
                    '...foi apresentada a lista DETALHE...
                    If Trim(flxDetalhe.TextMatrix(1, 0)) <> vbNullString Then
                        Call .GetText(COL_CODIGO_OCULTO, intCont, vntConteudoCelula)
                        
                        '...se sim, captura os dados da linha
                        strRetornoDetalhe = trvGeral.SelectedItem.Key & "|" & _
                                            vntConteudoCelula & "|" & _
                                            1 & ";"
                                
                        vntConteudoCelula = Empty
                    End If
                End If
                
                'Verifica se existe algum ITEM DE CAIXA expandido...
                If Not .TypePictPicture Is Nothing Then
                    If .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture Then
                        Call .GetText(COL_CODIGO_OCULTO, intCont, vntConteudoCelula)
                        
                        '...se sim, captura os dados da linha
                        strRetorno = strRetorno & _
                                trvGeral.SelectedItem.Key & "|" & _
                                vntConteudoCelula & "|" & _
                                0 & ";"
                        
                        vntConteudoCelula = Empty
                    End If
                End If
            Next
        Next
    End With
    
    'Adiciona as informações de DETALHE, se houver, no final da STRING
    strRetorno = strRetorno & strRetornoDetalhe
    If Trim(strRetorno) <> vbNullString Then
        strRetorno = Left(strRetorno, Len(strRetorno) - 1)
    End If

    flCapturaSituacaoGrid = strRetorno
    
End Function

' Monta novamente o grid a partir do que foi capturado.

Private Sub flRemontaGrids(ByVal pstrSituacaoGrid As String)

Dim intCont                                 As Integer
Dim intContAux                              As Integer
Dim intNivelItemDesejado                    As Integer
Dim strChave                                As String
Dim vntItemCaixa                            As Variant
Dim vntConteudoCelula                       As Variant
Dim blnAchouItemCaixa                       As Boolean
Dim blnAchouDetalhe                         As Boolean

    With vasLista
        'Percorre a lista de configurações dos GRIDs armazenada ANTES do REFRESH...
        For intCont = LBound(Split(pstrSituacaoGrid, ";")) To UBound(Split(pstrSituacaoGrid, ";"))
            '...captura as configurações (CHAVE, ITEM CAIXA, NÍVEL e se é ou não DETALHE)
            strChave = Split(Split(pstrSituacaoGrid, ";")(intCont), "|")(0)
            vntItemCaixa = Split(Split(pstrSituacaoGrid, ";")(intCont), "|")(1)
            intNivelItemDesejado = fgObterNivelItemCaixa(vntItemCaixa) + 1
            blnAchouDetalhe = (Split(Split(pstrSituacaoGrid, ";")(intCont), "|")(2) = 1)
            
            '...procura o ITEM CAIXA nas linhas ATUAIS do GRID...
            blnAchouItemCaixa = False
            For intContAux = ROW_MOVIMENTACAO + 1 To lngLinhaFinalSpread - 1
                Call .GetText(COL_CODIGO_OCULTO, intContAux, vntConteudoCelula)
                
                If Trim(vntConteudoCelula) = Trim(vntItemCaixa) Then
                    blnAchouItemCaixa = True
                    
                    Exit For
                End If
            Next
            
            '...se achou, verifica o tipo de AÇÃO a ser tomada...
            If blnAchouItemCaixa Then
                If blnAchouDetalhe Then
                    '...se for DETALHE, destaca a LINHA referente e carrega o detalhe
                    Call flSelecionarCelulaSpread(COL_PREVISTO, intContAux)
                    Call flCarregarDetalheMovimento(strChave, intContAux)
                Else
                    '...se for EXPLOSÃO de SUB-NÍVEIS, substitui a figura no GRID para MINUS e
                    '                                  carrega os dados respectivos
                    .Row = intContAux
                    .Col = intNivelItemDesejado
                    .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture
                    Call flCarregarListaPorNiveisItemCaixa(strChave, intNivelItemDesejado, intContAux)
                End If
            End If
        Next
    End With
    
End Sub
