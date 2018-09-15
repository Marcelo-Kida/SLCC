VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaixaFuturo 
   BackColor       =   &H8000000B&
   Caption         =   "Sub-reserva - Futuro"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDetalheVeicLega 
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3795
   End
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   661
   End
   Begin FPSpread.vaSpread vasLista 
      Height          =   1245
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Movimento Futuro"
      Top             =   930
      Width           =   9375
      _Version        =   196608
      _ExtentX        =   16536
      _ExtentY        =   2196
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
      SpreadDesigner  =   "frmCaixaFuturo.frx":0000
      UnitType        =   2
   End
   Begin MSComctlLib.ImageList imgOutrosIcones 
      Left            =   630
      Top             =   3990
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
            Picture         =   "frmCaixaFuturo.frx":01E8
            Key             =   "itemgrupo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":053A
            Key             =   "selectednode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":088C
            Key             =   "itemelementar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":0BDE
            Key             =   "treeminus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":0F00
            Key             =   "treeplus"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":1222
            Key             =   "leaf"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7560
      Width           =   13380
      _ExtentX        =   23601
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
         NumButtons      =   8
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
      EndProperty
   End
   Begin MSComctlLib.TreeView trvGeral 
      Height          =   3555
      Left            =   30
      TabIndex        =   3
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
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   30
      Top             =   3990
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
            Picture         =   "frmCaixaFuturo.frx":1574
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":1686
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":1798
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":1AEA
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":1E3C
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":218E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":24E0
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaixaFuturo.frx":2932
            Key             =   "posterior"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetalhe 
      Height          =   1185
      Left            =   60
      TabIndex        =   4
      Tag             =   "Detalhe do Movimento Futuro"
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
   Begin MSComctlLib.Toolbar tlbPaginacao 
      Height          =   330
      Left            =   3990
      TabIndex        =   6
      Top             =   3630
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   582
      ButtonWidth     =   2884
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Período Anterior"
            Key             =   "anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Período Posterior"
            Key             =   "proximo"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   30
      MousePointer    =   7  'Size N S
      Top             =   6180
      Width           =   13320
   End
   Begin VB.Image imgDummyV 
      Height          =   3465
      Left            =   3855
      MousePointer    =   9  'Size W E
      Top             =   435
      Width           =   90
   End
   Begin VB.Label lblBarra 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   3750
      TabIndex        =   1
      Top             =   0
      Width           =   615
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
      Left            =   7095
      TabIndex        =   0
      Top             =   30
      Width           =   300
   End
End
Attribute VB_Name = "frmCaixaFuturo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a visualização da movimentação futura de itens de caixa.

Option Explicit

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1
Private WithEvents objBuscaNo               As frmBuscaNo
Attribute objBuscaNo.VB_VarHelpID = -1
Private xmlFiltro                           As String

Private Enum enumTipoPesquisa
    None = 0
    PorGrupoVeiculoLegal = 1
    PorVeiculoLegal = 2
    PorNiveisItemCaixa = 3
    PorDetalheLancamento = 4
End Enum

Private xmlItensCaixa                       As MSXML2.DOMDocument40
Private strCarregaTreeView                  As String
Private strItemCaixaGrupoVeiculoLegal       As String

Private lngAlturaTableCombo                 As Long

Private Const strFuncionalidade             As String = "frmItemCaixa"
Private Const strTableComboInicial          As String = "Empresa"
Private Const intLinhaSdoInicial            As Integer = 2
Private Const intLinhaMovimentacao          As Integer = 3
Private Const intColunaInicioValores        As Integer = 9

Private Const COL_SISTEMA                   As Integer = 0
Private Const COL_EMPRESA                   As Integer = 1
Private Const COL_LOCAL_LIQUIDACAO          As Integer = 2
Private Const COL_TIPO_LIQUIDACAO           As Integer = 3
Private Const COL_DESCRICAO_ATIVO           As Integer = 4
Private Const COL_CNPJ_CONTRAPARTE          As Integer = 5
Private Const COL_NOME_CONTRAPARTE          As Integer = 6
Private Const COL_ENTRADA                   As Integer = 7
Private Const COL_SAIDA                     As Integer = 8
Private Const COL_DATA_MOVIMENTO            As Integer = 9
Private Const COL_DATA_RETORNO              As Integer = 10

Private Const MAX_FIXED_ROWS                As Integer = 100

Private fblnDummyV                          As Boolean
Private fblnDummyH                          As Boolean
Private blnEventByPass                      As Boolean

'>>>>> Declaração variáveis para controle geral do movimento
Private lngLinhaSdoFinal                    As Long
Private intTipoBackOfficeUsuario            As Integer
Private intQuantDiasPeriodo                 As Integer
Private datD0                               As Date
Private dblSaldoInicialD0                   As Double
Private dblSaldoInicialD1                   As Double

'>>>>> Declaração variáveis para controle de Refresh e Paginação da Tela
Private intUltPesqEfetuada                  As enumTipoPesquisa
Private objUltNodeSel                       As MSComctlLib.Node
Private lngUltColunaData                    As Long
Private lngUltLinhaItem                     As Long
Private datUltDataBase                      As Date
Private intUltPaginacao                     As enumPaginacao
Private intUltNivelDesejado                 As Integer

' Arranja telas de exibição de dados, conforme seleção pelo usuário.

Private Sub flArranjarJanelasExibicao(ByVal pstrJanelas As String)

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
    
    tlbPaginacao.Visible = vasLista.Visible
    txtDetalheVeicLega.Visible = trvGeral.Visible
    Call Form_Resize
    
End Sub

' Calcula saldos inicial e final de cada paginação de movimentação.

Private Sub flCalcularSaldosInicialEFinal()
        
Dim lngColunaGrid                           As Long

Dim varValorMovimento                       As Variant
Dim varValorSdoInicial                      As Variant

Dim dblTotal                                As Double
    
    With Me.vasLista
        .SetText intColunaInicioValores, intLinhaSdoInicial, fgVlrXml_To_Interface(dblSaldoInicialD0)
        .GetText intColunaInicioValores, intLinhaMovimentacao, varValorMovimento
        
        dblTotal = dblSaldoInicialD0 + fgVlrXml_To_Decimal(varValorMovimento)
        .SetText intColunaInicioValores, lngLinhaSdoFinal, fgVlrXml_To_Interface(dblTotal)
        
        dblTotal = dblSaldoInicialD0 + dblSaldoInicialD1
        .SetText intColunaInicioValores + 1, intLinhaSdoInicial, fgVlrXml_To_Interface(dblTotal)
        
        For lngColunaGrid = intColunaInicioValores + 1 To .MaxCols - 1
            .GetText lngColunaGrid, intLinhaSdoInicial, varValorSdoInicial
            .GetText lngColunaGrid, intLinhaMovimentacao, varValorMovimento
            
            dblTotal = fgVlrXml_To_Decimal(varValorSdoInicial) + fgVlrXml_To_Decimal(varValorMovimento)
            
            .SetText lngColunaGrid, lngLinhaSdoFinal, fgVlrXml_To_Interface(dblTotal)
            .SetText lngColunaGrid + 1, intLinhaSdoInicial, fgVlrXml_To_Interface(dblTotal)
        Next
    End With

End Sub

' Carrega detalhe da movimentação selecionada.

Private Sub flCarregarDetalheMovimento(ByVal pobjNodeSel As MSComctlLib.Node, _
                                       ByVal plngColunaData As Long, _
                                       ByVal plngLinhaItem As Long)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaFuturo  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaFuturo  As A6MIU.clsMonitoracaoFluxoCaixaFuturo
#End If

Dim xmlDomNode                          As IXMLDOMNode
Dim xmlDomLeitura                       As MSXML2.DOMDocument40
Dim strRetLeitura                       As String

Dim strCodVeicLegal                     As String
Dim strSiglaSistema                     As String
Dim strData                             As String
Dim varItemCaixa                        As Variant
Dim varData                             As Variant

Dim lngRow                              As Long
Dim arrPrimKey()                        As String
Dim vntCodErro                          As Variant
Dim vntMensagemErro                     As Variant

On Error GoTo ErrorHandler
    
    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMonitoracaoFluxoCaixaFuturo = fgCriarObjetoMIU("A6MIU.clsMonitoracaoFluxoCaixaFuturo")
    
    arrPrimKey = Split(pobjNodeSel.Key, "k_")
    strCodVeicLegal = arrPrimKey(2)
    strSiglaSistema = arrPrimKey(3)
        
    vasLista.GetText vasLista.MaxCols, plngLinhaItem, varItemCaixa
    vasLista.GetText plngColunaData, 1, varData
    
    strData = CStr(Format(CDate(varData), "DD/MM/YYYY"))
    
    Call flInicializarFlxDetalhe
    
    strRetLeitura = objMonitoracaoFluxoCaixaFuturo.ObterDetalheMovimento(strCodVeicLegal, _
                                                                         strSiglaSistema, _
                                                                         varItemCaixa, _
                                                                         strData, _
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
            For Each xmlDomNode In xmlDomLeitura.selectNodes("/Repeat_DetalheMovimento/*")
                             
                 If lngRow >= MAX_FIXED_ROWS Then
                    .Rows = lngRow + 1
                 End If
                             
                .TextMatrix(lngRow, COL_SISTEMA) = xmlDomNode.selectSingleNode("SG_SIST").Text
                .TextMatrix(lngRow, COL_EMPRESA) = xmlDomNode.selectSingleNode("NO_REDU_EMPR").Text
                .TextMatrix(lngRow, COL_LOCAL_LIQUIDACAO) = xmlDomNode.selectSingleNode("SG_LOCA_LIQU").Text
                .TextMatrix(lngRow, COL_TIPO_LIQUIDACAO) = xmlDomNode.selectSingleNode("DE_TIPO_LIQU").Text
                .TextMatrix(lngRow, COL_ENTRADA) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("ENTRADA").Text)
                .TextMatrix(lngRow, COL_SAIDA) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("SAIDA").Text)
                .TextMatrix(lngRow, COL_DESCRICAO_ATIVO) = xmlDomNode.selectSingleNode("DE_ATIV").Text
                .TextMatrix(lngRow, COL_CNPJ_CONTRAPARTE) = xmlDomNode.selectSingleNode("CO_CNPJ_CNPT").Text
                .TextMatrix(lngRow, COL_NOME_CONTRAPARTE) = xmlDomNode.selectSingleNode("NO_CNPT").Text
                .TextMatrix(lngRow, COL_DATA_MOVIMENTO) = fgDtXML_To_Interface(xmlDomNode.selectSingleNode("DT_LIQU_OPER").Text)
                
                If xmlDomNode.selectSingleNode("DT_RETN_OPER").Text <> gstrDataVazia Then
                    .TextMatrix(lngRow, COL_DATA_RETORNO) = fgDtXML_To_Interface(xmlDomNode.selectSingleNode("DT_RETN_OPER").Text)
                End If
                
                If fgVlrXml_To_Decimal(xmlDomNode.selectSingleNode("ENTRADA").Text) = 0 Then
                    .TextMatrix(lngRow, COL_ENTRADA) = vbNullString
                End If
                
                If fgVlrXml_To_Decimal(xmlDomNode.selectSingleNode("SAIDA").Text) = 0 Then
                    .TextMatrix(lngRow, COL_SAIDA) = vbNullString
                Else
                    .FillStyle = flexFillSingle
                    .Row = lngRow
                    .Col = COL_SAIDA
                    .RowSel = lngRow
                    .ColSel = COL_SAIDA
                    .CellForeColor = vbRed
                End If
                
                .FillStyle = flexFillSingle
                .Row = lngRow
                .Col = COL_ENTRADA
                .CellAlignment = flexAlignRightCenter
                .Col = COL_SAIDA
                .CellAlignment = flexAlignRightCenter
        
                lngRow = lngRow + 1
            Next
        
            .Redraw = True
        End With
    End If
            
    intUltPesqEfetuada = PorDetalheLancamento
    Set objUltNodeSel = pobjNodeSel
    lngUltColunaData = plngColunaData
    lngUltLinhaItem = plngLinhaItem
    
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flCarregarDetalheMovimento"

End Sub

' Carrega lista de itens de caixa.

Private Sub flCarregarListaItensCaixa()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A6MIU.clsMIU
#End If

Dim xmlDomNode             As IXMLDOMNode
Dim strLerTodos            As String
Dim xmlFiltroLeitura       As MSXML2.DOMDocument40
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set xmlItensCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlFiltroLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    
    If Not xmlItensCaixa.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.SBR, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlItensCaixa, App.EXEName, Me.Name, "flCarregarListaItensCaixa")
    End If
    
    Set xmlDomNode = xmlItensCaixa.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa")
    xmlDomNode.selectSingleNode("@Operacao").Text = "LerTodos"
    xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaFuturo

    Call xmlFiltroLeitura.loadXML(xmlDomNode.xml)
    
    If xmlFiltro <> vbNullString Then
        Call fgAppendXML(xmlFiltroLeitura, "Grupo_ItemCaixa", xmlFiltro)
    End If
    
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
    Set xmlFiltroLeitura = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlItensCaixa = Nothing
    Set xmlFiltroLeitura = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarListaItensCaixa", 0

End Sub

' Carrega movimentação por grupo de veículo legal.

Private Sub flCarregarMovPorGrupoVeiculoLegal(ByVal pobjNodeSel As MSComctlLib.Node, _
                                              ByVal pdatDataBase As Date, _
                                              ByVal pintPaginacao As enumPaginacao)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaFuturo  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaFuturo  As A6MIU.clsMonitoracaoFluxoCaixaFuturo
#End If

Dim xmlDomFiltro                        As MSXML2.DOMDocument40
Dim xmlDomNode                          As IXMLDOMNode
Dim xmlDomLeitura                       As MSXML2.DOMDocument40
Dim strRetLeitura                       As String

Dim lngLinhaGrid                        As Integer
Dim lngColunaGrid                       As Integer

Dim intCodGrupoVeicLegal                As Integer
Dim varVeicLegalAux                     As Variant
Dim varDataAux                          As Variant
Dim strCodVeiculoLegal                  As String
Dim strDataBase                         As String
Dim lngCodEmpresa                       As Long

Dim arrPrimKey()                        As String
Dim vntCodErro                          As Variant
Dim vntMensagemErro                     As Variant
    
On Error GoTo ErrorHandler
    
    If Trim(xmlFiltro) <> vbNullString Then
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltro.loadXML(xmlFiltro)
        
        If Not xmlDomFiltro.selectSingleNode("//VeiculoLegal") Is Nothing Then
            strCodVeiculoLegal = xmlDomFiltro.selectSingleNode("//VeiculoLegal").Text
        End If
        
        If Not xmlDomFiltro.selectSingleNode("//BancoLiquidante") Is Nothing Then
            lngCodEmpresa = Val(xmlDomFiltro.selectSingleNode("//BancoLiquidante").Text)
        End If
        
        Set xmlDomFiltro = Nothing
    End If
    
    With Me.vasLista
        .Redraw = False

        Call flInicializarVasLista
        
        .SetText 1, 1, pobjNodeSel.Text
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        arrPrimKey = Split(pobjNodeSel.Key, "k_")
        intCodGrupoVeicLegal = arrPrimKey(1)
        
        Set objMonitoracaoFluxoCaixaFuturo = fgCriarObjetoMIU("A6MIU.clsMonitoracaoFluxoCaixaFuturo")
        
        strDataBase = CStr(Format(pdatDataBase, "DD/MM/YYYY"))
        
        strRetLeitura = objMonitoracaoFluxoCaixaFuturo.ObterMovimentoPorGrupoVeiculoLegal(intCodGrupoVeicLegal, _
                                                                                          strDataBase, _
                                                                                          intQuantDiasPeriodo, _
                                                                                          pintPaginacao, _
                                                                                          vntCodErro, _
                                                                                          vntMensagemErro, _
                                                                                          lngCodEmpresa, _
                                                                                          strCodVeiculoLegal)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            '100 - Documento XML Inválido.
            lngCodigoErroNegocio = 100
            GoTo ErrorHandler
        End If
        
        '>>>>> Popula datas de movimento no Spread
        lngColunaGrid = intColunaInicioValores
        For Each xmlDomNode In xmlDomLeitura.selectNodes("//DatasMovimento/*")
            
            .SetText lngColunaGrid, 1, Trim$(fgDtXML_To_Interface(xmlDomNode.firstChild.Text))
            lngColunaGrid = lngColunaGrid + 1
            If lngColunaGrid = .MaxCols Then Exit For
        
        Next
        
        .MaxCols = intColunaInicioValores + intQuantDiasPeriodo + 1
        .ColWidth(.MaxCols) = 0
    
        If Not xmlDomLeitura.documentElement.selectSingleNode("Repeat_MovGrupoVeicLegal") Is Nothing Then
            
            '>>>>> Popula Grupos de Veículos Legais no Spread
            lngLinhaGrid = 2
            varVeicLegalAux = vbNullString
            
            For Each xmlDomNode In xmlDomLeitura.selectNodes("//Repeat_MovGrupoVeicLegal/*")
                
                If varVeicLegalAux <> xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text Then
                    varVeicLegalAux = xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SetText 2, lngLinhaGrid, varVeicLegalAux
                    lngLinhaGrid = lngLinhaGrid + 1
                End If
            
            Next
            
            lngLinhaSdoFinal = lngLinhaGrid - 1
            
            '>>>>> Atribui o movimento ao Grupo de Veículo Legal e à Data corretos
            For Each xmlDomNode In xmlDomLeitura.selectNodes("//Repeat_MovGrupoVeicLegal/*")
                varVeicLegalAux = vbNullString
                For lngLinhaGrid = 2 To .MaxRows
                    .GetText 2, lngLinhaGrid, varVeicLegalAux
                    If varVeicLegalAux = xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text Or _
                       varVeicLegalAux = vbNullString Then
                        Exit For
                    End If
                Next
                
                varDataAux = vbNullString
                For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
                    .GetText lngColunaGrid, 1, varDataAux
                    If Format$(varDataAux, "yyyymmdd") = xmlDomNode.selectSingleNode("DT_LIQU_OPER").Text Or varDataAux = vbNullString Then
                        Exit For
                    End If
                Next
                
                If varVeicLegalAux <> vbNullString And varDataAux <> vbNullString Then
                    .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_LIQU_OPER").Text)
                End If
            Next
            
            Call flIdentificarValoresNegativos
        
            For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
                If .MaxTextColWidth(lngColunaGrid) > .ColWidth(lngColunaGrid) Then
                    .ColWidth(lngColunaGrid) = .MaxTextColWidth(lngColunaGrid) + 200
                End If
            Next
        End If
        
        .Redraw = True
    End With
            
    intUltPesqEfetuada = PorGrupoVeiculoLegal
    Set objUltNodeSel = pobjNodeSel
    datUltDataBase = pdatDataBase
    intUltPaginacao = pintPaginacao
    
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    Set xmlDomFiltro = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flCarregarMovPorGrupoVeiculoLegal"

End Sub

' Carrega movimentação por níveis de item de caixa.

Private Sub flCarregarMovPorNiveisItemCaixa(ByVal pobjNodeSel As MSComctlLib.Node, _
                                            ByVal pdatDataBase As Date, _
                                            ByVal pintPaginacao As enumPaginacao, _
                                            ByVal pintNivelItemDesejado As Integer, _
                                            ByVal plngLinhaItemPai As Integer)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaFuturo  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaFuturo  As A6MIU.clsMonitoracaoFluxoCaixaFuturo
#End If

Dim xmlDomNode                          As IXMLDOMNode
Dim xmlDomNodeAux                       As IXMLDOMNode
Dim xmlDomLeitura                       As MSXML2.DOMDocument40
Dim xmlItemCaixaGrupoVeiculoLegal       As MSXML2.DOMDocument40
Dim strRetLeitura                       As String

Dim intCodGrupoVeicLegal                As Integer
Dim strCodVeicLegal                     As String
Dim strSiglaSistema                     As String
Dim strD0                               As String
Dim strDataBase                         As String
Dim intNivelItemCaixa                   As Integer

Dim lngLinhaGrid                        As Long
Dim lngColunaGrid                       As Long
Dim intLinhasAdicionadas                As Integer

Dim varItemCaixaAux                     As Variant
Dim blnTemFilhos                        As Boolean

Dim arrPrimKey()                        As String
Dim vntCodErro                          As Variant
Dim vntMensagemErro                     As Variant

On Error GoTo ErrorHandler
    
    With Me.vasLista
        .Redraw = False

        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set xmlItemCaixaGrupoVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
        
        Set objMonitoracaoFluxoCaixaFuturo = fgCriarObjetoMIU("A6MIU.clsMonitoracaoFluxoCaixaFuturo")
        
        arrPrimKey = Split(pobjNodeSel.Key, "k_")
        intCodGrupoVeicLegal = arrPrimKey(1)
        strCodVeicLegal = arrPrimKey(2)
        strSiglaSistema = arrPrimKey(3)
        
        strD0 = CStr(Format(datD0, "DD/MM/YYYY"))
        strDataBase = CStr(Format(pdatDataBase, "DD/MM/YYYY"))
        
        strRetLeitura = objMonitoracaoFluxoCaixaFuturo.ObterMovimentoPorVeiculoLegal(strCodVeicLegal, _
                                                                                     strSiglaSistema, _
                                                                                     pintNivelItemDesejado, _
                                                                                     strD0, _
                                                                                     strDataBase, _
                                                                                     intQuantDiasPeriodo, _
                                                                                     pintPaginacao, _
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
            
            Call xmlItemCaixaGrupoVeiculoLegal.loadXML(strItemCaixaGrupoVeiculoLegal)
            
            lngLinhaGrid = plngLinhaItemPai + 1
            intLinhasAdicionadas = 0
            
            '>>>>> Popula Itens de Caixa no Spread
            For Each xmlDomNode In xmlItemCaixaGrupoVeiculoLegal.selectNodes("/Repeat_ItensCaixa/*")
                For Each xmlDomNodeAux In xmlItensCaixa.selectNodes("Repeat_ItemCaixa/Grupo_ItemCaixa[CO_ITEM_CAIX_NIVE_01='" & xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "']")
                    If intCodGrupoVeicLegal = Val(xmlDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                        
                        intNivelItemCaixa = fgObterNivelItemCaixa(xmlDomNodeAux.selectSingleNode("CO_ITEM_CAIX").Text)
                        
                        .GetText .MaxCols, plngLinhaItemPai, varItemCaixaAux
                        varItemCaixaAux = Left$(varItemCaixaAux, ((pintNivelItemDesejado - 1) * 3) + 1)
                        
                        If intNivelItemCaixa = pintNivelItemDesejado And _
                            varItemCaixaAux = Left$(xmlDomNodeAux.selectSingleNode("CO_ITEM_CAIX").Text, ((pintNivelItemDesejado - 1) * 3) + 1) Then
                            
                            blnTemFilhos = IIf(xmlDomNodeAux.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar, False, True)
                            Call flInserirLinhaSpread(pintNivelItemDesejado + 1, lngLinhaGrid, blnTemFilhos)
                            
                            .SetText pintNivelItemDesejado + 2, lngLinhaGrid, xmlDomNodeAux.selectSingleNode("DE_ITEM_CAIX").Text
                            .SetText .MaxCols, lngLinhaGrid, xmlDomNodeAux.selectSingleNode("CO_ITEM_CAIX").Text
                            
                            lngLinhaGrid = lngLinhaGrid + 1
                            intLinhasAdicionadas = intLinhasAdicionadas + 1
                            
                        End If
                        
                    End If
                Next
                
                If intCodGrupoVeicLegal < Val(xmlDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                    Exit For
                End If
                        
            Next
            
            lngLinhaSdoFinal = lngLinhaSdoFinal + intLinhasAdicionadas
            
            '>>>>> Atribui o movimento ao Item de Caixa e à Data corretos
            For Each xmlDomNode In xmlDomLeitura.selectNodes("//Repeat_MovAgrupItemCaixa/*")
                Call flLocalizarCelulaMovimentoSpread(xmlDomNode, _
                                                      pintNivelItemDesejado, _
                                                      plngLinhaItemPai + 1, _
                                                      plngLinhaItemPai + intLinhasAdicionadas)
            Next
        End If
        
        For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
            If .MaxTextColWidth(lngColunaGrid) > .ColWidth(lngColunaGrid) Then
                .ColWidth(lngColunaGrid) = .MaxTextColWidth(lngColunaGrid) + 200
            End If
        Next
        
        Call flIdentificarValoresNegativos
        
        .Redraw = True
    End With
            
    intUltPesqEfetuada = PorNiveisItemCaixa
    Set objUltNodeSel = pobjNodeSel
    datUltDataBase = pdatDataBase
    intUltPaginacao = pintPaginacao
    intUltNivelDesejado = pintNivelItemDesejado
    
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flCarregarMovPorNiveisItemCaixa"

End Sub

' Carrega movimentação por veículo legal.

Private Sub flCarregarMovPorVeiculoLegal(ByVal pobjNodeSel As MSComctlLib.Node, _
                                         ByVal pdatDataBase As Date, _
                                         ByVal pintPaginacao As enumPaginacao)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaFuturo  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaFuturo  As A6MIU.clsMonitoracaoFluxoCaixaFuturo
#End If

Dim xmlDomNode                          As IXMLDOMNode
Dim xmlDomNodeAux                       As IXMLDOMNode
Dim xmlDomLeitura                       As MSXML2.DOMDocument40
Dim xmlItemCaixaGrupoVeiculoLegal       As MSXML2.DOMDocument40
Dim strRetLeitura                       As String
Dim blnItemGenerico                     As Boolean

Dim strCodVeicLegal                     As String
Dim intCodGrupoVeicLegal                As Integer
Dim strSiglaSistema                     As String
Dim strD0                               As String
Dim strDataBase                         As String
Dim intNivelItemCaixa                   As Integer

Dim lngLinhaGrid                        As Long
Dim lngColunaGrid                       As Long

Dim varItemCaixaAux                     As Variant
Dim varDataAux                          As Variant

Dim arrPrimKey()                        As String
Dim vntCodErro                          As Variant
Dim vntMensagemErro                     As Variant

On Error GoTo ErrorHandler
    
    With Me.vasLista
        .Redraw = False

        Call flInicializarVasLista
        .ColsFrozen = .ColsFrozen + 1
        
        .SetText 1, 1, pobjNodeSel.Text
        .SetText 2, 2, "Saldo Inicial"
        .SetText 2, 3, "Movimentação"
        
        .BlockMode = False
        .Col = 1
        .Row = 3
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignCenter
        .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture
    
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set xmlItemCaixaGrupoVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
        
        Set objMonitoracaoFluxoCaixaFuturo = fgCriarObjetoMIU("A6MIU.clsMonitoracaoFluxoCaixaFuturo")
        
        arrPrimKey = Split(pobjNodeSel.Key, "k_")
        intCodGrupoVeicLegal = arrPrimKey(1)
        strCodVeicLegal = arrPrimKey(2)
        strSiglaSistema = arrPrimKey(3)
        
        '>>>>> Obtenção da sistuação atual do Veículo Legal
        Call flObterD0SaldoInicialCaixaSubReserva(strCodVeicLegal, strSiglaSistema, datD0, dblSaldoInicialD0)
        
        strD0 = CStr(Format(datD0, "DD/MM/YYYY"))
        strDataBase = CStr(Format(pdatDataBase, "DD/MM/YYYY"))
        
        strRetLeitura = objMonitoracaoFluxoCaixaFuturo.ObterMovimentoPorVeiculoLegal(strCodVeicLegal, _
                                                                                     strSiglaSistema, _
                                                                                     1, _
                                                                                     strD0, _
                                                                                     strDataBase, _
                                                                                     intQuantDiasPeriodo, _
                                                                                     pintPaginacao, _
                                                                                     vntCodErro, _
                                                                                     vntMensagemErro)
        
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            '100 - Documento XML Inválido.
            lngCodigoErroNegocio = 100
            GoTo ErrorHandler
        End If
        
        '>>>>> População das datas de movimento no Spread
        lngColunaGrid = intColunaInicioValores
        
        For Each xmlDomNode In xmlDomLeitura.selectNodes("//DatasMovimento/*")
            
            .SetText lngColunaGrid, 1, Trim$(fgDtXML_To_Interface(xmlDomNode.firstChild.Text))
            lngColunaGrid = lngColunaGrid + 1
            If lngColunaGrid = .MaxCols Then Exit For
        
        Next
        
        .MaxCols = intColunaInicioValores + intQuantDiasPeriodo + 1
        .ColWidth(.MaxCols) = 0
        
        If Not xmlDomLeitura.documentElement.selectSingleNode("Repeat_MovAgrupItemCaixa") Is Nothing Then
            '>>>>> Atribui o Valor do Saldo Inicial D1 para a paginação de movimento corrente
            If Not xmlDomLeitura.documentElement.selectSingleNode("//Repeat_SaldoInicialMovimento/Grupo_SaldoInicialMovimento/VA_LIQU_OPER") Is Nothing Then
                dblSaldoInicialD1 = fgVlrXml_To_Decimal(xmlDomLeitura.documentElement.selectSingleNode("//Repeat_SaldoInicialMovimento/Grupo_SaldoInicialMovimento/VA_LIQU_OPER").Text)
            Else
                dblSaldoInicialD1 = 0
            End If
            
            lngLinhaGrid = 4
            
            '>>>>> População dos itens de caixa no Spread
            If strItemCaixaGrupoVeiculoLegal <> vbNullString Then
                Call xmlItemCaixaGrupoVeiculoLegal.loadXML(strItemCaixaGrupoVeiculoLegal)

                For Each xmlDomNode In xmlItemCaixaGrupoVeiculoLegal.selectNodes("/Repeat_ItensCaixa/*")
                    
                    If Not blnItemGenerico Then
                        blnItemGenerico = True
                    
                        .BlockMode = False
                        .Col = 2
                        .Row = lngLinhaGrid
                        .CellType = CellTypePicture
                        .TypePictCenter = True
                        .TypeHAlign = TypeHAlignCenter
                        
                        .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
                        .SetText 3, lngLinhaGrid, gstrItemGenerico
                        .SetText .MaxCols, lngLinhaGrid, xmlItensCaixa.selectSingleNode("Repeat_ItemCaixa/Grupo_ItemCaixa[DE_ITEM_CAIX='" & gstrItemGenerico & "']/CO_ITEM_CAIX").Text
                        lngLinhaGrid = lngLinhaGrid + 1
                    End If
                    
                    For Each xmlDomNodeAux In xmlItensCaixa.selectNodes("Repeat_ItemCaixa/Grupo_ItemCaixa[CO_ITEM_CAIX_NIVE_01='" & xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "']")
                        If intCodGrupoVeicLegal = Val(xmlDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                            intNivelItemCaixa = fgObterNivelItemCaixa(xmlDomNodeAux.selectSingleNode("CO_ITEM_CAIX").Text)
                            
                            If intNivelItemCaixa = 1 Then
                                .BlockMode = False
                                .Col = 2
                                .Row = lngLinhaGrid
                                .CellType = CellTypePicture
                                .TypePictCenter = True
                                .TypeHAlign = TypeHAlignCenter
                                
                                If xmlDomNodeAux.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar Then
                                    .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
                                Else
                                    .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture
                                End If
                                
                                .SetText 3, lngLinhaGrid, xmlDomNodeAux.selectSingleNode("DE_ITEM_CAIX").Text
                                .SetText .MaxCols, lngLinhaGrid, xmlDomNodeAux.selectSingleNode("CO_ITEM_CAIX").Text
                                lngLinhaGrid = lngLinhaGrid + 1
                            End If

                        End If
                    Next
                    
                    If intCodGrupoVeicLegal < Val(xmlDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                        Exit For
                    End If
                
                Next
                
            Else
            
                .BlockMode = False
                .Col = 2
                .Row = lngLinhaGrid
                .CellType = CellTypePicture
                .TypePictCenter = True
                .TypeHAlign = TypeHAlignCenter
                
                .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
                .SetText 3, lngLinhaGrid, gstrItemGenerico
                .SetText .MaxCols, lngLinhaGrid, xmlItensCaixa.selectSingleNode("Repeat_ItemCaixa/Grupo_ItemCaixa[DE_ITEM_CAIX='" & gstrItemGenerico & "']/CO_ITEM_CAIX").Text
                lngLinhaGrid = lngLinhaGrid + 1
            
            End If
            
            .SetText 2, lngLinhaGrid, "Saldo Final"
            lngLinhaSdoFinal = lngLinhaGrid
            
            '>>>>> População do movimento para as datas e itens de caixa corretos
            For Each xmlDomNode In xmlDomLeitura.selectNodes("//Repeat_MovAgrupItemCaixa/*")
                Call flLocalizarCelulaMovimentoSpread(xmlDomNode, _
                                                      1, _
                                                      4, _
                                                      .MaxRows)
            Next
        
            Call flTotalizarMovimento(pintPaginacao)
        
            For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
                If .MaxTextColWidth(lngColunaGrid) > .ColWidth(lngColunaGrid) Then
                    .ColWidth(lngColunaGrid) = .MaxTextColWidth(lngColunaGrid) + 200
                End If
            Next
        
        Else
            lngLinhaSdoFinal = 4
            .SetText 2, lngLinhaSdoFinal, "Saldo Final"
        
        End If
        
        .Redraw = True
    End With
            
    intUltPesqEfetuada = PorVeiculoLegal
    Set objUltNodeSel = pobjNodeSel
    datUltDataBase = pdatDataBase
    intUltPaginacao = pintPaginacao
    
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set xmlItemCaixaGrupoVeiculoLegal = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flCarregarMovPorVeiculoLegal"

End Sub

' Carrega treeview de grupos e veículos legais.

Private Sub flCarregarTrvGeral(ByRef pstrXMLDocFiltros As String)

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
    
    Call flInicializarVasLista
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Set objItemCaixa = fgCriarObjetoMIU("A6MIU.clsItemCaixa")
    Call xmlRetorno.loadXML(objItemCaixa.ObterRelacaoItensCaixaGrupoVeicLegal(pstrXMLDocFiltros, _
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
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flCarregarTrvGeral"

End Sub

' Identifica e colore valores negativos.

Private Sub flIdentificarValoresNegativos()

Dim lngLinhaGrid                            As Long
Dim lngColunaGrid                           As Long
Dim varConteudoCelula                       As Variant
Dim varData                                 As Variant

On Error GoTo ErrorHandler

    With Me.vasLista
        For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
            For lngLinhaGrid = 2 To lngLinhaSdoFinal
                .GetText lngColunaGrid, 1, varData
                .GetText lngColunaGrid, lngLinhaGrid, varConteudoCelula
                
                .BlockMode = False
                .Col = lngColunaGrid
                .Row = lngLinhaGrid
                
                If Left$(Trim(varConteudoCelula), 1) = "-" Then
                    .ForeColor = vbRed
                Else
                    .ForeColor = vbBlack
                End If
                
                If varConteudoCelula = vbNullString And varData <> vbNullString Then
                    .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface("0")
                End If
            Next
        Next
    End With

    Exit Sub

ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flIdentificarValoresNegativos", 0

End Sub

' Formatação inicial do grid de detalhe de movimento.

Private Sub flInicializarFlxDetalhe()

On Error GoTo ErrorHandler

    intUltPesqEfetuada = None
    
    With Me.flxDetalhe
        .Redraw = False

        .Rows = 0
        .Rows = MAX_FIXED_ROWS
        .FixedRows = 1
        .FixedCols = 0
        
        .Cols = 11
        .ColWidth(0) = 800
        .ColWidth(1) = 2500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 2150
        .ColWidth(7) = 1500
        .ColWidth(8) = 1500
        .ColWidth(9) = 1500
        .ColWidth(10) = 1500
        
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
        .TextMatrix(0, 9) = "Data Movimento"
        .TextMatrix(0, 10) = "Data Retorno"
        
        .Redraw = True
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInicializarFlxDetalhe", 0
            
End Sub

' Formatação inicial do grid de movimentação futura.

Private Sub flInicializarVasLista()

Dim lngColunas                              As Long

On Error GoTo ErrorHandler

    With Me.vasLista
        .Redraw = False

        .MaxRows = 0
        .MaxRows = 100
        .RowsFrozen = 3
        
        .MaxCols = 100
        .ColWidth(1) = 200
        .ColWidth(2) = 200
        .ColWidth(3) = 200
        .ColWidth(4) = 200
        .ColWidth(5) = 200
        .ColWidth(6) = 200
        .ColWidth(7) = 2000
        .ColWidth(8) = 15
        .ColsFrozen = 8
        
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
        .BackColor = vbBlack
        .ForeColor = vbWhite
        .RowHeight(1) = 300
        .FontSize = 10
        .FontBold = True
        .TypeHAlign = TypeHAlignLeft
        .BlockMode = False
        
        .BlockMode = True
        .Col = intColunaInicioValores
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = intColunaInicioValores
        .Row = 2
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
        
        For lngColunas = intColunaInicioValores To .MaxCols
            .ColWidth(lngColunas) = 1400
        Next
        
        .CursorStyle = CursorStyleArrow
        
        .Redraw = True
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInicializarVasLista", 0
            
End Sub

' Insere linhas ao grid de movimentação.

Private Sub flInserirLinhaSpread(ByVal plngColunaPicture As Long, _
                                 ByVal plngLinhaInserir As Long, _
                                 ByVal pblnTemFilhos As Boolean)

On Error GoTo ErrorHandler

    With Me.vasLista
        .BlockMode = False
        .Col = plngColunaPicture
        .Row = plngLinhaInserir
        .Action = ActionInsertRow
        
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignCenter
        
        If pblnTemFilhos Then
            .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture
        Else
            .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture
        End If
        
        .BlockMode = True
        .Col = 1
        .Row = plngLinhaInserir
        .Col2 = 7
        .Row2 = plngLinhaInserir
        .BackColorStyle = BackColorStyleOverVertGridOnly
        .BackColor = &H8000000E
        .BlockMode = False
    
        .BlockMode = True
        .Col = intColunaInicioValores
        .Row = plngLinhaInserir
        .Col2 = .MaxCols
        .Row2 = plngLinhaInserir
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInserirLinhaSpread", 0

End Sub
                                 
' Identifica uma determinada célular no grid de movimentação.

Private Sub flLocalizarCelulaMovimentoSpread(ByVal pxmlDomNode As IXMLDOMNode, _
                                             ByVal pintNivelItemDesejado As Integer, _
                                             ByVal plngLinhaInicioBusca As Long, _
                                             ByVal plngLinhaFinalBusca As Long)

Dim blnAchouData                            As Boolean
Dim blnAchouItem                            As Boolean
Dim varDataAux                              As Variant
Dim varItemCaixaAux                         As Variant
Dim lngColunaGrid                           As Long
Dim lngLinhaGrid                            As Long

    With Me.vasLista
        blnAchouData = False
        blnAchouItem = False
        
        For lngColunaGrid = intColunaInicioValores To .MaxCols
            .GetText lngColunaGrid, 1, varDataAux
            varDataAux = Format$(varDataAux, "yyyymmdd")
            If varDataAux = pxmlDomNode.selectSingleNode("DT_LIQU_OPER").Text Then
                blnAchouData = True
                Exit For
            End If
        Next
        
        For lngLinhaGrid = plngLinhaInicioBusca To plngLinhaFinalBusca
            .GetText .MaxCols, lngLinhaGrid, varItemCaixaAux
            varItemCaixaAux = Left$(varItemCaixaAux, (pintNivelItemDesejado * 3) + 1)
            If varItemCaixaAux = enumTipoCaixa.CaixaFuturo & pxmlDomNode.selectSingleNode("AGRUP_ITEM_CAIX").Text Then
                blnAchouItem = True
                Exit For
            End If
        Next
        
        If blnAchouItem And blnAchouData Then
            .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface(pxmlDomNode.selectSingleNode("VA_LIQU_OPER").Text)
        End If
    End With

End Sub

' Obtém D0 e saldo inicial de determinado veículo legal.

Private Sub flObterD0SaldoInicialCaixaSubReserva(ByVal pstrVeiculoLegal As String, _
                                                 ByVal pstrSiglaSistema As String, _
                                                 ByRef pdatDataD0 As Date, _
                                                 ByRef pdblSaldoInicialD0 As Double)

Dim xmlPosCaixa                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlPosCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlPosCaixa.loadXML(fgObterPosicaoCaixaSubReserva(pstrVeiculoLegal, pstrSiglaSistema))
    
    With xmlPosCaixa.documentElement
        If Not .selectSingleNode("//DT_CAIX_SUB_RESE_ATUAL") Is Nothing Then
            pdatDataD0 = fgDtXML_To_Date(.selectSingleNode("//DT_CAIX_SUB_RESE_ATUAL").Text)
        End If
        If Not .selectSingleNode("//VA_UTLZ_ABER_CAIX_SUB_RESE") Is Nothing Then
            pdblSaldoInicialD0 = fgVlrXml_To_Decimal(.selectSingleNode("//VA_UTLZ_ABER_CAIX_SUB_RESE").Text)
        End If
    End With
    
    Set xmlPosCaixa = Nothing
    Exit Sub

ErrorHandler:
    Set xmlPosCaixa = Nothing
    mdiSBR.uctLogErros.MostrarErros Err, "flObterD0SaldoInicialCaixaSubReserva"

End Sub

' Refresh da tela montada por níveis de item de caixa.

Private Sub flRefreshMovPorNiveisItemCaixa(ByVal pobjNodeSel As MSComctlLib.Node, _
                                           ByVal pdatDataBase As Date, _
                                           ByVal pintPaginacao As enumPaginacao, _
                                           ByVal pintNivelItemDesejado As Integer)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaFuturo  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaFuturo  As A6MIU.clsMonitoracaoFluxoCaixaFuturo
#End If

Dim xmlDomNode                          As IXMLDOMNode
Dim xmlDomLeitura                       As MSXML2.DOMDocument40
Dim strRetLeitura                       As String

Dim intCodGrupoVeicLegal                As Integer
Dim strCodVeicLegal                     As String
Dim strSiglaSistema                     As String
Dim intNivelItemCaixa                   As Integer
Dim intCountNiveis                      As Integer
Dim strItemCaixaAux                     As String
Dim strD0                               As String
Dim strDataBase                         As String
Dim lngColunaGrid                       As Long

Dim arrPrimKey()                        As String
Dim vntCodErro                          As Variant
Dim vntMensagemErro                     As Variant

On Error GoTo ErrorHandler
    
    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMonitoracaoFluxoCaixaFuturo = fgCriarObjetoMIU("A6MIU.clsMonitoracaoFluxoCaixaFuturo")
    
    arrPrimKey = Split(pobjNodeSel.Key, "k_")
    intCodGrupoVeicLegal = arrPrimKey(1)
    strCodVeicLegal = arrPrimKey(2)
    strSiglaSistema = arrPrimKey(3)
    
    With Me.vasLista
        .Redraw = False
        
        .BlockMode = True
        .Col = intColunaInicioValores
        .Row = intLinhaMovimentacao + 1
        .Col2 = .MaxCols - 1
        .Row2 = lngLinhaSdoFinal
        .Text = fgVlrXml_To_Interface("0")
        .BlockMode = False
        
        strD0 = CStr(Format(datD0, "DD/MM/YYYY"))
        strDataBase = CStr(Format(pdatDataBase, "DD/MM/YYYY"))
        
        For intCountNiveis = pintNivelItemDesejado To 1 Step -1
        
            strRetLeitura = objMonitoracaoFluxoCaixaFuturo.ObterMovimentoPorVeiculoLegal(strCodVeicLegal, _
                                                                                         strSiglaSistema, _
                                                                                         intCountNiveis, _
                                                                                         strD0, _
                                                                                         strDataBase, _
                                                                                         intQuantDiasPeriodo, _
                                                                                         pintPaginacao, _
                                                                                         vntCodErro, _
                                                                                         vntMensagemErro)
        
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If
            
            Call xmlDomLeitura.loadXML(strRetLeitura)
        
            If intCountNiveis = pintNivelItemDesejado Then
                '>>>>> Obtenção da sistuação atual do Veículo Legal
                Call flObterD0SaldoInicialCaixaSubReserva(strCodVeicLegal, strSiglaSistema, datD0, dblSaldoInicialD0)
                
                '>>>>> População das datas de movimento no Spread
                lngColunaGrid = intColunaInicioValores
                
                For Each xmlDomNode In xmlDomLeitura.selectNodes("//DatasMovimento/*")
                    
                    .SetText lngColunaGrid, 1, Trim$(fgDtXML_To_Interface(xmlDomNode.firstChild.Text))
                    lngColunaGrid = lngColunaGrid + 1
                    If lngColunaGrid = .MaxCols Then Exit For
                
                Next
        
                .MaxCols = intColunaInicioValores + intQuantDiasPeriodo + 1
                .ColWidth(.MaxCols) = 0
            
                '>>>>> Atribui o Valor do Saldo Inicial D1 para a paginação de movimento corrente
                If Not xmlDomLeitura.documentElement.selectSingleNode("//Repeat_SaldoInicialMovimento/Grupo_SaldoInicialMovimento/VA_LIQU_OPER") Is Nothing Then
                    dblSaldoInicialD1 = fgVlrXml_To_Decimal(xmlDomLeitura.documentElement.selectSingleNode("//Repeat_SaldoInicialMovimento/Grupo_SaldoInicialMovimento/VA_LIQU_OPER").Text)
                Else
                    dblSaldoInicialD1 = 0
                End If
                
            End If
        
            For Each xmlDomNode In xmlDomLeitura.selectNodes("//Repeat_MovAgrupItemCaixa/*")
                
                strItemCaixaAux = enumTipoCaixa.CaixaFuturo & _
                                  xmlDomNode.selectSingleNode("AGRUP_ITEM_CAIX").Text & _
                                  String$(15 - Len(xmlDomNode.selectSingleNode("AGRUP_ITEM_CAIX").Text), "0")
                
                intNivelItemCaixa = fgObterNivelItemCaixa(strItemCaixaAux)
                
                If intNivelItemCaixa = intCountNiveis Then
                    Call flLocalizarCelulaMovimentoSpread(xmlDomNode, _
                                                          intCountNiveis, _
                                                          intLinhaMovimentacao + 1, _
                                                          lngLinhaSdoFinal - 1)
                End If
                
            Next
        Next
        
        For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
            If .MaxTextColWidth(lngColunaGrid) > .ColWidth(lngColunaGrid) Then
                .ColWidth(lngColunaGrid) = .MaxTextColWidth(lngColunaGrid) + 200
            End If
        Next
        
        Call flTotalizarMovimento(pintPaginacao)
        
        .Redraw = True
    End With
            
    intUltPesqEfetuada = PorNiveisItemCaixa
    Set objUltNodeSel = pobjNodeSel
    datUltDataBase = pdatDataBase
    intUltPaginacao = pintPaginacao
    intUltNivelDesejado = pintNivelItemDesejado
    
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoFluxoCaixaFuturo = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flRefreshMovPorNiveisItemCaixa"

End Sub

' Refresh da posição atual da tela.

Private Sub flRefreshPosicaoAtualTela(Optional ByVal pstrPosicaoAtualPesquisa As String = vbNullString)

Dim xmlPosPesquisa                          As MSXML2.DOMDocument40

    Select Case intUltPesqEfetuada
        Case enumTipoPesquisa.PorDetalheLancamento
             Call flRefreshMovPorNiveisItemCaixa(objUltNodeSel, _
                                                 datUltDataBase, _
                                                 intUltPaginacao, _
                                                 intUltNivelDesejado)
             
             Call flCarregarDetalheMovimento(objUltNodeSel, _
                                             lngUltColunaData, _
                                             lngUltLinhaItem)
             Exit Sub
             
        Case enumTipoPesquisa.PorGrupoVeiculoLegal
             Call flCarregarMovPorGrupoVeiculoLegal(objUltNodeSel, _
                                                    datUltDataBase, _
                                                    intUltPaginacao)
             Exit Sub
        
        Case enumTipoPesquisa.PorNiveisItemCaixa
             Call flRefreshMovPorNiveisItemCaixa(objUltNodeSel, _
                                                 datUltDataBase, _
                                                 intUltPaginacao, _
                                                 intUltNivelDesejado)
             Exit Sub

        Case enumTipoPesquisa.PorVeiculoLegal
             Call flRefreshMovPorNiveisItemCaixa(objUltNodeSel, _
                                                 datUltDataBase, _
                                                 intUltPaginacao, _
                                                 1)
             Exit Sub
    
    End Select
    
    If pstrPosicaoAtualPesquisa <> vbNullString Then
        Set xmlPosPesquisa = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call xmlPosPesquisa.loadXML(pstrPosicaoAtualPesquisa)
        
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
            
            intUltPesqEfetuada = None
        End With
        
        Set xmlPosPesquisa = Nothing
    End If
    
End Sub

' Seleciona uma determinada célula no grid de movimentação.

Private Function flSelecionarCelulaSpread(ByVal plngCol As Long, _
                                          ByVal plngRow As Long) As Boolean

Dim varConteudoCelula                       As Variant
Dim lngColunaGrid                           As Long

On Error GoTo ErrorHandler

    flSelecionarCelulaSpread = True
    
    With Me.vasLista
        
        If plngCol < intColunaInicioValores Then
            flSelecionarCelulaSpread = False
            Exit Function
        End If
        
        .GetText plngCol, 1, varConteudoCelula
        If varConteudoCelula = vbNullString Then
            flSelecionarCelulaSpread = False
            Exit Function
        End If
        
        .BlockMode = False
        .Row = plngRow
        lngColunaGrid = 1
        
        Do
            
            .Col = lngColunaGrid
            If Not .TypePictPicture Is Nothing Then
                If .TypePictPicture = imgOutrosIcones.ListImages("leaf").Picture Then Exit Do
            End If
            lngColunaGrid = lngColunaGrid + 1
            
            If lngColunaGrid >= intColunaInicioValores Then
                flSelecionarCelulaSpread = False
                Exit Function
            End If
            
        Loop
        
        lngColunaGrid = lngColunaGrid + 1
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .ForeColor = vbWhite
        
        .Col = lngColunaGrid
        .Row = plngRow
        .Col2 = 7
        .Row2 = plngRow
        .BackColor = vbYellow
    
        .BlockMode = False
        .Col = plngCol
        .Row = plngRow
        .BackColor = vbYellow
    
        .Col = plngCol
        .Row = 1
        .ForeColor = vbYellow
    
    End With

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flSelecionarCelulaSpread", 0
    
End Function

' Totaliza movimentação por data.

Private Sub flTotalizarMovimento(ByVal pintPaginacao As enumPaginacao)
                                 
Dim lngLinhaGrid                            As Long
Dim lngColunaGrid                           As Long
Dim intNivelItemCaixa                       As Integer

Dim varValorItemCaixa                       As Variant
Dim varCodItemCaixa                         As Variant
Dim varData                                 As Variant

Dim dblTotal                                As Double

On Error GoTo ErrorHandler

    With Me.vasLista
        For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
            For lngLinhaGrid = intLinhaSdoInicial To lngLinhaSdoFinal
                
                .GetText lngColunaGrid, 1, varData
                .GetText lngColunaGrid, lngLinhaGrid, varValorItemCaixa
                
                If varValorItemCaixa = vbNullString And varData <> vbNullString Then
                    .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface("0")
                End If
            Next
        Next
        
        For lngColunaGrid = intColunaInicioValores To .MaxCols - 1
            dblTotal = 0
            For lngLinhaGrid = intLinhaMovimentacao + 1 To lngLinhaSdoFinal - 1
                .GetText .MaxCols, lngLinhaGrid, varCodItemCaixa
                intNivelItemCaixa = fgObterNivelItemCaixa(varCodItemCaixa)
                
                If intNivelItemCaixa = 1 Then
                    .GetText lngColunaGrid, lngLinhaGrid, varValorItemCaixa
                    dblTotal = dblTotal + fgVlrXml_To_Decimal(varValorItemCaixa)
                End If
            Next
            .SetText lngColunaGrid, intLinhaMovimentacao, fgVlrXml_To_Interface(dblTotal)
        Next
        
        Call flCalcularSaldosInicialEFinal
        Call flIdentificarValoresNegativos
        
        .BlockMode = True
        .Col = 1
        .Row = intLinhaSdoInicial
        .Col2 = .MaxCols
        .Row2 = intLinhaSdoInicial
        .FontBold = True
        
        .Col = 1
        .Row = intLinhaMovimentacao
        .Col2 = .MaxCols
        .Row2 = intLinhaMovimentacao
        .FontBold = True
        
        .Col = 1
        .Row = lngLinhaSdoFinal
        .Col2 = .MaxCols
        .Row2 = lngLinhaSdoFinal
        .FontBold = True
        .BlockMode = False
    End With

    Exit Sub

ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flTotalizarMovimento", 0

End Sub
                                 
Private Sub ctlTableCombo_AplicarFiltro(xmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCursor True
    Call flInicializarFlxDetalhe
    
    If Trim(xmlFiltro) = vbNullString Then
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
        Call xmlDomFiltro.loadXML(xmlFiltro)
        
        If Not xmlDomFiltro.selectSingleNode("//Grupo_BancoLiquidante") Is Nothing Then
            Call fgRemoveNode(xmlDomFiltro, "Grupo_BancoLiquidante")
        End If
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", _
                                        xmlDomLeitura.selectSingleNode("//BancoLiquidante").Text)
                                        
        Call flCarregarTrvGeral(xmlDomFiltro.xml)
        xmlFiltro = xmlDomFiltro.xml
    End If
    
    fgCursor
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - ctlTableCombo_AplicarFiltro"
    
End Sub

Private Sub ctlTableCombo_DropDown()

On Error GoTo ErrorHandler
    
    lngAlturaTableCombo = ctlTableCombo.Height
    Call ctlTableCombo.fgCarregarCombo
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - ctlTableCombo_DropDown"

End Sub

Private Sub ctlTableCombo_MouseMove()

On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButton3D
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - ctlTableCombo_MouseMove"

End Sub

Private Sub flxDetalhe_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButtonFlat
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - flxDetalhe_MouseMove"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        fgCursor True
        Call flRefreshPosicaoAtualTela
        fgCursor
    End If
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    intUltPesqEfetuada = None
    
    fgCursor True
    Call fgCenterMe(Me)
    
    Me.Icon = mdiSBR.Icon
    
    intTipoBackOfficeUsuario = fgObterTipoBackOfficeUsuario
    intQuantDiasPeriodo = 5
    intUltNivelDesejado = 1
    
    Call flInicializarVasLista
    Call flInicializarFlxDetalhe
        
    lblData.Caption = fgDataHoraServidor(enumFormatoDataHora.Data)
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
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmCaixaFuturo
    Load objFiltro
    
    Call flCarregarListaItensCaixa
    Call objFiltro.fgCarregarPesquisaAnterior
    
    DoEvents
    Me.Show
    
    fgCursor

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - Form_Load"

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
        .vasLista.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .trvGeral.Top - .tlbPaginacao.Height, .tlbButtons.Top - .trvGeral.Top - .tlbPaginacao.Height)
        .vasLista.Width = .ScaleWidth - .vasLista.Left
        
        .tlbPaginacao.Left = .vasLista.Left
        .tlbPaginacao.Top = .vasLista.Top + .vasLista.Height
        .tlbPaginacao.Width = .vasLista.Width
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo ErrorHandler
    
    Set xmlItensCaixa = Nothing
    Unload objBuscaNo
    Set objBuscaNo = Nothing
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - Form_Unload"

End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    fblnDummyH = True

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - imgDummyH_MouseDown"

End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButtonFlat
    
    If Not fblnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyH.Top = y + imgDummyH.Top

    On Error Resume Next
    
    With Me
        If .imgDummyH.Top < 1926 Then
            .imgDummyH.Top = 1926
        End If
        If .imgDummyH.Top > (.Width - 500) And (.Width - 500) > 0 Then
            .imgDummyH.Top = .Width - 500
        End If
        
        .flxDetalhe.Top = .imgDummyH.Top + .imgDummyH.Height
        .flxDetalhe.Height = .tlbButtons.Top - .imgDummyH.Top - .imgDummyH.Height
        .trvGeral.Height = .imgDummyH.Top - .trvGeral.Top - .txtDetalheVeicLega.Height
        .txtDetalheVeicLega.Top = .imgDummyH.Top - .txtDetalheVeicLega.Height
        .vasLista.Height = .trvGeral.Height - .tlbPaginacao.Height + .txtDetalheVeicLega.Height
        .tlbPaginacao.Top = .vasLista.Top + .vasLista.Height
    End With

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - imgDummyH_MouseMove"

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    fblnDummyH = False

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - imgDummyH_MouseUp"

End Sub

Private Sub imgDummyV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    fblnDummyV = True

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - imgDummyV_MouseDown"

End Sub

Private Sub imgDummyV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButtonFlat
    
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
        .tlbPaginacao.Left = .vasLista.Left
        .tlbPaginacao.Width = .vasLista.Width
    End With

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - imgDummyV_MouseMove"

End Sub

Private Sub imgDummyV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    fblnDummyV = False

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - imgDummyV_MouseUp"

End Sub

Private Sub lblBarra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - lblBarra_MouseMove"

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)
    
On Error GoTo ErrorHandler
    
    fgCursor True
    
    tlbButtons.Buttons("aplicarfiltro").Value = tbrPressed
    Call flCarregarTrvGeral(xmlDocFiltros)
    xmlFiltro = xmlDocFiltros
    If InStr(1, xmlFiltro, "BackOfficePerfilGeral") <> 0 Then
        Call flCarregarListaItensCaixa
    End If
    Call flInicializarFlxDetalhe
    Call flRefreshPosicaoAtualTela(xmlFiltro)
    
    fgCursor

    ctlTableCombo.TituloCombo = IIf(strTituloTableCombo = vbNullString, strTableComboInicial, strTituloTableCombo)
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - objFiltro_AplicarFiltro"

End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strJanelas                              As String

On Error GoTo ErrorHandler
    
    If tlbButtons.Buttons("showtreeview").Value = tbrPressed Then
        strJanelas = strJanelas & "1"
    End If
    
    If tlbButtons.Buttons("showlist").Value = tbrPressed Then
        strJanelas = strJanelas & "2"
    End If
    
    If tlbButtons.Buttons("showdetail").Value = tbrPressed Then
        strJanelas = strJanelas & "3"
    End If
    
    Call flArranjarJanelasExibicao(strJanelas)
    
    Select Case Button.Key
    Case "showfiltro"
        Set objFiltro = New frmFiltro
        Set objFiltro.FormOwner = Me
        objFiltro.TipoFiltro = enumTipoFiltroA6.frmCaixaFuturo
        objFiltro.Show vbModal
        
    Case "refresh"
        fgCursor True
        Call flRefreshPosicaoAtualTela
        fgCursor
    
    End Select

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - tlbButtons_ButtonClick"

End Sub

Private Sub tlbPaginacao_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim varDataBase                             As Variant
Dim intPaginacao                            As enumPaginacao

On Error GoTo ErrorHandler
    
    If trvGeral.SelectedItem Is Nothing Then Exit Sub
    
    intPaginacao = IIf(Button.Key = "proximo", enumPaginacao.Proximo, enumPaginacao.Anterior)
    
    If intPaginacao = enumPaginacao.Proximo Then
        vasLista.GetText vasLista.MaxCols - 1, 1, varDataBase
    Else
        If intUltPesqEfetuada = PorGrupoVeiculoLegal Then
            vasLista.GetText intColunaInicioValores, 1, varDataBase
        Else
            vasLista.GetText intColunaInicioValores + 2, 1, varDataBase
        End If
    End If
    
    If Not IsDate(varDataBase) Then varDataBase = fgDataHoraServidor(enumFormatoDataHora.Data)
    If CDate(varDataBase) - intQuantDiasPeriodo <= fgDataHoraServidor(enumFormatoDataHora.Data) Then Exit Sub
    
    fgCursor True
    
    Select Case intUltPesqEfetuada
        Case enumTipoPesquisa.PorGrupoVeiculoLegal
             Call flCarregarMovPorGrupoVeiculoLegal(trvGeral.SelectedItem, varDataBase, intPaginacao)
    
        Case enumTipoPesquisa.PorVeiculoLegal
             Call flRefreshMovPorNiveisItemCaixa(trvGeral.SelectedItem, varDataBase, intPaginacao, 1)
    
        Case Else
             Call fgSelecionarLinhaSpread(Me.vasLista, 1, 3)
             Call flInicializarFlxDetalhe
             If intUltNivelDesejado > 0 Then
                 Call flRefreshMovPorNiveisItemCaixa(trvGeral.SelectedItem, varDataBase, intPaginacao, intUltNivelDesejado)
             Else
                 Call flCarregarMovPorVeiculoLegal(trvGeral.SelectedItem, varDataBase, intPaginacao)
             End If
    
    End Select
        
    fgCursor
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - tlbPaginacao_ButtonClick"

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
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - trvGeral_MouseDown"

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
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - trvGeral_MouseMove"

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
    Call flInicializarFlxDetalhe
    
    If InStr(2, Node.Key, "k_") = 0 Then
        Call flCarregarMovPorGrupoVeiculoLegal(Node, fgDataHoraServidor(enumFormatoDataHora.Data), enumPaginacao.Proximo)
    Else
        Call flCarregarMovPorVeiculoLegal(Node, fgDataHoraServidor(enumFormatoDataHora.Data), enumPaginacao.Proximo)
    End If
    
    fgCursor
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - trvGeral_NodeClick"

End Sub

Private Sub vasLista_Click(ByVal Col As Long, ByVal Row As Long)

Dim lngLinhaGrid                            As Long
Dim varConteudoCelula                       As Variant
Dim varRadicalItemPai                       As Variant
Dim intNivelItemPai                         As Integer
Dim intNivelItemCaixa                       As Integer
Dim varDataInicioAux                        As Variant

On Error GoTo ErrorHandler
    
    Call flInicializarFlxDetalhe
    
    With vasLista
        .BlockMode = False
        .Col = Col
        .Row = Row

        If Not .TypePictPicture Is Nothing Then
            .GetText .MaxCols, .Row, varConteudoCelula
            If varConteudoCelula = vbNullString Then
                intNivelItemPai = 0
                varRadicalItemPai = "2"
            Else
                intNivelItemPai = fgObterNivelItemCaixa(varConteudoCelula)
                varRadicalItemPai = Left$(varConteudoCelula, (intNivelItemPai * 3) + 1)
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
                    If lngLinhaGrid >= lngLinhaSdoFinal Then Exit Do
                Loop
                    
            ElseIf .TypePictPicture = imgOutrosIcones.ListImages("treeplus").Picture Then
                .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture

                If .RowHeight(.Row + 1) <> 0 Then
                    If Not Me.trvGeral.SelectedItem Is Nothing Then
                        .GetText intColunaInicioValores + 1, 1, varDataInicioAux
                        If varDataInicioAux = vbNullString Then
                            .GetText intColunaInicioValores, 1, varDataInicioAux
                        End If
                        
                        If varDataInicioAux = vbNullString Then Exit Sub
                        fgCursor True
                        Call flCarregarMovPorNiveisItemCaixa(Me.trvGeral.SelectedItem, varDataInicioAux, enumPaginacao.Proximo, .Col, .Row)
                        fgCursor
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
                        If lngLinhaGrid >= lngLinhaSdoFinal Then Exit Do
                    Loop
                End If
                
            End If
        
        Else
            If Row >= intLinhaSdoInicial And Row <= lngLinhaSdoFinal Then
                Call fgSelecionarLinhaSpread(Me.vasLista, Col, Row)
            End If
        
        End If

    End With
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - vasLista_Click"

End Sub

Private Sub vasLista_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim varTextoSpread                          As Variant
Dim objNodeBusca                            As MSComctlLib.Node
    
On Error GoTo ErrorHandler
    
    If flSelecionarCelulaSpread(Col, Row) Then
        If Not Me.trvGeral.SelectedItem Is Nothing Then
            fgCursor True
            Call flCarregarDetalheMovimento(trvGeral.SelectedItem, Col, Row)
            fgCursor
        End If
    Else
        vasLista.GetText 2, Row, varTextoSpread
        If varTextoSpread <> vbNullString And Row > 1 Then
            Set objNodeBusca = objBuscaNo.ProcuraNoProx(varTextoSpread)
            If Not objNodeBusca Is Nothing Then
                objNodeBusca.EnsureVisible
                objNodeBusca.Selected = True
                Call trvGeral_NodeClick(objNodeBusca)
            End If
        End If
    End If

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - vasLista_DblClick"

End Sub

Private Sub vasLista_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler
        
    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmCaixaFuturo - vasLista_MouseMove"

End Sub
