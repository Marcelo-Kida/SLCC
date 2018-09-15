VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubReservaMonitoracao 
   Caption         =   "Sub-reserva - Monitoração de Movimentação"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   14040
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDetalheVeicLega 
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3795
   End
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   661
   End
   Begin FPSpread.vaSpread vasLista 
      Height          =   1245
      Left            =   3960
      TabIndex        =   0
      Tag             =   "Totais por Local Liquidação"
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
      SpreadDesigner  =   "frmSubReservaMonitoracao.frx":0000
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":01E8
            Key             =   "itemgrupo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":053A
            Key             =   "itemelementar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":088C
            Key             =   "treeminus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":0BAE
            Key             =   "treeplus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5055
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   635
      ButtonWidth     =   2646
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Árvore"
            Key             =   "showtreeview"
            Object.ToolTipText     =   "Mostrar TreeView"
            ImageIndex      =   3
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Lista"
            Key             =   "showlist"
            Object.ToolTipText     =   "Mostrar Lista"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Detalhe"
            Key             =   "showdetail"
            Object.ToolTipText     =   "Mostrar Detalhe"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "refresh"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvGeral 
      Height          =   3555
      Left            =   30
      TabIndex        =   2
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
            Picture         =   "frmSubReservaMonitoracao.frx":0ED0
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":0FE2
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":10F4
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":1446
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":1798
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":1AEA
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":1E3C
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaMonitoracao.frx":228E
            Key             =   "posterior"
         EndProperty
      EndProperty
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
      TabIndex        =   4
      Top             =   30
      Width           =   300
   End
   Begin VB.Label lblBarra 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   3390
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Image imgDummyV 
      Height          =   3465
      Left            =   3855
      MousePointer    =   9  'Size W E
      Top             =   435
      Width           =   90
   End
End
Attribute VB_Name = "frmSubReservaMonitoracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a consulta do total de movimento de caixa
' por local de liquidação.

Option Explicit

Private WithEvents objBuscaNo               As frmBuscaNo
Attribute objBuscaNo.VB_VarHelpID = -1
Private blnEventByPass                      As Boolean

Private strCarregaTreeView                  As String

Private fbDummyV                            As Boolean
Private lngMaxColsGrid                      As Long
Private intTipoBackOfficeUsuario            As enumTipoBackOffice

Private Const MAX_FIXED_ROWS                As Integer = 100

' Arranja janelas de exibição conforme seleção do usuário.

Private Sub flArranjarJanelasExibicao(ByVal psJanelas As String)

On Error GoTo ErrorHandler

    Select Case psJanelas
    Case ""
        imgDummyV.Visible = False
        trvGeral.Visible = False
        vasLista.Visible = False
    
    Case "1"
        imgDummyV.Visible = False
        trvGeral.Visible = True
        vasLista.Visible = False
    
    Case "2"
        imgDummyV.Visible = False
        trvGeral.Visible = False
        vasLista.Visible = True
    
    Case "12"
        imgDummyV.Visible = True
        trvGeral.Visible = True
        vasLista.Visible = True
    
    End Select
    
    txtDetalheVeicLega.Visible = trvGeral.Visible
    Call Form_Resize

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flArranjarJanelasExibicao", 0
    
End Sub

' Carrega totalização de movimento por grupo de veículo legal.

Private Sub flCarregarTotPorGrupoVeiculoLegal(ByVal objNodeSel As MSComctlLib.Node)

#If EnableSoap = 1 Then
    Dim objMonitoracaoMovimentacaoFundos    As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoMovimentacaoFundos    As A6MIU.clsMonitoracaoMovimentacaoFundos
#End If

Dim objDomNode                              As IXMLDOMNode
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String

Dim lngLinhaGrid                            As Integer
Dim lngColunaGrid                           As Integer

Dim strGrupVeicLega                         As String
Dim arrPrimKey()                            As String

Dim dblTotalVeiculo                         As Double
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    With Me.vasLista
        .Redraw = False
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMonitoracaoMovimentacaoFundos = fgCriarObjetoMIU("A6MIU.clsMonitoracaoMovimentacaoFundos")
        
        arrPrimKey = Split(objNodeSel.Key, "k_")
        strGrupVeicLega = Trim(arrPrimKey(1))
        
        strRetLeitura = objMonitoracaoMovimentacaoFundos.ObterTotalPorGrupoVeiculoLegal(strGrupVeicLega, _
                                                                                        vntCodErro, _
                                                                                        vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        .MaxCols = 0
        If strRetLeitura <> vbNullString Then
            flInicializarVasLista
            
            If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                '100 - Documento XML Inválido.
                lngCodigoErroNegocio = 100
                GoTo ErrorHandler
            End If
            
            lngMaxColsGrid = xmlDomLeitura.selectNodes("/Repeat_TotalGrupoVeiculoLegal/*").length + 1
            .MaxCols = lngMaxColsGrid
            
            lngColunaGrid = 1
            For Each objDomNode In xmlDomLeitura.selectNodes("/Repeat_TotalGrupoVeiculoLegal/*")
                .SetText lngColunaGrid, 2, Trim$(objDomNode.selectSingleNode("SG_LOCA_LIQU").Text)
                lngColunaGrid = lngColunaGrid + 1
            Next
        
            lngLinhaGrid = 3
            lngColunaGrid = 1
            For Each objDomNode In xmlDomLeitura.selectNodes("/Repeat_TotalGrupoVeiculoLegal/*")
                .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
                dblTotalVeiculo = dblTotalVeiculo + fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
                lngColunaGrid = lngColunaGrid + 1
            Next
            .SetText .MaxCols, lngLinhaGrid, fgVlrXml_To_Interface(dblTotalVeiculo)
            .SetText .MaxCols, 2, "Total"
            
            .BlockMode = False
            .Row = 3
            For lngColunaGrid = 1 To .MaxCols
                .Col = lngColunaGrid
                If .MaxTextCellWidth > .ColWidth(lngColunaGrid) Then
                    .ColWidth(lngColunaGrid) = .MaxTextCellWidth + 200
                End If
            Next
        End If
        
        .Redraw = True
    End With
            
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoMovimentacaoFundos = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoMovimentacaoFundos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - flCarregarTotPorGrupoVeiculoLegal"

End Sub

' Carrega totalização de movimento por veículo legal.

Private Sub flCarregarTotPorVeiculoLegal(ByVal objNodeSel As MSComctlLib.Node)

#If EnableSoap = 1 Then
    Dim objMonitoracaoMovimentacaoFundos    As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoMovimentacaoFundos    As A6MIU.clsMonitoracaoMovimentacaoFundos
#End If

Dim objDomNode                              As IXMLDOMNode
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String

Dim lngLinhaGrid                            As Integer
Dim lngColunaGrid                           As Integer
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String

Dim dblTotalVeiculo                         As Double
Dim arrPrimKey()                            As String

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    With Me.vasLista
        .Redraw = False
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMonitoracaoMovimentacaoFundos = fgCriarObjetoMIU("A6MIU.clsMonitoracaoMovimentacaoFundos")
        
        arrPrimKey = Split(objNodeSel.Key, "k_")
        strCodigoVeiculoLegal = arrPrimKey(2)
        strSiglaSistema = arrPrimKey(3)
        
        strRetLeitura = objMonitoracaoMovimentacaoFundos.ObterTotalPorVeiculoLegal(strCodigoVeiculoLegal, _
                                                                                   strSiglaSistema, _
                                                                                   vntCodErro, _
                                                                                   vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        .MaxCols = 0
        If strRetLeitura <> vbNullString Then
            flInicializarVasLista
            
            If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                '100 - Documento XML Inválido.
                lngCodigoErroNegocio = 100
                GoTo ErrorHandler
            End If
            
            lngMaxColsGrid = xmlDomLeitura.selectNodes("/Repeat_TotalVeiculoLegal/*").length + 1
            .MaxCols = lngMaxColsGrid
            
            lngColunaGrid = 1
            For Each objDomNode In xmlDomLeitura.selectNodes("/Repeat_TotalVeiculoLegal/*")
                .SetText lngColunaGrid, 2, Trim$(objDomNode.selectSingleNode("SG_LOCA_LIQU").Text)
                lngColunaGrid = lngColunaGrid + 1
            Next
        
            lngLinhaGrid = 3
            lngColunaGrid = 1
            For Each objDomNode In xmlDomLeitura.selectNodes("/Repeat_TotalVeiculoLegal/*")
                .SetText lngColunaGrid, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
                dblTotalVeiculo = dblTotalVeiculo + fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
                lngColunaGrid = lngColunaGrid + 1
            Next
            
            .SetText .MaxCols, lngLinhaGrid, fgVlrXml_To_Interface(dblTotalVeiculo)
            .SetText .MaxCols, 2, "Total"
            
            .BlockMode = False
            .Row = 3
            For lngColunaGrid = 1 To .MaxCols
                .Col = lngColunaGrid
                If .MaxTextCellWidth > .ColWidth(lngColunaGrid) Then
                    .ColWidth(lngColunaGrid) = .MaxTextCellWidth + 200
                End If
            Next
        End If
        
        .Redraw = True
    End With
            
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoMovimentacaoFundos = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set objMonitoracaoMovimentacaoFundos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - flCarregarTotPorVeiculoLegal"

End Sub

' Carrega treeview de grupos e veículos legais.

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
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Set objItemCaixa = fgCriarObjetoMIU("A6MIU.clsItemCaixa")
    Call xmlRetorno.loadXML(objItemCaixa.ObterRelacaoItensCaixaGrupoVeicLegal(xmlDocFiltros, _
                                                                              False, _
                                                                              0, _
                                                                              vntCodErro, _
                                                                              vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
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
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - flCarregarTrvGeral"

End Sub

' Inicialização do grid de pesquisa de totalização de movimento.

Private Sub flInicializarVasLista()

Dim lngColunas                              As Long

On Error GoTo ErrorHandler

    With Me.vasLista
        .Redraw = False

        .MaxRows = 0
        .MaxRows = MAX_FIXED_ROWS
        .RowsFrozen = 3
        
        .MaxCols = 100 'lngMaxColsGrid
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .AllowCellOverflow = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
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
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .TypeHAlign = TypeHAlignLeft
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 2
        .Col2 = .MaxCols
        .Row2 = 2
        .BackColor = &H8000000F
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        For lngColunas = 1 To .MaxCols
            .ColWidth(lngColunas) = 1700
        Next
        
        .SetText 1, 1, "Saldo Total da Posição"
        
        .BlockMode = True
        .Col = 1
        .Row = 3
        .Col2 = lngMaxColsGrid
        .Row2 = .MaxRows
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
        
        
        .Redraw = True
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInicializarVasLista", 0
            
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        If trvGeral.SelectedItem Is Nothing Then Exit Sub
        
        fgCursor True
        Call trvGeral_NodeClick(trvGeral.SelectedItem)
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - Form_KeyDown"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    Me.Icon = mdiSBR.Icon

    fgCursor True

    flArranjarJanelasExibicao ("12")
    
    ctlTableCombo.TituloCombo = "Grupos de Veículos Legais"

    Call fgCenterMe(Me)
        
    intTipoBackOfficeUsuario = fgObterTipoBackOfficeUsuario
    Call flCarregarTrvGeral(vbNullString)
    
    Call flInicializarVasLista
    vasLista.MaxCols = 0
    
    '>>>>> Inicialização Formulário de Busca
    Set objBuscaNo = New frmBuscaNo
    Load objBuscaNo
    objBuscaNo.Criterio = "Veículo Legal"
    Set objBuscaNo.objTreeView = trvGeral
    
    DoEvents
    Me.Show
    
    vasLista.CursorStyle = CursorStyleArrow
    txtDetalheVeicLega.Visible = True
    
    fgCursor

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - Form_Load"

End Sub

Private Sub Form_Resize()
    
On Error Resume Next

    With Me
        .ctlTableCombo.Left = 0
        .ctlTableCombo.Top = 0
        
        .tlbButtons.Top = .ScaleHeight - .tlbButtons.Height
        
        .trvGeral.Left = 0
        .trvGeral.Top = .ctlTableCombo.Height
        .trvGeral.Height = .tlbButtons.Top - .trvGeral.Top - .txtDetalheVeicLega.Height
        .trvGeral.Width = IIf(.imgDummyV.Visible, .imgDummyV.Left, .ScaleWidth)
        
        .txtDetalheVeicLega.Left = 0
        .txtDetalheVeicLega.Top = .trvGeral.Top + .trvGeral.Height
        .txtDetalheVeicLega.Width = .trvGeral.Width
        
        .lblBarra.Left = 0
        .lblBarra.Top = 0
        .lblBarra.Width = .ScaleWidth
        
        .lblData.Left = .ScaleWidth - .lblData.Width
        .lblData.Top = 0
        
        .imgDummyV.Top = 0
        .imgDummyV.Height = .trvGeral.Height
        
        .vasLista.Left = IIf(.imgDummyV.Visible, .imgDummyV.Left + .imgDummyV.Width, 0)
        .vasLista.Top = .trvGeral.Top
        .vasLista.Height = .tlbButtons.Top - .trvGeral.Top
        .vasLista.Width = .ScaleWidth - .vasLista.Left
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo ErrorHandler
    
    Set frmSubReservaMonitoracao = Nothing
    Unload objBuscaNo
    Set objBuscaNo = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_Unload"

End Sub

Private Sub imgDummyV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    fbDummyV = True

End Sub

Private Sub imgDummyV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
If Not fbDummyV Or Button = vbRightButton Then
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

Private Sub imgDummyV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    fbDummyV = False

End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim lsJanelas                               As String

On Error GoTo ErrorHandler

    If tlbButtons.Buttons("showtreeview").Value = tbrPressed Then
        lsJanelas = lsJanelas & "1"
    End If
    
    If tlbButtons.Buttons("showlist").Value = tbrPressed Then
        lsJanelas = lsJanelas & "2"
    End If
    
    Call flArranjarJanelasExibicao(lsJanelas)

    Select Case Button.Key
    Case "refresh"
        If trvGeral.SelectedItem Is Nothing Then Exit Sub
        
        fgCursor True
        Call trvGeral_NodeClick(trvGeral.SelectedItem)
        fgCursor
        
    End Select

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - tlbButtons_ButtonClick"
    
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
    
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - trvGeral_MouseDown"

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

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - trvGeral_MouseMove"
   
End Sub

Private Sub trvGeral_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnEventByPass = False
    
End Sub

Private Sub trvGeral_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim arrNodeKey()                            As String
    
If blnEventByPass Then Exit Sub

On Error GoTo ErrorHandler
    
    objBuscaNo.Hide
    DoEvents
    
    fgCursor True
    
    If InStr(2, Node.Key, "k_") = 0 Then
        Call flCarregarTotPorGrupoVeiculoLegal(Node)
    Else
        arrNodeKey = Split(Node.Key, "k_")
        If Val(arrNodeKey(4)) <> intTipoBackOfficeUsuario Then
            frmMural.Display = "Usuário não autorizado a visualizar detalhe deste Veículo Legal."
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
        Else
            Call flCarregarTotPorVeiculoLegal(Node)
        End If
    End If
    
    fgCursor

    Exit Sub
    
ErrorHandler:
   
   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - trvGeral_NodeClick"
    
End Sub
