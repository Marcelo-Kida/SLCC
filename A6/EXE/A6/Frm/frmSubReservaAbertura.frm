VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmSubReservaAbertura 
   Caption         =   "Sub-reserva - Abertura"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9615
   Tag             =   "Veículos Legais"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   4020
      Top             =   0
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
            Picture         =   "frmSubReservaAbertura.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":0224
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":0F6C
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":13BE
            Key             =   "posterior"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaAbertura.frx":1810
            Key             =   "aplicar"
         EndProperty
      EndProperty
   End
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   661
   End
   Begin A6.ctlMenu ctlMenu1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   1508
      _ExtentY        =   661
   End
   Begin FPSpread.vaSpread vasAbertura 
      Height          =   4755
      Left            =   60
      TabIndex        =   1
      Tag             =   "Sub Reserva Abertura"
      Top             =   480
      Visible         =   0   'False
      Width           =   9555
      _Version        =   196608
      _ExtentX        =   16854
      _ExtentY        =   8387
      _StockProps     =   64
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
      GrayAreaBackColor=   16777215
      MaxCols         =   0
      MaxRows         =   0
      SpreadDesigner  =   "frmSubReservaAbertura.frx":1C62
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5340
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      ButtonWidth     =   2328
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
            Caption         =   "Abrir Caixa"
            Key             =   "abrircaixa"
            Object.ToolTipText     =   "Abrir Caixa"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
            Key             =   "showfiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
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
   Begin VB.Label lblGrupoVeiculo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5130
   End
End
Attribute VB_Name = "frmSubReservaAbertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a abertura do caixa sub-reserva.

Option Explicit

Private strOperacao                         As String
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlDocFiltros                       As String
Private lngMaxLinha                         As Long
Private strSepMilhar                        As String
Private strSepDecimal                       As String

Private Const strTableComboInicial          As String = "Grupos de Veículos Legais"

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

    On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, _
             enumTipoSelecao.DesmarcarTodas
            Call flMarcarDesmarcarTodas(Retorno)
        Case enumTipoAbertura.SaldoConta, _
             enumTipoAbertura.ValorD1, _
             enumTipoAbertura.ValorInformado
            Call flDefinirTipoAbertura(Retorno)
    End Select
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - ctlMenu1_ClickMenu "

End Sub

'Define o tipo de Abertura Selecionado no Menu Popup

Private Sub flDefinirTipoAbertura(ByVal plngTipoAbertura As enumTipoAbertura)

On Error GoTo ErrorHandler

    With Me.vasAbertura
        .Col = 4
        If .Text <> "Aberto" And .Text <> "" Then
            .Col = 8
            .Row = .ActiveRow
            Select Case plngTipoAbertura
                Case enumTipoAbertura.SaldoConta
                    .SetText 8, .ActiveRow, "Valor Recebido Legado"
                    .ForeColor = vbBlack
                    .SetFloat 7, .ActiveRow, 0
                Case enumTipoAbertura.ValorD1
                    .SetText 8, .ActiveRow, "Valor Calculado A6"
                    .ForeColor = vbBlack
                    .SetFloat 7, .ActiveRow, 0
                Case enumTipoAbertura.ValorInformado
                    With frmMural
                        .Caption = Me.Caption
                        .Display = "Para definir este tipo de abertura, preencha o campo Valor Informado Usuário."
                        .Show vbModal
                    End With
                    .Col = 7
                    .Action = ActionActiveCell
            End Select
        End If
    End With
    
Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaAbertura - flDefinirTipoAbertura ", 0
End Sub

' Marca ou desmarca todos os itens do grid.

Private Sub flMarcarDesmarcarTodas(ByVal plngTipoSelecao As enumTipoSelecao)

Dim lngLinha                         As Long

On Error GoTo ErrorHandler

    With Me.vasAbertura
    For lngLinha = 1 To .MaxRows
        .Col = 4
        .Row = lngLinha
        If .Text <> "Aberto" And .Text <> "" Then
            If plngTipoSelecao = enumTipoSelecao.MarcarTodas Then
                .SetText 1, lngLinha, 1
            End If
        End If
        If plngTipoSelecao = enumTipoSelecao.DesmarcarTodas Then
            .SetText 1, lngLinha, 0
        End If
    Next
    End With

Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, Me.Name, "frmSubReservaAbertura - flMarcarDesmarcarTodas ", 0
End Sub

' Exibe resultado do processamento em lote.

Private Sub flMostrarResultado(ByVal pstrResultado As String)

    With frmResultOperacaoLote
        .strDescricaoOperacao = " abertos "
        .Resultado = pstrResultado
        .Show vbModal
    End With

End Sub

Private Sub ctlTableCombo_AplicarFiltro(pxmlDocFiltros As String)
    
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCursor True
        
    If Trim(xmlDocFiltros) = vbNullString Then
        Call flCarregarVasAbertura(pxmlDocFiltros)
        xmlDocFiltros = pxmlDocFiltros
    Else
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltro.loadXML(xmlDocFiltros)
        
        If Not xmlDomFiltro.selectSingleNode("//Grupo_GrupoVeiculoLegal") Is Nothing Then
            Call fgRemoveNode(xmlDomFiltro, "Grupo_GrupoVeiculoLegal")
        End If
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(pxmlDocFiltros)
        
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_GrupoVeiculoLegal", "GrupoVeiculoLegal", _
                                        xmlDomLeitura.selectSingleNode("//GrupoVeiculoLegal").Text)
                                        
        Call flCarregarVasAbertura(xmlDomFiltro.xml)
        xmlDocFiltros = xmlDomFiltro.xml
    End If
    
    fgCursor
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - ctlTableCombo_AplicarFiltro"
    
End Sub

Private Sub ctlTableCombo_DropDown()

    On Error GoTo ErrorHandler

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
   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - ctlTableCombo_MouseMove"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        fgCursor True
        flCarregarVasAbertura xmlDocFiltros
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - Form_KeyDown"

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Me.Icon = mdiSBR.Icon

    fgCursor True
    fgCenterMe Me
    
    strSepDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    strSepMilhar = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SThousand")
    
    vasAbertura.MaxRows = 1
    Call flFormatarvasAbertura
    Call flCarregarFiltro
    
    Me.Show
    DoEvents
    
    fgCursor
    vasAbertura.Visible = True

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - Form_Load"

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next

    With Me
        .lblGrupoVeiculo.Left = 0
        .lblGrupoVeiculo.Top = 0
        .lblGrupoVeiculo.Width = .ScaleWidth
        .lblGrupoVeiculo.Height = .ctlTableCombo.Height
        
        .vasAbertura.Left = 0
        .vasAbertura.Top = .lblGrupoVeiculo.Height
        .vasAbertura.Width = .ScaleWidth
        .vasAbertura.Height = .ScaleHeight - vasAbertura.Top - .tlbButtons.Height
    End With

End Sub

' Aciona a abertura do caixa para os veículos legais selecionados.

Private Function flAbrirCaixa() As String

#If EnableSoap = 1 Then
    Dim objAbrirCaixa   As MSSOAPLib30.SoapClient30
#Else
    Dim objAbrirCaixa   As A6MIU.clsCaixaSubReserva
#End If

Dim xmlAbrirCaixa       As MSXML2.DOMDocument40
Dim lvwItemAbertura     As MSComctlLib.ListItem
Dim blnTemChecked       As Boolean
Dim lngLinha            As Long

Dim intValorCelula      As Integer
Dim vntConteudoCelula   As Variant

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    'Verificar se existe algum item selecionado para a abertura
    blnTemChecked = False
    With Me.vasAbertura
        .BlockMode = False
    
        For lngLinha = 1 To .MaxRows
            .Col = 1
            .Row = lngLinha
            intValorCelula = .Value
            .GetText 2, lngLinha, vntConteudoCelula
            
            If intValorCelula = vbChecked And vntConteudoCelula <> vbNullString Then
                blnTemChecked = True
                Exit For
            End If
        Next
    End With
    
    If Not blnTemChecked Then
        frmMural.Display = "Selecionar Veículo(s) Legal(ais) para Abertura."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Function
    End If

    If MsgBox("Confirma a Abertura dos Veículos Legais", vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbNo Then Exit Function

    Set xmlAbrirCaixa = CreateObject("MSXML2.DOMDocument.4.0")

    fgCursor True
    fgAppendNode xmlAbrirCaixa, "", "CaixaSubReservaAbertura", ""

    blnTemChecked = False
    'Monta XML com os dados necessários para Abertura dos Veículos Legais
    With Me.vasAbertura
        For lngLinha = 1 To .MaxRows
            .Col = 1
            .Row = lngLinha
            If .Value = 1 Then
                blnTemChecked = True
                fgAppendNode xmlAbrirCaixa, "CaixaSubReservaAbertura", "GrupoVeiculoLegal", ""
                .Col = 2
                fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "NO_VEIC_LEGA", .Text, "CaixaSubReservaAbertura"
                .Col = 12
                fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "CO_VEIC_LEGA", .Text, "CaixaSubReservaAbertura"
                .Col = 13
                fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "SG_SIST", .Text, "CaixaSubReservaAbertura"
                .Col = 5
                fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "VA_ABER_RECE_CAIX_SUB_RESE", fgVlr_To_Xml(.Text), "CaixaSubReservaAbertura"
                .Col = 6
                fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "VA_MOVI_CAIX_SUB_RESE", fgVlr_To_Xml(.Text), "CaixaSubReservaAbertura"
                .Col = 7
                fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "VA_ABER_INFO_CAIX_SUB_RESE", fgVlr_To_Xml(.Text), "CaixaSubReservaAbertura"
                
                .Col = 8
                Select Case .Text
                    Case "Valor Calculado A6"
                        fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "TipoAbertura", enumTipoAbertura.ValorD1, "CaixaSubReservaAbertura"
                    Case "Valor Informado Usuário"
                        fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "TipoAbertura", enumTipoAbertura.ValorInformado, "CaixaSubReservaAbertura"
                    Case "Valor Recebido Legado"
                        fgAppendNode xmlAbrirCaixa, "GrupoVeiculoLegal", "TipoAbertura", enumTipoAbertura.SaldoConta, "CaixaSubReservaAbertura"
                End Select
    
            End If
        Next
    End With

    Set objAbrirCaixa = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")
    flAbrirCaixa = objAbrirCaixa.AbrirCaixa(xmlAbrirCaixa.xml, _
                                            vntCodErro, _
                                            vntMensagemErro)
    Set objAbrirCaixa = Nothing

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    fgCursor

    Set xmlAbrirCaixa = Nothing

    Exit Function

ErrorHandler:
    Set xmlAbrirCaixa = Nothing
    Set objAbrirCaixa = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaAbertura - flAbrirCaixa ", 0

End Function

' Exibe tela de filtro.

Private Sub flCarregarFiltro()

On Error GoTo ErrorHandler

    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaAbertura
    Load objFiltro

    Call objFiltro.fgCarregarPesquisaAnterior

    Exit Sub
ErrorHandler:

     fgRaiseError App.EXEName, Me.Name, "frmSubReservaAbertura - flCarregarFiltro ", 0

End Sub

Private Sub lblGrupoVeiculo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub

ErrorHandler:
   
   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - lblGrupoVeiculo_MouseMove"

End Sub

' Carrega formatação inicial do grid de veículos legais.

Private Sub flFormatarvasAbertura()
Dim intCol                                  As Integer
Dim lngTamanhoAtributo                      As Long
Dim lngCasasDecimais                        As Long
Dim strMascara                              As String

On Error GoTo ErrorHandler

    lngTamanhoAtributo = 15
    lngCasasDecimais = 2
    
    With Me.vasAbertura
    
        .Redraw = False
        .DisplayRowHeaders = False
        .CursorStyle = CursorStyleArrow
        
        .MaxCols = 13
        
        .Col = 1
        .CellType = CellTypeCheckBox
        .ColWidth(1) = 2
        
        .SetText 1, 0, " "
        .SetText 2, 0, "Veículo Legal"
        .SetText 3, 0, "Data para Abertura"
        .SetText 4, 0, "Status Caixa"
        .SetText 5, 0, "Valor Calculado A6"
        .SetText 6, 0, "Valor Recebido Legado"
        .SetText 7, 0, "Valor Informado Usuário"
        .SetText 8, 0, "Valor Utilizado"
        .SetText 9, 0, "Informada"
        .SetText 10, 0, "Aceita"
        .SetText 11, 0, "Rejeitada"
        
        .ColWidth(2) = 32
        .ColWidth(3) = 13
        .ColWidth(4) = 8
        .ColWidth(5) = 17
        .ColWidth(6) = 17
        .ColWidth(7) = 17
        .ColWidth(8) = 17
        
        For intCol = 9 To .MaxCols
            .ColWidth(intCol) = 0
        Next
        
        .Row = 1
        .Row2 = .MaxRows
        .Col = 2
        .Col2 = 11
        .BlockMode = True
        .CellType = CellTypeStaticText
        .BlockMode = False
        
        .Row = 1
        .Row2 = .MaxRows
        .Col = 5
        .Col2 = 11
        .BlockMode = True
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
        
        .Row = 1
        .Row2 = .MaxRows
        .Col = 7
        .Col2 = 7
        .BlockMode = True
        .CellType = CellTypeFloat
        strMascara = String(lngTamanhoAtributo, "9") & strSepDecimal & String(lngCasasDecimais, "9")
        .TypeFloatMax = strMascara
        .TypeFloatMin = "-" & strMascara
        .TypeFloatDecimalPlaces = lngCasasDecimais
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignTop
        .TypeFloatMoney = False
        .TypeFloatSeparator = True
        .TypeFloatDecimalChar = Asc(strSepDecimal)
        .TypeFloatSepChar = Asc(strSepMilhar)
        .BlockMode = False
        
        .Redraw = True
    
    End With

    Exit Sub
ErrorHandler:
     
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaAbertura - flFormatarlvwAbertura ", 0
    
End Sub

' Aciona a leitura de veículos legais em condições de serem abertos.

Private Sub flLerTodos()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim strRetorno          As String
Dim strPropriedades     As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    strOperacao = "LerTodos"

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")

    xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_MovimentoSubReserva").attributes.getNamedItem("Operacao").Text = strOperacao
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_MovimentoSubReserva").xml

    Call objMIU.Executar(strPropriedades, _
                         vntCodErro, _
                         vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaAbertura - flLerTodos ", 0

End Sub

' Carrega lista de veículos legais após a leitura.

Private Sub flCarregarVasAbertura(ByRef pstrDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objAberturaCaixa    As MSSOAPLib30.SoapClient30
#Else
    Dim objAberturaCaixa    As A6MIU.clsCaixaSubReserva
#End If

Dim xmlAbertura             As MSXML2.DOMDocument40
Dim xmlDomNode              As MSXML2.IXMLDOMNode
Dim lngLinha                As Long
Dim strDataServidor         As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If pstrDocFiltros = vbNullString Then
        frmMural.Display = "A seleção do Grupo de Veículo Legal é obrigatória. Por favor, clique em Definir Filtro, selecione um Grupo de Veículo Legal, e tente novamente."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If

    Set objAberturaCaixa = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")

    Set xmlAbertura = CreateObject("MSXML2.DOMDocument.4.0")
    xmlAbertura.loadXML (objAberturaCaixa.ObterValoresAbertura(pstrDocFiltros, _
                                                               vntCodErro, _
                                                               vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    strDataServidor = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    
    With Me.vasAbertura
        
        .Redraw = False
        
        .MaxRows = 0
        .MaxRows = 1
        
        lngLinha = 1

        For Each xmlDomNode In xmlAbertura.selectNodes("/Repeat_CaixaSubReserva/*")
           
           .SetText 2, lngLinha, xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text
           .SetText 3, lngLinha, fgDtXML_To_Interface(xmlDomNode.selectSingleNode("DT_ABER_CAIX").Text)
           
           If Val(xmlDomNode.selectSingleNode("CO_SITU_CAIX_SUB_RESE").Text) = enumEstadoCaixa.Fechado And _
              strDataServidor = xmlDomNode.selectSingleNode("DT_CAIX_SUB_RESE").Text Then
               .SetText 4, lngLinha, fgDescricaoEstadoCaixa(Fechado)
           Else
               .SetText 4, lngLinha, fgDescricaoEstadoCaixa(Disponivel)
           End If
           
           .SetText 5, lngLinha, fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_FECH_CAIX_SUB_RESE").Text)
           .SetText 6, lngLinha, fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_SALD_FECH").Text)
           .SetText 7, lngLinha, fgVlrXml_To_Interface("0")
           .SetText 8, lngLinha, "Valor Recebido Legado"
           .SetText 12, lngLinha, xmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text
           .SetText 13, lngLinha, xmlDomNode.selectSingleNode("SG_SIST").Text
           
           lngLinha = lngLinha + 1
    
           If .MaxRows < lngLinha Then
              .MaxRows = .MaxRows + 1
           End If
       
       Next
       
       If .MaxRows > 1 Then
          .MaxRows = .MaxRows - 1
       End If
       
       .Refresh
       .Redraw = True
       
    End With
    
    Call flFormatarvasAbertura

RemoverInstancias:
    Set objAberturaCaixa = Nothing
    Set xmlAbertura = Nothing
    Set xmlDomNode = Nothing
    
    Exit Sub

ErrorHandler:
    Set objAberturaCaixa = Nothing
    Set xmlAbertura = Nothing
    Set xmlDomNode = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarVasAbertura", lngCodigoErroNegocio

End Sub

Private Sub objFiltro_AplicarFiltro(pxmlDocFiltros As String, plsTituloTableCombo As String)

On Error GoTo ErrorHandler

    DoEvents
    Call flCarregarVasAbertura(pxmlDocFiltros)
    xmlDocFiltros = pxmlDocFiltros

    ctlTableCombo.TituloCombo = IIf(plsTituloTableCombo = vbNullString, strTableComboInicial, plsTituloTableCombo)

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - objFiltro_AplicarFiltro"
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultado                            As String

On Error GoTo ErrorHandler

    Select Case Button.Key
    Case "abrircaixa"
        strResultado = flAbrirCaixa
        If strResultado <> vbNullString Then
            fgCursor True
            
            Call flMostrarResultado(strResultado)
            Call flCarregarVasAbertura(xmlDocFiltros)
            Call flMarcarDesmarcarTodas(DesmarcarTodas)
            
        End If
    
    Case "showfiltro"
        Set objFiltro = New frmFiltro
        Set objFiltro.FormOwner = Me
        objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaAbertura
        objFiltro.Show vbModal
        
    Case "refresh"
        fgCursor True
        flCarregarVasAbertura xmlDocFiltros
        fgCursor
    
    End Select
    
    fgCursor
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura  - tlbButtons_ButtonClick"

End Sub

Private Sub vasAbertura_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo ErrorHandler

   If Col = 7 Then
        With vasAbertura
            .Row = Row
            .Col = Col
            If .Value <> 0 Then
                .SetText 8, .Row, "Valor Informado Usuário"
                .Col = 8
                .ForeColor = vbRed
            Else
                .SetText 8, .Row, "Valor Recebido Legado"
                .Col = 8
                .ForeColor = vbBlack
            End If
        End With
    End If

Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - vasAbertura_Change"
    
End Sub

Private Sub vasAbertura_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo ErrorHandler
    
    If Col = 1 And Row > 0 Then
       vasAbertura.Col = 2
       vasAbertura.Row = Row
       If vasAbertura.Text = "" Then
          vasAbertura.SetText Col, Row, 0
       Else
          vasAbertura.Col = 4
          vasAbertura.Row = Row
          If vasAbertura.Text = "Aberto" Then
             vasAbertura.Text = 0
           End If
       End If
    End If

    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura  - vasAbertura_Click"

End Sub

Private Sub vasAbertura_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo ErrorHandler

   If Col = 7 Then
        With vasAbertura
            .Row = Row
            .Col = Col
            If .Value <> 0 Then
                .SetText 8, .Row, "Valor Informado Usuário"
                .Col = 8
                .ForeColor = vbRed
            Else
                .SetText 8, .Row, "Valor Recebido Legado"
                .Col = 8
                .ForeColor = vbBlack
            End If
        End With
    End If

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - vasAbertura_EditChange"
   
End Sub

Private Sub vasAbertura_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim Col                                     As Long
Dim Row                                     As Long

On Error GoTo ErrorHandler

    With vasAbertura
        Call .GetCellFromScreenCoord(Col, Row, x, y)
        If Row > 0 Then
            .Col = Col
            .Row = Row
            .Action = ActionActiveCell
        End If

    End With

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuCaixaSubReservaAbertura
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura  - vasAbertura_MouseDown"

End Sub

Private Sub vasAbertura_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub

ErrorHandler:
   
   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - vasAbertura_MouseMove"

End Sub

