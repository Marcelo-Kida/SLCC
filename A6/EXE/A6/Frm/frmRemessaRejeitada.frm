VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRemessaRejeitada 
   Caption         =   "Sub-reserva - Remessa Rejeitada"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleMode       =   0  'User
   ScaleWidth      =   10530
   WindowState     =   2  'Maximized
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   661
   End
   Begin MSFlexGridLib.MSFlexGrid flxMonitoracao 
      Height          =   6225
      Left            =   0
      TabIndex        =   0
      Tag             =   "Remessas Rejeitadas"
      Top             =   495
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   10980
      _Version        =   393216
      Rows            =   10
      Cols            =   8
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   -2147483638
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6750
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      ButtonWidth     =   2434
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir &Filtro"
            Key             =   "showfiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgOrder 
      Left            =   4860
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0000
            Key             =   "Cima"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0192
            Key             =   "Baixo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   4290
      Top             =   0
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
            Picture         =   "frmRemessaRejeitada.frx":0324
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0436
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0548
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":089A
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0BEC
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0F3E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":1290
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":16E2
            Key             =   "posterior"
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
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10650
   End
End
Attribute VB_Name = "frmRemessaRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a consulta às remessas rejeitadas pelo A6.

Option Explicit

Private strOperacao                         As String

'Variaveis para a utilização do Filtro
Private strFiltroXML                        As String
Private blnPrimeiraConsulta                 As Boolean

Private Const strFuncionalidade             As String = "REMESSAREJEITADA"
Private Const strTableComboInicial          As String = "Empresa"

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private intCompareCol                       As Integer
Private intCompareOrder                     As Integer

Private Const COMPARE_ASC                   As Integer = 1
Private Const COMPARE_DESC                  As Integer = -1

'Definição das Colunas do Grid
Private Const COL_DH_REME_REJE              As Integer = 0
Private Const COL_EMPRESA                   As Integer = 1
Private Const COL_SISTEMA                   As Integer = 2
Private Const COL_VEICULO_LEGAL             As Integer = 3
Private Const COL_TIPO_MENSAGEM             As Integer = 4
Private Const COL_KEY                       As Integer = 5

Private Const MAX_FIXED_ROWS                As Integer = 50

' Carrega grid com o resultado da pesquisa à tabela de remessas rejeitadas pelo A6.

Private Sub flCarregarFlexGrid(ByRef pxmlDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objRemessaRejeitada As MSSOAPLib30.SoapClient30
#Else
    Dim objRemessaRejeitada As A6MIU.clsConsultaRemessaRejeitada
#End If

Dim xmlRejeitada            As MSXML2.DOMDocument40
Dim xmlDomRejeitada         As MSXML2.IXMLDOMNode
Dim lngLinhaGrid            As Long
Dim strXMLRetorno           As String
Dim intCount                As Integer
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    gintRowPositionAnt = 0

    Set objRemessaRejeitada = fgCriarObjetoMIU("A6MIU.clsConsultaRemessaRejeitada")

    Set xmlRejeitada = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objRemessaRejeitada.LerTodos(pxmlDocFiltros, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flFormatarFlxMonitoracao
    
    'caso a tabela esteja sem registros não tem como carregar um XML,
    'sendo assim vai para o fim da rotina.
    If Trim(strXMLRetorno) <> "" Then
       If Not xmlRejeitada.loadXML(strXMLRetorno) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Exit Sub
    End If

    lngLinhaGrid = 1

    With Me.flxMonitoracao
    
        .Redraw = False

        For Each xmlDomRejeitada In xmlRejeitada.documentElement.selectNodes("//Repeat_Erro/*")
        
            .TextMatrix(lngLinhaGrid, COL_SISTEMA) = xmlDomRejeitada.selectSingleNode("SG_SIST_ORIG_INFO").Text
            
            If xmlDomRejeitada.selectSingleNode("NO_SIST").Text <> vbNullString Then
                .TextMatrix(lngLinhaGrid, COL_SISTEMA) = .TextMatrix(lngLinhaGrid, COL_SISTEMA) & " - " & _
                                                         xmlDomRejeitada.selectSingleNode("NO_SIST").Text
            End If
            
            .TextMatrix(lngLinhaGrid, COL_EMPRESA) = IIf(Val(xmlDomRejeitada.selectSingleNode("CO_EMPR").Text) = 0, vbNullString, xmlDomRejeitada.selectSingleNode("CO_EMPR").Text)
            
            If xmlDomRejeitada.selectSingleNode("NO_REDU_EMPR").Text <> vbNullString Then
                .TextMatrix(lngLinhaGrid, COL_EMPRESA) = .TextMatrix(lngLinhaGrid, COL_EMPRESA) & " - " & _
                                                         xmlDomRejeitada.selectSingleNode("NO_REDU_EMPR").Text
            End If
            
            .TextMatrix(lngLinhaGrid, COL_VEICULO_LEGAL) = xmlDomRejeitada.selectSingleNode("CO_VEIC_LEGA").Text
            
            If xmlDomRejeitada.selectSingleNode("NO_VEIC_LEGA").Text <> vbNullString Then
                .TextMatrix(lngLinhaGrid, COL_VEICULO_LEGAL) = .TextMatrix(lngLinhaGrid, COL_VEICULO_LEGAL) & " - " & _
                                                               xmlDomRejeitada.selectSingleNode("NO_VEIC_LEGA").Text
            End If
            
            .TextMatrix(lngLinhaGrid, COL_TIPO_MENSAGEM) = xmlDomRejeitada.selectSingleNode("NO_TIPO_MESG").Text
            
            .TextMatrix(lngLinhaGrid, COL_DH_REME_REJE) = fgDtHrStr_To_DateTime(xmlDomRejeitada.selectSingleNode("DH_REME_REJE").Text)
            .TextMatrix(lngLinhaGrid, COL_KEY) = ";" & xmlDomRejeitada.selectSingleNode("SG_SIST_ORIG_INFO").Text & _
                                                 ";" & xmlDomRejeitada.selectSingleNode("TP_MESG_INTE").Text & _
                                                 ";" & xmlDomRejeitada.selectSingleNode("CO_EMPR").Text & _
                                                 ";" & xmlDomRejeitada.selectSingleNode("CO_TEXT_XML_REJE").Text & _
                                                 ";" & xmlDomRejeitada.selectSingleNode("DH_REME_REJE").Text & _
                                                 ";" & xmlDomRejeitada.selectSingleNode("OWNER").Text
                
            If (.Rows - 1) > lngLinhaGrid Then
               lngLinhaGrid = lngLinhaGrid + 1
            Else
               lngLinhaGrid = lngLinhaGrid + 1
               .Rows = .Rows + 1
            End If
            
        Next xmlDomRejeitada
        
        .Rows = .Rows - 1
        
        .Col = COL_DH_REME_REJE
        .Sort = 9

        .Redraw = True
        
    End With
    
    flxMonitoracao.Col = COL_EMPRESA
    flxMonitoracao.Row = 1

    Set xmlRejeitada = Nothing
    Set objRemessaRejeitada = Nothing

    Exit Sub

ErrorHandler:

    Set xmlRejeitada = Nothing
    Set objRemessaRejeitada = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmRemessaRejeitada - flCarregaFlexGrid", 0

End Sub

Private Sub flxMonitoracao_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

Dim vntRow1Value                            As Variant
Dim vntRow2Value                            As Variant

On Error GoTo ErrorHandler

    With flxMonitoracao
        
        .Col = intCompareCol
        
        If .TextMatrix(Row1, .Col) = vbNullString Then
            If .TextMatrix(Row2, .Col) = vbNullString Then
                Cmp = 0
            Else
                Cmp = 1
            End If
            Exit Sub
        End If
        
        If .TextMatrix(Row2, .Col) = vbNullString Then
            Cmp = -1
            Exit Sub
        End If
        
        Select Case intCompareCol
            Case COL_DH_REME_REJE
                vntRow1Value = CDate(.TextMatrix(Row1, .Col))
                vntRow2Value = CDate(.TextMatrix(Row2, .Col))
                
            Case COL_EMPRESA
                vntRow1Value = CLng(fgObterCodigoCombo(.TextMatrix(Row1, .Col)))
                vntRow2Value = CLng(fgObterCodigoCombo(.TextMatrix(Row2, .Col)))
                
            Case COL_SISTEMA
                vntRow1Value = fgObterCodigoCombo(.TextMatrix(Row1, .Col))
                vntRow2Value = fgObterCodigoCombo(.TextMatrix(Row2, .Col))
                
            Case COL_VEICULO_LEGAL
                vntRow1Value = fgObterCodigoCombo(.TextMatrix(Row1, .Col))
                vntRow2Value = fgObterCodigoCombo(.TextMatrix(Row2, .Col))
                
            Case COL_TIPO_MENSAGEM
                vntRow1Value = .TextMatrix(Row1, .Col)
                vntRow2Value = .TextMatrix(Row2, .Col)
                
        End Select
        
        If vntRow1Value > vntRow2Value Then
            Cmp = -1 * intCompareOrder
        ElseIf vntRow1Value = vntRow2Value Then
            Cmp = 0
        Else
            Cmp = 1 * intCompareOrder
        End If
        
    End With

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - flxMonitoracao_Compare"

End Sub

Private Sub flxMonitoracao_DblClick()

#If EnableSoap = 1 Then
    Dim objRemessa          As MSSOAPLib30.SoapClient30
#Else
    Dim objRemessa          As A6MIU.clsConsultaRemessaRejeitada
#End If

Dim arrChave()              As String
Dim strXML                  As String
Dim xmlLer                  As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    With flxMonitoracao
        
        If Len(.TextMatrix(.Row, COL_KEY)) = 0 Then Exit Sub
        
        fgCursor True
        arrChave = Split(.TextMatrix(.Row, COL_KEY), ";")
    
        Set objRemessa = fgCriarObjetoMIU("A6MIU.clsConsultaRemessaRejeitada")
    
        strXML = objRemessa.Ler(arrChave(1), _
                                CInt(arrChave(2)), _
                                CLng(arrChave(3)), _
                                CLng(arrChave(4)), _
                                arrChave(5), _
                                arrChave(6), _
                                vntCodErro, _
                                vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If Len(strXML) = 0 Then Exit Sub
        
        Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
        If Not xmlLer.loadXML(strXML) Then
            fgErroLoadXML xmlLer, "flxMonitoracao_DblClick", "", ""
        End If
        
        strXML = xmlLer.documentElement.selectSingleNode("TX_XML_ERRO").Text
        xmlLer.loadXML strXML
        
        If arrChave(6) = "A6HIST" Then
            frmDetalheRemessa.lngCO_TEXT_XML_REJE = CLng(arrChave(4)) * -1
        Else
            frmDetalheRemessa.lngCO_TEXT_XML_REJE = CLng(arrChave(4))
        End If
        frmDetalheRemessa.strXMLErro = strXML
        frmDetalheRemessa.Show vbModal
        fgCursor
    
    End With

Exit Sub

ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmRemessaRejeitada - flxMonitoracao_DblClick", Me.Caption

End Sub

Private Sub flxMonitoracao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim lngColuna                               As Long
Dim lngCont                                 As Long

On Error GoTo ErrorHandler

    With flxMonitoracao

        If .RowPos(0) > y Or .RowPos(0) + .RowHeight(0) < y Then Exit Sub
        
        lngColuna = -1
        
        For lngCont = 0 To .Cols - 1
            If .ColPos(lngCont) < x And .ColPos(lngCont) + .ColWidth(lngCont) > x Then
                lngColuna = lngCont
                Exit For
            End If
        Next lngCont
        
        If lngColuna = -1 Then Exit Sub
    
        fgLockWindow Me.hwnd
        fgCursor True
        .Redraw = False
        
        If lngColuna <> intCompareCol Then
            intCompareCol = lngColuna
            intCompareOrder = COMPARE_ASC
        Else
            'Inverte a ordem de comparação
            intCompareOrder = intCompareOrder * -1
        End If
        
        .Sort = 9
        
        .Row = 0
        For lngCont = 0 To .Cols - 1
            .Col = lngCont
            .CellPictureAlignment = flexAlignRightCenter
            If lngColuna = lngCont Then
                If intCompareOrder = COMPARE_ASC Then
                    Set .CellPicture = imgOrder.ListImages("Cima").Picture
                Else
                    Set .CellPicture = imgOrder.ListImages("Baixo").Picture
                End If
            Else
                Set .CellPicture = Nothing
            End If
        Next lngCont
    
        .Redraw = True
    
    End With

    fgLockWindow 0
    fgCursor
    
    Exit Sub
ErrorHandler:

    fgLockWindow 0
    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - flxMonitoracao_MouseDown"
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        If Trim(strFiltroXML) = vbNullString Then Exit Sub
        
        fgCursor True
        Call flCarregarFlexGrid(strFiltroXML)
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmRemessaRejeitada - Form_KeyDown"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    Me.Icon = mdiSBR.Icon

    intCompareCol = COL_DH_REME_REJE
    intCompareOrder = 1
    
    fgCenterMe Me
    
    fgCursor True
    
    blnPrimeiraConsulta = True
    ctlTableCombo.TituloCombo = strTableComboInicial
    
    flFormatarFlxMonitoracao
    
    Me.Show
    DoEvents
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmRemessaRejeitada
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    objFiltro.Show vbModal
    
    blnPrimeiraConsulta = False

    DoEvents
    Me.Refresh
    
    fgCursor

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "frmRemessaRejeitada - Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .ctlTableCombo.Left = 0
        .ctlTableCombo.Top = 0
        
        .lblGrupoVeiculo.Left = 0
        .lblGrupoVeiculo.Top = 0
        .lblGrupoVeiculo.Width = .ScaleWidth
        .lblGrupoVeiculo.Height = .ctlTableCombo.Height
        
        .flxMonitoracao.Left = 0
        .flxMonitoracao.Top = .lblGrupoVeiculo.Height
        .flxMonitoracao.Width = .ScaleWidth
        .flxMonitoracao.Height = .ScaleHeight - .flxMonitoracao.Top - .tlbButtons.Height
    End With

End Sub

' Inicialização do grid de consulta às remessas rejeitadas.

Private Sub flFormatarFlxMonitoracao()

Dim intCount                                As Integer
Dim intLinhaGridFix                         As Integer

On Error GoTo ErrorHandler

    intLinhaGridFix = 0
    
    With Me.flxMonitoracao
    
        .Clear
        .Rows = MAX_FIXED_ROWS
        .Cols = 6
        .FixedRows = 1
        
        For intCount = 0 To .Cols - 1
            .ColAlignment(intCount) = MSFlexGridLib.flexAlignLeftCenter
        Next
        
        .GridColorFixed = &HE6E6E6

        .TextMatrix(intLinhaGridFix, COL_SISTEMA) = "Sistema"
        .ColWidth(COL_SISTEMA) = 2500
        
        .TextMatrix(intLinhaGridFix, COL_EMPRESA) = "Empresa"
        .ColWidth(COL_EMPRESA) = 3500
        
        .TextMatrix(intLinhaGridFix, COL_TIPO_MENSAGEM) = "Tipo"
        .ColWidth(COL_TIPO_MENSAGEM) = 3500
        
        .TextMatrix(intLinhaGridFix, COL_DH_REME_REJE) = "Data/Hora Rejeição"
        .ColWidth(COL_DH_REME_REJE) = 2000
        
        .TextMatrix(intLinhaGridFix, COL_VEICULO_LEGAL) = "Veículo Legal"
        .ColWidth(COL_VEICULO_LEGAL) = 3300
        
        .ColWidth(COL_KEY) = 0

        .GridLinesFixed = flexGridInset
        .SelectionMode = flexSelectionByRow
        
    End With

    Exit Sub
ErrorHandler:
   
   fgRaiseError App.EXEName, Me.Name, "frmRemessaRejeitada - flFormataflxMonitoracao", 0
   
End Sub

Private Sub ctlTableCombo_AplicarFiltro(xmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCursor True
    
    If Trim(strFiltroXML) = vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltro, "", "Repeat_Filtros", "")
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", _
                                        xmlDomLeitura.selectSingleNode("//BancoLiquidante").Text)
                                        
        If xmlDomFiltro.selectSingleNode("//Grupo_Data") Is Nothing Then
            fgAppendNode xmlDomFiltro, "Repeat_Filtros", "Grupo_Data", ""
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)))
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle("99991231")
        End If
    
        Call flCarregarFlexGrid(xmlDomFiltro.xml)
        strFiltroXML = xmlDomFiltro.xml
    
    Else
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltro.loadXML(strFiltroXML)
        
        If Not xmlDomFiltro.selectSingleNode("//Grupo_BancoLiquidante") Is Nothing Then
            Call fgRemoveNode(xmlDomFiltro, "Grupo_BancoLiquidante")
        End If
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", _
                                        xmlDomLeitura.selectSingleNode("//BancoLiquidante").Text)
                                        
        If xmlDomFiltro.selectSingleNode("//Grupo_Data") Is Nothing Then
            fgAppendNode xmlDomFiltro, "Repeat_Filtros", "Grupo_Data", ""
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)))
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle("99991231")
        End If
    
        Call flCarregarFlexGrid(xmlDomFiltro.xml)
        strFiltroXML = xmlDomFiltro.xml
    End If
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    fgCursor

    Exit Sub

ErrorHandler:
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing

    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - ctlTableCombo_AplicarFiltro"

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

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - ctlTableCombo_MouseMove"
   
End Sub

Private Sub flxMonitoracao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - flxMonitoracao_MouseMove"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmRemessaRejeitada = Nothing
    gintRowPositionAnt = 0

End Sub

Private Sub lblGrupoVeiculo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - lblGrupoVeiculo_MouseMove"
   
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, lsTituloTableCombo As String)
    
On Error GoTo ErrorHandler

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            If blnPrimeiraConsulta Then
                blnPrimeiraConsulta = False
                
                If InStr(1, strFiltroXML, "DataIni") = 0 Then
                    frmMural.Caption = Me.Caption
                    frmMural.Display = "Obrigatória a seleção do filtro DATA."
                    frmMural.Show vbModal
                    Exit Sub
                End If
                
'                Call tlbButtons_ButtonClick(tlbButtons.Buttons("showfiltro"))
            Else
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
                
                Exit Sub
            End If
        End If
        
        fgCursor True
        Call flCarregarFlexGrid(strFiltroXML)
        fgCursor
        Me.Refresh
    
        ctlTableCombo.TituloCombo = IIf(lsTituloTableCombo = vbNullString, strTableComboInicial, lsTituloTableCombo)
        
    End If
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmRemessaRejeitada - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "showfiltro"
             Set objFiltro = New frmFiltro
             Set objFiltro.FormOwner = Me
             objFiltro.TipoFiltro = enumTipoFiltroA6.frmRemessaRejeitada
             objFiltro.Show vbModal
    
        Case "refresh"
             If Trim(strFiltroXML) = vbNullString Then Exit Sub
            
             If InStr(1, strFiltroXML, "DataIni") = 0 Then
                 frmMural.Caption = Me.Caption
                 frmMural.Display = "Obrigatória a seleção do filtro DATA."
                 frmMural.Show vbModal
                 Exit Sub
             End If
        
             fgCursor True
             Call flCarregarFlexGrid(strFiltroXML)
             fgCursor
    End Select

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbButtons_ButtonClick"
   
End Sub
