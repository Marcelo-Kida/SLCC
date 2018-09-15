VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSubReservaFechamento 
   Caption         =   "Sub-reserva - Fechamento"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9525
   Tag             =   "Veículos Legais"
   WindowState     =   2  'Maximized
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
      Left            =   5400
      Top             =   120
      _ExtentX        =   1720
      _ExtentY        =   450
   End
   Begin MSComctlLib.ListView lvwFechamento 
      Height          =   4845
      Left            =   0
      TabIndex        =   0
      Tag             =   "Sub Reserva Fechamento"
      Top             =   445
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   8546
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5310
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   635
      ButtonWidth     =   2381
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
            Caption         =   "Fechar Caixa"
            Key             =   "fecharcaixa"
            Object.ToolTipText     =   "Fechar Caixa"
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
            Picture         =   "frmSubReservaFechamento.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":0224
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":0F6C
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":13BE
            Key             =   "posterior"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaFechamento.frx":1810
            Key             =   "aplicar"
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
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4110
   End
End
Attribute VB_Name = "frmSubReservaFechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário o fechamento do caixa sub-reserva.

Option Explicit

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private Const strTableComboInicial          As String = "Grupos de Veículos Legais"

Private strOperacao                         As String
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlDocFiltros                       As String

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, _
             enumTipoSelecao.DesmarcarTodas
            Call flMarcarDesmarcatTodas(Retorno)
    End Select
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - ctlMenu1_ClickMenu "
    
End Sub

' Marca ou desmarca todas as linhas do grid.

Private Sub flMarcarDesmarcatTodas(ByVal plngTipoSelecao As enumTipoSelecao)

Dim lvwItemFechamento                         As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    For Each lvwItemFechamento In lvwFechamento.ListItems
        If lvwItemFechamento.SubItems(lvwFechamento.ColumnHeaders("StatusCaixa").SubItemIndex) = "Aberto" Then
            If plngTipoSelecao = enumTipoSelecao.MarcarTodas Then
                lvwItemFechamento.Checked = True
            End If
        End If
    
        If plngTipoSelecao = enumTipoSelecao.DesmarcarTodas Then
            lvwItemFechamento.Checked = False
        End If
    Next

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - flMarcarDesmarcatTodas "

End Sub

' Aciona a exibição do resultado do processamento em lote.

Private Sub flMostrarResultado(ByVal pstrResultado As String)

    With frmResultOperacaoLote
        .strDescricaoOperacao = " fechados "
        .Resultado = pstrResultado
        .Show vbModal
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        fgCursor True
        flCarregarlvwFechamento xmlDocFiltros
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura - Form_KeyDown"

End Sub

Private Sub lblGrupoVeiculo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura  - lblGrupoVeiculo_MouseMove"

End Sub

Private Sub lvwFechamento_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuCaixaSubReservaFechamento
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaAbertura  - lvwFechamento_MouseDown"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    Me.Icon = mdiSBR.Icon

    fgCenterMe Me

    fgCursor True

    Call flFormatarlvwFechamento
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaResumo
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    
    DoEvents
    Me.Show
    
    fgCursor

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
        
        .lvwFechamento.Left = 0
        .lvwFechamento.Top = .lblGrupoVeiculo.Height
        .lvwFechamento.Width = .ScaleWidth
        .lvwFechamento.Height = .ScaleHeight - lvwFechamento.Top - .tlbButtons.Height
    End With

End Sub

' Aciona o fechamento do caixa dos veículos legais selecionados.

Private Function flFecharCaixa() As String

#If EnableSoap = 1 Then
    Dim objFechamento   As MSSOAPLib30.SoapClient30
#Else
    Dim objFechamento   As A6MIU.clsCaixaSubReserva
#End If

Dim xmlFecharCaixa      As MSXML2.DOMDocument40
Dim xmlAux              As MSXML2.DOMDocument40
Dim lvwItemFechamento   As MSComctlLib.ListItem
Dim blnTemChecked       As Boolean
Dim arrItemKey()        As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlFecharCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    
    If lvwFechamento.ListItems.Count = 0 Then
        frmMural.Display = "Carregar Veículo Legal para Fechamento."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Function
    End If
    
    If MsgBox("Confirma o Fechamento dos Veículos Legais", vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbNo Then Exit Function
    
    fgCursor True
    fgAppendNode xmlFecharCaixa, "", "Repeat_CaixaSubReserva", ""
    
    blnTemChecked = False
    For Each lvwItemFechamento In lvwFechamento.ListItems
        If lvwItemFechamento.Checked Then
            arrItemKey = Split(lvwItemFechamento.Key, "k_")
        
            blnTemChecked = True
            
            Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
            fgAppendNode xmlAux, "", "Grupo_VeiculoLegal", ""
            fgAppendNode xmlAux, "Grupo_VeiculoLegal", "NO_VEIC_LEGA", lvwItemFechamento.Text
            fgAppendNode xmlAux, "Grupo_VeiculoLegal", "CO_VEIC_LEGA", arrItemKey(1)
            fgAppendNode xmlAux, "Grupo_VeiculoLegal", "SG_SIST", arrItemKey(2)
            
            fgAppendXML xmlFecharCaixa, "Repeat_CaixaSubReserva", xmlAux.xml
            Set xmlAux = Nothing
            
        End If
    Next

    If blnTemChecked = False Then
        frmMural.Display = "Selecionar Veículo(s) Legal(ais) para Fechamento."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Function
    End If

    Set objFechamento = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")
    flFecharCaixa = objFechamento.FecharCaixa(xmlFecharCaixa.xml, _
                                              vntCodErro, _
                                              vntMensagemErro)
    Set objFechamento = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set xmlFecharCaixa = Nothing
    Set xmlAux = Nothing
    Set objFechamento = Nothing

    Exit Function
    
ErrorHandler:
    Set xmlFecharCaixa = Nothing
    Set objFechamento = Nothing
    Set xmlAux = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - tbrFecharCaixa_ButtonClick "

End Function

'Define as colunas do ListView de fechamento do Caixa

Private Sub flFormatarlvwFechamento()

On Error GoTo ErrorHandler

    With Me.lvwFechamento
        .View = MSComctlLib.lvwReport
        .ColumnHeaders.Add , "VeículoLegal", "Veículo Legal", 2350
        .ColumnHeaders.Add , "DataCaixa", "Data Caixa", 1230, vbLeftJustify
        .ColumnHeaders.Add , "StatusCaixa", "Status Caixa", 1289, vbLeftJustify

        .ColumnHeaders.Add , "ValorAbertura", "Valor Abertura", 2069, vbRightJustify
        .ColumnHeaders.Add , "ValorMovimentacao", "Valor Movimentação", 2385, vbRightJustify
        .ColumnHeaders.Add , "ValorFechamento", "Valor de Fechamento", 2385, vbRightJustify
    End With

Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - flFormatarlvwFechamento"
End Sub

' Carrega veículos legais a serem fechados.

Private Sub flCarregarlvwFechamento(ByRef strDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objCaixaSubReserva  As MSSOAPLib30.SoapClient30
#Else
    Dim objCaixaSubReserva  As A6MIU.clsCaixaSubReserva
#End If

Dim xmlFechamento           As MSXML2.DOMDocument40
Dim xmlDomNode              As MSXML2.IXMLDOMNode
Dim lvwItemFechamento       As MSComctlLib.ListItem
Dim dblTotal                As Double
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If strDocFiltros = vbNullString Then
        frmMural.Display = "A seleção do Grupo de Veículo Legal é obrigatória. Por favor, clique em Definir Filtro, selecione um Grupo de Veículo Legal, e tente novamente."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If

    Set objCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")

    Set xmlFechamento = CreateObject("MSXML2.DOMDocument.4.0")
    xmlFechamento.loadXML (objCaixaSubReserva.ObterValoresFechamento(strDocFiltros, _
                                                                     vntCodErro, _
                                                                     vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If xmlFechamento.xml = vbNullString Then
        '100 - Documento XML Inválido.
        lvwFechamento.ListItems.Clear
        GoTo RemoverInstancias
    End If

    With Me.lvwFechamento
    
        lvwFechamento.ListItems.Clear

        For Each xmlDomNode In xmlFechamento.documentElement.childNodes
            Set lvwItemFechamento = .ListItems.Add(, "k_" & xmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                                                     "k_" & xmlDomNode.selectSingleNode("SG_SIST").Text, _
                                                            xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text)
            
            lvwItemFechamento.SubItems(1) = fgDtXML_To_Interface(xmlDomNode.selectSingleNode("DT_CAIX_DISP").Text)
            lvwItemFechamento.SubItems(2) = fgDescricaoEstadoCaixa(Val(xmlDomNode.selectSingleNode("CO_SITU_CAIX").Text))
            lvwItemFechamento.SubItems(3) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_UTLZ_ABER_CAIX").Text)
            lvwItemFechamento.SubItems(4) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
            
            dblTotal = fgVlrXml_To_Decimal(xmlDomNode.selectSingleNode("VA_UTLZ_ABER_CAIX").Text) + fgVlrXml_To_Decimal(xmlDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
            
            lvwItemFechamento.SubItems(5) = fgVlrXml_To_Interface(dblTotal)
                
       Next
    
    End With

RemoverInstancias:
    Set objCaixaSubReserva = Nothing
    Set xmlFechamento = Nothing
    Set xmlDomNode = Nothing
    Set lvwItemFechamento = Nothing
    
    Exit Sub

ErrorHandler:
    Set objCaixaSubReserva = Nothing
    Set xmlFechamento = Nothing
    Set xmlDomNode = Nothing
    Set xmlDomNode = Nothing
    Set lvwItemFechamento = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarlvwFechamento", Err.Number

End Sub

Private Sub lvwFechamento_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler
    
    Call ctlTableCombo.fgMakeButtonFlat
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "lvwFechamento_MouseMove"
    
End Sub

Private Sub objFiltro_AplicarFiltro(pxmlDocFiltros As String, lsTituloTableCombo As String)

On Error GoTo ErrorHandler
    
    fgCursor True
    Call flCarregarlvwFechamento(pxmlDocFiltros)
    xmlDocFiltros = pxmlDocFiltros
    fgCursor

    ctlTableCombo.TituloCombo = IIf(lsTituloTableCombo = vbNullString, strTableComboInicial, lsTituloTableCombo)
    
    Exit Sub
    
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - objFiltro_AplicarFiltro"
    
End Sub

Private Sub ctlTableCombo_AplicarFiltro(pxmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    fgCursor True
    
    If Trim(xmlDocFiltros) = vbNullString Then
        Call flCarregarlvwFechamento(pxmlDocFiltros)
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
                                        
        Call flCarregarlvwFechamento(xmlDomFiltro.xml)
        xmlDocFiltros = xmlDomFiltro.xml
    End If
    
    fgCursor
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaMonitoracao - ctlTableCombo_AplicarFiltro"

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
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - ctlTableCombo_MouseMove"

End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultado                            As String

On Error GoTo ErrorHandler

    Select Case Button.Key
    Case "fecharcaixa"
        strResultado = flFecharCaixa
        If strResultado <> vbNullString Then
            Call flMostrarResultado(strResultado)
            Call flCarregarlvwFechamento(xmlDocFiltros)
        End If
    
    Case "showfiltro"
        Set objFiltro = New frmFiltro
        Set objFiltro.FormOwner = Me
        objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaFechamento
        objFiltro.Show vbModal
        
    Case "refresh"
        fgCursor True
        flCarregarlvwFechamento xmlDocFiltros
        
    End Select
    
    fgCursor
    
    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - tlbButtons_ButtonClick"
    
End Sub
