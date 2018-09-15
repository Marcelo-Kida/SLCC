VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultaAberturaFechamento 
   Caption         =   "Sub-reserva - Consulta Abertura / Fechamento"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9540
   WindowState     =   2  'Maximized
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   661
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   5070
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
            Picture         =   "frmConsultaAberturaFechamento.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":0224
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":0F6C
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAberturaFechamento.frx":13BE
            Key             =   "posterior"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   9540
      _ExtentX        =   16828
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
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
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
      EndProperty
   End
   Begin MSComctlLib.ListView lvwHistoricoCaixa 
      Height          =   3480
      Left            =   45
      TabIndex        =   3
      Tag             =   "Histórico Abertura Fechamento"
      Top             =   405
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   6138
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgToolBar"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Veículo Legal"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Situação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Abertura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Fechamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Movimentação"
         Object.Width           =   2540
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
      Width           =   4110
   End
End
Attribute VB_Name = "frmConsultaAberturaFechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a consulta ao histórico de abertura e fechamento do caixa.

Option Explicit

'Este objeto ObjDomDocument é carregado com as propriedades para o formulário
' e todas as coleções que este for utilizar
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

'Variaveis para a utilização do Filtro
Private strFiltroXML                        As String
Private blnPrimeiraConsulta                 As Boolean

Private Const strFuncionalidade             As String = "CONSULTAABERTURAFECHAMENTO"
Private Const strTableComboInicial          As String = "Grupos de Veículos Legais"

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

' Retorna descrição da situação do caixa.

Private Function flSituacaoCaixa(ByVal penumTipo As enumEstadoCaixa) As String

Dim strRetornaMensagem                      As String

On Error GoTo ErrorHandler

    Select Case penumTipo
       Case enumEstadoCaixa.Aberto
            strRetornaMensagem = "Aberto"
       Case enumEstadoCaixa.Fechado
            strRetornaMensagem = "Fechado"
       Case enumEstadoCaixa.Disponivel
            strRetornaMensagem = "Disponível"
    End Select

    flSituacaoCaixa = strRetornaMensagem

    Exit Function
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flSituacaoCaixa", 0
End Function

' Carrega lista com o histórico de abertura e fechamento do caixa.

Private Sub flCarregarList(ByRef pstrXMLDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objCaixaSubReserva  As MSSOAPLib30.SoapClient30
#Else
    Dim objCaixaSubReserva  As A6MIU.clsCaixaSubReserva
#End If

Dim xmlHistorico            As MSXML2.DOMDocument40
Dim xmlDomNode              As MSXML2.IXMLDOMNode
Dim strXMLRetorno           As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    gintRowPositionAnt = 0

    Set objCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")

    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objCaixaSubReserva.ObterHistoricoPosicaoCaixa(pstrXMLDocFiltros, _
                                                                  vntCodErro, _
                                                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lvwHistoricoCaixa.ListItems.Clear
    
    'caso a tabela esteja sem registros não tem como carregar um XML,
    'sendo assim vai para o fim da rotina.
    If Trim(strXMLRetorno) <> "" Then
       If Not xmlHistorico.loadXML(strXMLRetorno) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Exit Sub
    End If

    For Each xmlDomNode In xmlHistorico.documentElement.selectNodes("//Repeat_HistoricoPosicaoCaixa/*")
            
        With lvwHistoricoCaixa.ListItems.Add(, , xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text, , "showdetail")
            .SubItems(1) = fgDtXML_To_Date(xmlDomNode.selectSingleNode("DT_CAIX_SUB_RESE").Text)
            .SubItems(2) = flSituacaoCaixa(CLng(xmlDomNode.selectSingleNode("CO_SITU_CAIX_SUB_RESE").Text))
            .SubItems(3) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VALOR_ABERTURA").Text)
            .SubItems(4) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_FECH_CAIX_SUB_RESE").Text)
            .SubItems(5) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
        End With
        
    Next xmlDomNode
    
    Set xmlHistorico = Nothing
    Set objCaixaSubReserva = Nothing
    
    Exit Sub

ErrorHandler:
    Set objCaixaSubReserva = Nothing
    Set xmlHistorico = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "flCarregarList"

End Sub

Private Sub ctlTableCombo_AplicarFiltro(xmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCursor True
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Trim(strFiltroXML) = vbNullString Then
        fgAppendNode xmlDomFiltro, "", "Repeat_Filtros", ""
    Else
        Call xmlDomFiltro.loadXML(strFiltroXML)
    End If
    
    If Not xmlDomFiltro.selectSingleNode("//Grupo_GrupoVeiculoLegal") Is Nothing Then
        Call fgRemoveNode(xmlDomFiltro, "Grupo_GrupoVeiculoLegal")
    End If
    
    If xmlDomFiltro.selectSingleNode("//Grupo_Data") Is Nothing Then
        fgAppendNode xmlDomFiltro, "Repeat_Filtros", "Grupo_Data", ""
        fgAppendNode xmlDomFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)))
        fgAppendNode xmlDomFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle("99991231")
    End If
            
    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDomLeitura.loadXML(xmlDocFiltros)
    
    Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
    Call fgAppendNode(xmlDomFiltro, "Grupo_GrupoVeiculoLegal", "GrupoVeiculoLegal", _
                                    xmlDomLeitura.selectSingleNode("//GrupoVeiculoLegal").Text)
                                    
    Call flCarregarList(xmlDomFiltro.xml)
    strFiltroXML = xmlDomFiltro.xml
    
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
    mdiSBR.uctLogErros.MostrarErros Err, "ctlTableCombo_MouseMove"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        If Trim(strFiltroXML) = vbNullString Then Exit Sub
        
        fgCursor True
        Call flCarregarList(strFiltroXML)
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmConsultaAberturaFechamento - Form_KeyDown"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCenterMe Me
     
    Set Me.Icon = mdiSBR.Icon
    
    Me.Show
    DoEvents

    fgCursor True
    
    blnPrimeiraConsulta = True
    ctlTableCombo.TituloCombo = strTableComboInicial

    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaConsultaAberturaFechamento
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    objFiltro.Show vbModal
    
    blnPrimeiraConsulta = False
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .lblGrupoVeiculo.Left = 0
        .lblGrupoVeiculo.Top = 0
        .lblGrupoVeiculo.Width = .ScaleWidth
        
        .ctlTableCombo.Left = 0
        .ctlTableCombo.Top = 0
        
        .lvwHistoricoCaixa.Left = 0
        .lvwHistoricoCaixa.Top = .lblGrupoVeiculo.Height
        .lvwHistoricoCaixa.Width = .ScaleWidth
        .lvwHistoricoCaixa.Height = .ScaleHeight - .lvwHistoricoCaixa.Top - .tlbButtons.Height
    End With

End Sub

' Carrega configurações iniciais do formulário.

Private Sub flInit()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim strMapaNavegacao    As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    'strMapaNavegacao = objMIU.ObterMapaNavegacao(strFuncionalidade,vntCodErro,vntMensagemErro)
    Set objMIU = Nothing
    
    'If vntCodErro <> 0 Then
    '    GoTo ErrorHandler
    'End If

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConsultaAberturaFechamento", "flInit")
    Else

    End If

    Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    'Err.Number = vntCodErro
    'Err.Description = vntMensagemErro

    Call mdiSBR.uctLogErros.MostrarErros(Err, "flInit")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConsultaAberturaFechamento = Nothing
    gintRowPositionAnt = 0
End Sub

Private Sub lblGrupoVeiculo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmConsultaAberturaFechamento - lblGrupoVeiculo_MouseMove", Me.Caption
    
End Sub

Private Sub lvwHistoricoCaixa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    fgCursor True
    fgLockWindow Me.hwnd
    fgClassificarListview lvwHistoricoCaixa, ColumnHeader.Index
    fgLockWindow 0
    fgCursor
    
    Exit Sub
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmConsultaAberturaFechamento - lvwHistoricoCaixa_ColumnClick", Me.Caption
End Sub

Private Sub lvwHistoricoCaixa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat
    
    Exit Sub
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmConsultaAberturaFechamento - lblGrupoVeiculo_MouseMove", Me.Caption
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

On Error GoTo ErrorHandler

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            If blnPrimeiraConsulta Then
                blnPrimeiraConsulta = False
                
                'Call tlbButtons_ButtonClick(tlbButtons.Buttons("showfiltro"))
            Else
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
                
                Exit Sub
            End If
        End If
        
        fgCursor True
        Call flCarregarList(strFiltroXML)
        fgCursor
    
        ctlTableCombo.TituloCombo = IIf(strTituloTableCombo = vbNullString, strTableComboInicial, strTituloTableCombo)
        
    End If
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmConsultaAberturaFechamento - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
    Case "showfiltro"
        Set objFiltro = New frmFiltro
        Set objFiltro.FormOwner = Me
        objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaConsultaAberturaFechamento
        objFiltro.Show vbModal
    
    Case "refresh"
        If Trim(strFiltroXML) = vbNullString Then Exit Sub
        
        fgCursor True
        Call flCarregarList(strFiltroXML)
        fgCursor
    
    End Select
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmConsultaAberturaFechamento - tlbButtons_ButtonClick", Me.Caption
End Sub

