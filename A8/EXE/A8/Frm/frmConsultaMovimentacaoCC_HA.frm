VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultaMovimentacaoCC_HA 
   Caption         =   "Consulta - Movimentações de Conta Corrente e Contabilidade"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstMovimentacaoCC_HA 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13996
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status da Integração"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo da Operação"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Local de Liquidação"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Empresa"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sistema"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Erro da Integração"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Justificativa"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Quantidade"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tipo BackOffice"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8430
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   582
      ButtonWidth     =   2434
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir &Filtro"
            Key             =   "Filtro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "Atualizar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   7800
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
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacaoCC_HA.frx":0F6C
            Key             =   "sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaMovimentacaoCC_HA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
''-------------------------------------------------
' Gerado por Adrian Pretel
' Atualização em:      24-jun-2005
''-------------------------------------------------
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmConsultaMovimentacaoCC_HA"
Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private Const COMPARE_ASC                   As Integer = 1
Private Const COMPARE_DESC                  As Integer = -1

Private intCompareCol                       As Integer
Private intCompareOrder                     As Integer

'Variaveis para a utilização do Filtro
Private strFiltroXML                        As String
Private blnUtilizaFiltro                    As Boolean
Private blnOrigemBotaoRefresh               As Boolean
Private blnPrimeiraConsulta                 As Boolean
Private intRefresh                          As Integer

Private lngCodigoErroNegocio                As Long

Private lngIndexClassifList                 As Long

Private Sub Form_Unload(Cancel As Integer)

    Set objFiltro = Nothing
    Set frmConsultaMovimentacaoCC_HA = Nothing
    gintRowPositionAnt = 0

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

'Inicializa o conteúdo dos controles de tela e variáveis
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao                        As String

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set objMIU = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConsultaMovimentacaoCC_HA", "flInicializar")
    Else

    End If

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    fgRaiseError App.EXEName, Me.Name, "frmConsultaMovimentacaoCC_HA - flInicializar", 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        fgLockWindow Me.hwnd
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons(gstrAtualizar))
        fgLockWindow 0
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    fgLockWindow 0
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMovimentacaoCC_HA - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    intCompareOrder = COMPARE_ASC
    
    fgLockWindow Me.hwnd
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    
    blnPrimeiraConsulta = True
    
    fgLockWindow 0

    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaMovimentacaoCC_HA
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior

    Me.Refresh

    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMovimentacaoCC_HA - Form_Load", Me.Caption
    
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
           Case "Sair"
                Unload Me
           
           Case gstrAtualizar
                If InStr(1, strFiltroXML, "DataIni") = 0 Then
                    frmMural.Caption = Me.Caption
                    frmMural.Display = "Obrigatória a seleção do filtro DATA."
                    frmMural.Show vbModal
                Else
                    Call fgCursor(True)
                    Call flCarregarFlexGrid(strFiltroXML)
                    Call fgCursor(False)
                End If
           
           Case Else
                blnPrimeiraConsulta = False
            
                Set objFiltro = New frmFiltro
                Set objFiltro.FormOwner = Me
                objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaMovimentacaoCC_HA
                objFiltro.Show vbModal
    
    End Select
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor True
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbFiltro_ButtonClick"
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With Me
        .lstMovimentacaoCC_HA.Left = 0
        .lstMovimentacaoCC_HA.Width = .ScaleWidth
        .lstMovimentacaoCC_HA.Height = .ScaleHeight - .lstMovimentacaoCC_HA.Top - .tlbFiltro.Height
    End With

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, _
                                    lsTituloTableCombo As String)

On Error GoTo ErrorHandler

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            blnPrimeiraConsulta = False
            Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("Filtro"))
        End If
        
        If InStr(1, strFiltroXML, "DataIni") = 0 Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
        Else
            Call fgCursor(True)
            Call flCarregarFlexGrid(strFiltroXML)
            Call fgCursor(False)
        End If
    
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMovimentacaoCC_HA - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub lstMovimentacaoCC_HA_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   
On Error GoTo ErrorHandler
        
    Call fgClassificarListview(Me.lstMovimentacaoCC_HA, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lstMovimentacaoCC_HA_ColumnClick", Me.Caption

End Sub

Private Sub flCarregarFlexGrid(ByRef pxmlDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objMovimentacao As MSSOAPLib30.SoapClient30
#Else
    Dim objMovimentacao As A8MIU.clsConsultaMovimentacao
#End If

Dim xmlMovimentacao     As MSXML2.DOMDocument40
Dim xmlLer              As MSXML2.DOMDocument40
Dim xmlDomMovimentacao  As MSXML2.IXMLDOMNode

Dim lerXMLMotivo        As MSXML2.DOMDocument40

Dim strXMLRetorno       As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    gintRowPositionAnt = 0

    Set objMovimentacao = fgCriarObjetoMIU("A8MIU.clsConsultaMovimentacao")
    Set lerXMLMotivo = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMovimentacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objMovimentacao.LerCC_HA(pxmlDocFiltros, _
                                             vntCodErro, _
                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lstMovimentacaoCC_HA.ListItems.Clear
    
    'caso a tabela esteja sem registros não tem como carregar um XML,
    'sendo assim vai para o fim da rotina.
    If Trim(strXMLRetorno) <> "" Then
       If Not xmlMovimentacao.loadXML(strXMLRetorno) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Call fgCursor(False)
       Exit Sub
    End If

    For Each xmlDomMovimentacao In xmlMovimentacao.documentElement.selectNodes("//Repeat_Erro/*")
        With lstMovimentacaoCC_HA.ListItems.Add(, , CStr(fgDtXML_To_Interface(xmlDomMovimentacao.selectSingleNode(".//DTINTEGRACAO").Text)))
            .SubItems(1) = xmlDomMovimentacao.selectSingleNode(".//STATUS").Text
            .SubItems(2) = xmlDomMovimentacao.selectSingleNode(".//DESCTPOPERACAO").Text
            .SubItems(3) = xmlDomMovimentacao.selectSingleNode(".//LOCALLIQUIDACAO").Text & _
                           " - " & _
                           xmlDomMovimentacao.selectSingleNode(".//DESCLOCALIQUIDACAO").Text
            .SubItems(4) = xmlDomMovimentacao.selectSingleNode(".//TPEMPRESA").Text & _
                           " - " & _
                           xmlDomMovimentacao.selectSingleNode(".//EMPRESA").Text
            .SubItems(5) = xmlDomMovimentacao.selectSingleNode(".//TPSISTEMA").Text & _
                           " - " & _
                           xmlDomMovimentacao.selectSingleNode(".//SISTEMA").Text
            .SubItems(6) = xmlDomMovimentacao.selectSingleNode(".//TPERRO").Text & _
                           " - " & _
                           xmlDomMovimentacao.selectSingleNode(".//ERRO").Text
            .SubItems(7) = xmlDomMovimentacao.selectSingleNode(".//JUSTIFICATIVA").Text
            .SubItems(8) = xmlDomMovimentacao.selectSingleNode(".//QUANTIDADE").Text
            .SubItems(9) = xmlDomMovimentacao.selectSingleNode("TP_BKOF").Text
            If Trim$(xmlDomMovimentacao.selectSingleNode("DE_BKOF").Text) <> vbNullString Then
                .SubItems(9) = .SubItems(9) & " - " & xmlDomMovimentacao.selectSingleNode("DE_BKOF").Text
            End If
        End With
    Next xmlDomMovimentacao
    
    Call fgClassificarListview(Me.lstMovimentacaoCC_HA, lngIndexClassifList, True)
    
    Set xmlMovimentacao = Nothing
    Set objMovimentacao = Nothing
    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Set xmlMovimentacao = Nothing
    Set objMovimentacao = Nothing
    Call fgCursor(False)
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, Me.Name, "frmConsultaMovimentacaoCC_HA - flCarregarFlexGrid", 0

End Sub



