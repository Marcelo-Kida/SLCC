VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultaMovimentacao 
   Caption         =   "Consulta - Movimentações Operações"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   14670
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstMovimentacao 
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Empresas"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sistemas"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Local de Liquidação"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tipo das Operações"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Situação das Operações"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Motivo Rejeição"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Quantidade"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Tipo BackOffice"
         Object.Width           =   3528
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
            Picture         =   "frmConsultaMovimentacao.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacao.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacao.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacao.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacao.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacao.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMovimentacao.frx":0F6C
            Key             =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8385
      Width           =   14670
      _ExtentX        =   25876
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
End
Attribute VB_Name = "frmConsultaMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
''-------------------------------------------------
' Gerado por Adrian Pretel
' Atualização em:      20-jun-2005
''-------------------------------------------------
''
Option Explicit

'Este objeto xmlMapaNavegacao é carregado com as propriedades do objRemessaRejeitada
'e todas as coleções que este for utilizar
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmConsultaMovimentacao"
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
    Set frmConsultaMovimentacao = Nothing
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
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConsultaMovimentacao", "flInicializar")
    Else

    End If

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    fgRaiseError App.EXEName, Me.Name, "frmConsultaMovimentacao - flInicializar", 0

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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMovimentacao - Form_KeyDown", Me.Caption

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
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaMovimentacao
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior

    Me.Refresh

    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMovimentacao - Form_Load", Me.Caption
    
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
                objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaMovimentacao
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
        .lstMovimentacao.Left = 0
        .lstMovimentacao.Width = .ScaleWidth
        .lstMovimentacao.Height = .ScaleHeight - .lstMovimentacao.Top - .tlbFiltro.Height
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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMovimentacao - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub lstMovimentacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   
On Error GoTo ErrorHandler
        
    Call fgClassificarListview(Me.lstMovimentacao, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lstMovimentacao_ColumnClick", Me.Caption

End Sub

Private Sub flCarregarFlexGrid(ByRef pxmlDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objMovimentacao As MSSOAPLib30.SoapClient30
#Else
    Dim objMovimentacao As A8MIU.clsConsultaMovimentacao
#End If

Dim xmlMovimentacao     As MSXML2.DOMDocument40
Dim xmlDomMovimentacao  As MSXML2.IXMLDOMNode
Dim strXMLRetorno       As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    gintRowPositionAnt = 0

    Set objMovimentacao = fgCriarObjetoMIU("A8MIU.clsConsultaMovimentacao")

    Set xmlMovimentacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objMovimentacao.Ler(pxmlDocFiltros, _
                                        vntCodErro, _
                                        vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lstMovimentacao.ListItems.Clear
    
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
    
    For Each xmlDomMovimentacao In xmlMovimentacao.documentElement.selectNodes("//Repeat_ConsultaMovimentacao/*")
        With lstMovimentacao.ListItems.Add(, , CStr(fgDtXML_To_Interface(xmlDomMovimentacao.selectSingleNode(".//DT_OPER_ATIV").Text)))
            .SubItems(1) = xmlDomMovimentacao.selectSingleNode(".//NO_EMPR").Text
            .SubItems(2) = xmlDomMovimentacao.selectSingleNode(".//NO_SIST").Text
            .SubItems(3) = xmlDomMovimentacao.selectSingleNode(".//CO_LOCA_LIQU").Text & " - " & xmlDomMovimentacao.selectSingleNode(".//DE_LOCA_LIQU").Text
            .SubItems(4) = xmlDomMovimentacao.selectSingleNode(".//NO_TIPO_OPER").Text
            .SubItems(5) = xmlDomMovimentacao.selectSingleNode(".//DE_SITU_PROC").Text
            .SubItems(6) = xmlDomMovimentacao.selectSingleNode(".//DE_SITU_PROC_MESG_SPB_RECB").Text
            .SubItems(7) = xmlDomMovimentacao.selectSingleNode(".//QUANTIDADE").Text
            .SubItems(8) = xmlDomMovimentacao.selectSingleNode("TP_BKOF").Text
            If Trim$(xmlDomMovimentacao.selectSingleNode("DE_BKOF").Text) <> vbNullString Then
                .SubItems(8) = .SubItems(8) & " - " & xmlDomMovimentacao.selectSingleNode("DE_BKOF").Text
            End If
        End With
    Next xmlDomMovimentacao
    
    Call fgClassificarListview(Me.lstMovimentacao, lngIndexClassifList, True)
    
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

    fgRaiseError App.EXEName, Me.Name, "frmConsultaMovimentacao - flCarregarFlexGrid", 0

End Sub
