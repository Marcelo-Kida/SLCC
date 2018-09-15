VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSoliRepasseFincPagtoDespesas 
   Caption         =   "Solicitação de Repasse Financeiro / Pagamento de Despesas"
   ClientHeight    =   5760
   ClientLeft      =   810
   ClientTop       =   960
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   10785
   Begin MSComctlLib.ListView lvwConciliacao 
      Height          =   4635
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sequência"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Número Comando Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Indentificação do Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Preço Unitário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Número Comando Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Indentificação do Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Preço Unitário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Quantidade Conciliada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tipo Justificativa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Texto Justificativa"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox cboTipoOperacao 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   330
      Width           =   4350
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   4350
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   582
      ButtonWidth     =   2408
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela "
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   5475
      Top             =   10380
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
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoliRepasseFincPagtoDespesas.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Operação"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   90
      Width           =   615
   End
End
Attribute VB_Name = "frmSoliRepasseFincPagtoDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável por Disponibilizar Solicitação de Repasse Financeiro e Pagemtno de Despesas

Option Explicit

Private Const COL_CONC_SEQUENCIA            As Integer = 0

Private Const COL_OPER_NUMERO_COMANDO       As Integer = 1
Private Const COL_OPER_IDENT_ATIVO          As Integer = 2
Private Const COL_OPER_QTDE                 As Integer = 3
Private Const COL_OPER_PU                   As Integer = 4
Private Const COL_OPER_VALOR                As Integer = 5

Private Const COL_MESG_NUMERO_COMANDO       As Integer = 6
Private Const COL_MESG_IDENT_ATIVO          As Integer = 7
Private Const COL_MESG_QTDE                 As Integer = 8
Private Const COL_MESG_PU                   As Integer = 9
Private Const COL_MESG_VALOR                As Integer = 10

Private Const COL_CONC_QTDE                 As Integer = 11
Private Const COL_CONC_TIPO_JUST            As Integer = 12
Private Const COL_CONC_TEXTO_JUST           As Integer = 13

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLerTodos                         As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmSoliRepasseFincPagtoDespesas"

Private lngIndexClassifList                 As Long

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler

    fgCursor True
    flCarregarConciliacoes
    fgCursor False

Exit Sub
ErrorHandler:
    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmSoliRepasseFincPagtoDespesas - cboEmpresa_Click", Me.Caption

End Sub

Private Sub cboTipoOperacao_Click()

On Error GoTo ErrorHandler
    
    fgCursor True
    flCarregarConciliacoes
    fgCursor False

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmSoliRepasseFincPagtoDespesas - cboTipoOperacao_Click", Me.Caption
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCursor True
    Set Me.Icon = mdiLQS.Icon
    fgCenterMe Me

    Me.Show
    DoEvents

    flInicializar

    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    Call flCarregarComboTipoOperacao
    
    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor False

    mdiLQS.uctlogErros.MostrarErros Err, "frmSoliRepasseFincPagtoDespesas - Form_Load", Me.Caption

End Sub

'Inicializa o Form e os controles subjacentes
Public Function flInicializar() As Boolean

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")

    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConciliacaoOperacao", "flInicializar")
    End If

    Set objMIU = Nothing

Exit Function
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Function

'Carregar as conciliações que já foram efetuadas
Private Sub flCarregarConciliacoes()

#If EnableSoap = 1 Then
    Dim objConciliacao                      As MSSOAPLib30.SoapClient30
#Else
    Dim objConciliacao                      As A8MIU.clsOperacao
#End If

Dim strLerTodos                             As String
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lstItem                                 As MSComctlLib.ListItem

Dim lngTipoOperacao                         As Long
Dim lngEmpresa                              As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    If cboEmpresa.ListIndex > -1 Then
        lngEmpresa = fgObterCodigoCombo(cboEmpresa.Text)
    Else
        Exit Sub
    End If

    If cboTipoOperacao.ListIndex > -1 Then
        lngTipoOperacao = fgObterCodigoCombo(cboTipoOperacao.Text)
    Else
        Exit Sub
    End If

    Set objConciliacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strLerTodos = objConciliacao.ObterConciliacao(0, _
                                                  vbNullString, _
                                                  vbNullString, _
                                                  fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)), _
                                                  lngTipoOperacao, _
                                                  lngEmpresa, _
                                                  enumStatusOperacao.LiquidadaFisicamente, _
                                                  vntCodErro, _
                                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objConciliacao = Nothing

    If Trim$(strLerTodos) = vbNullString Then
        Exit Sub
    End If

    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlLerTodos.loadXML(strLerTodos) Then
        Call fgErroLoadXML(xmlLerTodos, App.EXEName, TypeName(Me), "flCarregarConciliacoes")
    End If

    lvwConciliacao.ListItems.Clear

    For Each objDomNode In xmlLerTodos.documentElement.childNodes

        With objDomNode
            Set lstItem = lvwConciliacao.ListItems.Add(, "k" & .selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text)

            lstItem.Text = .selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text

            lstItem.SubItems(COL_OPER_NUMERO_COMANDO) = .selectSingleNode("NU_COMD_OPER_OPER").Text
            lstItem.SubItems(COL_OPER_IDENT_ATIVO) = .selectSingleNode("NU_ATIV_MERC_OPER").Text
            lstItem.SubItems(COL_OPER_QTDE) = .selectSingleNode("QT_ATIV_MERC_OPER").Text
            lstItem.SubItems(COL_OPER_PU) = .selectSingleNode("PU_ATIV_MERC_OPER").Text
            lstItem.SubItems(COL_OPER_VALOR) = .selectSingleNode("VA_OPER_ATIV").Text

            lstItem.SubItems(COL_MESG_NUMERO_COMANDO) = .selectSingleNode("NU_COMD_OPER").Text
            lstItem.SubItems(COL_MESG_IDENT_ATIVO) = .selectSingleNode("NU_ATIV_MERC_MESG").Text
            lstItem.SubItems(COL_MESG_QTDE) = .selectSingleNode("QT_ATIV_MERC_MESG").Text
            lstItem.SubItems(COL_MESG_PU) = .selectSingleNode("PU_ATIV_MERC_MESG").Text
            lstItem.SubItems(COL_MESG_VALOR) = .selectSingleNode("VA_FINC").Text

            lstItem.SubItems(COL_CONC_QTDE) = .selectSingleNode("QT_ATIV_MERC_CNCL").Text
            lstItem.SubItems(COL_CONC_TIPO_JUST) = .selectSingleNode("NO_TIPO_JUST_CNCL").Text
            lstItem.SubItems(COL_CONC_TEXTO_JUST) = .selectSingleNode("TX_JUST").Text

        End With

    Next objDomNode

    Call fgClassificarListview(Me.lvwConciliacao, lngIndexClassifList, True)
    
Exit Sub
ErrorHandler:

    Set objConciliacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarConciliacoes", 0

End Sub

'Carrega o combo com os tipos de operação tratados nesta tela
Private Sub flCarregarComboTipoOperacao()

    cboTipoOperacao.AddItem "4- Operação Compromissada Volta Com Conciliação"
    cboTipoOperacao.ItemData(cboTipoOperacao.NewIndex) = 4

    cboTipoOperacao.AddItem "15- Eventos - Juros"
    cboTipoOperacao.ItemData(cboTipoOperacao.NewIndex) = 15

    cboTipoOperacao.AddItem "14- Eventos - Amortização"
    cboTipoOperacao.ItemData(cboTipoOperacao.NewIndex) = 14

    cboTipoOperacao.AddItem "13- Eventos - Resgate"
    cboTipoOperacao.ItemData(cboTipoOperacao.NewIndex) = 13

    cboTipoOperacao.AddItem "12- Despesas Selic"
    cboTipoOperacao.ItemData(cboTipoOperacao.NewIndex) = 12

End Sub

Private Sub Form_Resize()

On Error Resume Next

    With lvwConciliacao
        .Width = Me.Width - (100 + .Left)
        .Height = (Me.Height - tlbFiltro.Height) - (1000 + .Top)
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    Set xmlLerTodos = Nothing

End Sub

Private Sub lvwConciliacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwConciliacao, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwConciliacao_ColumnClick"
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
        Case "Refresh"
            flCarregarConciliacoes
        Case "Sair"
            Unload Me
    End Select

    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmSoliRepasseFincPagtoDespesas - tlbFiltro_ButtonClick", Me.Caption

End Sub
