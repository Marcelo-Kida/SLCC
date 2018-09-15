VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAlteracaoEmpresaVeiculoLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Veículo Legal X Empresa"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   10560
   Begin VB.ComboBox cboEmpresaNova 
      Height          =   315
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   450
      Width           =   6855
   End
   Begin VB.ComboBox cboEmpresaAtual 
      Height          =   315
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   6855
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   0
      Top             =   7950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":0112
            Key             =   "AtualizarExcecao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":5922
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":5F56
            Key             =   "Nova"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoEmpresaVeiculoLegal.frx":6270
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   8550
      TabIndex        =   1
      Top             =   8400
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alterar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair   "
            Key             =   "sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwVeiculoLegal 
      Height          =   7455
      Left            =   60
      TabIndex        =   2
      Top             =   870
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   13150
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Excluir"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sistema"
         Object.Width           =   5028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Local Liquidação"
         Object.Width           =   5028
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Grupo Usuário"
         Object.Width           =   5028
      EndProperty
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   8940
      Top             =   0
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin VB.Label lblAdm 
      AutoSize        =   -1  'True
      Caption         =   "Nova Empresa"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   510
      Width           =   1050
   End
   Begin VB.Label lblAdm 
      AutoSize        =   -1  'True
      Caption         =   "Empresa Atual"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmAlteracaoEmpresaVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Possibilita ao usuário a alteração de empresas relacionadas à veículos legais.

Option Explicit

Private Const COL_VEICULO_LEGAL             As Integer = 0
Private Const COL_SIGLA_SISTEMA             As Integer = 1
Private Const COL_EMPRESA                   As Integer = 2
Private Const COL_GRUPO_VEICULO_LEGAL       As Integer = 3
Private Const COL_TIPO_BACKOFFICE           As Integer = 4
Private Const COL_NOME                      As Integer = 5
Private Const COL_NOME_REDUZIDO             As Integer = 6
Private Const COL_CNPJ                      As Integer = 7
Private Const COL_INDENTIFICADOR_CETIP      As Integer = 8
Private Const COL_CONTA_PADRAO_SELIC        As Integer = 9
Private Const COL_TIPO_TITULAR_BMA          As Integer = 10
Private Const COL_CODIGO_TITULAR_BMA        As Integer = 11
Private Const COL_DATA_INICIO_VIG           As Integer = 12
Private Const COL_DATA_FINAL_VIG            As Integer = 13

'Carregar Listview
Private Sub flCarregarListview()

Dim intEmpresa                              As Integer
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objXMLNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    intEmpresa = fgObterCodigoCombo(cboEmpresaAtual.Text)
    If intEmpresa = 0 Then
        frmMural.Display = "Selecione a Empresa Atual."
        frmMural.Show vbModal
        Exit Sub
    End If

    fgCursor True
    lvwVeiculoLegal.ListItems.Clear
    
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlLeitura, vbNullString, "Repeat_Filtros", vbNullString)
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_BancoLiquidante", vbNullString)
    Call fgAppendNode(xmlLeitura, "Grupo_BancoLiquidante", "BancoLiquidante", intEmpresa)
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("ObterDetalheVeiculoLegal", "A6A7A8.clsVeiculoLegal", xmlLeitura))
    
    For Each objXMLNode In xmlLeitura.selectNodes("Repeat_Repeat_VeiculoLegal/*")
        With lvwVeiculoLegal.ListItems.Add
            .Text = objXMLNode.selectSingleNode("CO_VEIC_LEGA").Text
            .SubItems(COL_SIGLA_SISTEMA) = objXMLNode.selectSingleNode("SG_SIST").Text
            .SubItems(COL_EMPRESA) = objXMLNode.selectSingleNode("CO_EMPR").Text & " - " & _
                                     objXMLNode.selectSingleNode("NO_EMPR").Text
            .SubItems(COL_GRUPO_VEICULO_LEGAL) = objXMLNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & " - " & _
                                                 objXMLNode.selectSingleNode("NO_GRUP_VEIC_LEGA").Text
            .SubItems(COL_TIPO_BACKOFFICE) = objXMLNode.selectSingleNode("TP_BKOF").Text & " - " & _
                                             objXMLNode.selectSingleNode("DE_BKOF").Text
            .SubItems(COL_NOME) = objXMLNode.selectSingleNode("NO_VEIC_LEGA").Text
            .SubItems(COL_NOME_REDUZIDO) = objXMLNode.selectSingleNode("NO_REDU_VEIC_LEGA").Text
            .SubItems(COL_CNPJ) = fgFormataCnpj(objXMLNode.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text)
            .SubItems(COL_INDENTIFICADOR_CETIP) = objXMLNode.selectSingleNode("ID_PART_CAMR_CETIP").Text
            .SubItems(COL_CONTA_PADRAO_SELIC) = objXMLNode.selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text
            .SubItems(COL_TIPO_TITULAR_BMA) = objXMLNode.selectSingleNode("TP_TITL_BMA").Text
            .SubItems(COL_CODIGO_TITULAR_BMA) = objXMLNode.selectSingleNode("CO_TITL_BMA").Text
            .SubItems(COL_DATA_INICIO_VIG) = fgDtXML_To_Interface(objXMLNode.selectSingleNode("DT_INIC_VIGE").Text)
            .SubItems(COL_DATA_FINAL_VIG) = fgDtXML_To_Interface(objXMLNode.selectSingleNode("DT_FIM_VIGE").Text)
        End With
    Next
    
    Set xmlLeitura = Nothing
    fgCursor
    
    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListview", 0

End Sub

'Inicializa controles de tela e variáveis
Private Sub flInicializarFormulario()

Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsEmpresa", xmlLeitura))
    
    Call fgCarregarCombos(Me.cboEmpresaAtual, xmlLeitura, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    Call fgCarregarCombos(Me.cboEmpresaNova, xmlLeitura, "Empresa", "CO_EMPR", "NO_REDU_EMPR")

    Call flInicializarListview
    
    Set xmlLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarFormulario", 0

End Sub

'Formata as colunas das listas da tela
Private Sub flInicializarListview()

    On Error GoTo ErrorHandler

    lvwVeiculoLegal.ListItems.Clear
    
    With Me.lvwVeiculoLegal.ColumnHeaders
        .Clear
        .Add , , "Veiculo Legal", 1310
        .Add , , "Sistema", 870
        .Add , , "Empresa", 2760
        .Add , , "Grupo Veículo Legal", 2040
        .Add , , "Tipo BackOffice", 1380
        .Add , , "Nome", 3075
        .Add , , "Nome Reduzido", 2489
        .Add , , "CNPJ do Veículo Legal", 1890
        .Add , , "Identificador Participante CETIP", 2459
        .Add , , "Conta Própria Custódia SELIC", 2489
        .Add , , "Tipo Titular BMA", 1725
        .Add , , "Código Titular BMA", 1725
        .Add , , "Data Início Vigência", 1666
        .Add , , "Data Fim Vigência", 1785
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarListview", 0

End Sub

'Monta string XML para processamento em lote
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemProc                             As MSXML2.DOMDocument40

Dim intEmpresaAtual                         As Integer
Dim intEmpresaNova                          As Integer

    On Error GoTo ErrorHandler
    
    intEmpresaAtual = Val(fgObterCodigoCombo(cboEmpresaAtual.Text))
    If intEmpresaAtual = 0 Then
        frmMural.Display = "Selecione a Empresa Atual."
        frmMural.Show vbModal
        Exit Function
    End If

    intEmpresaNova = Val(fgObterCodigoCombo(cboEmpresaNova.Text))
    If intEmpresaNova = 0 Then
        frmMural.Display = "Selecione a Empresa Nova."
        frmMural.Show vbModal
        Exit Function
    End If

    If intEmpresaAtual = intEmpresaNova Then
        frmMural.Display = "Empresa Atual é igual à Nova Empresa selecionada. Alteração não permitida."
        frmMural.Show vbModal
        Exit Function
    End If

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, vbNullString, "Repeat_Processamento", vbNullString)
    
    For Each objListItem In lvwVeiculoLegal.ListItems
        With objListItem
            If .Checked Then
                
                Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "CO_EMPR", intEmpresaNova)
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "CO_VEIC_LEGA", .Text)
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "SG_SIST", .SubItems(COL_SIGLA_SISTEMA))
        
                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemProc.xml)
                
                Set xmlItemProc = Nothing
                
            End If
        
        End With
    Next

    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLProcessamento = vbNullString
        With frmMural
            .Display = "Não existem itens selecionados para o processamento."
            .Show vbModal
        End With
    Else
        flMontarXMLProcessamento = xmlProcessamento.xml
    End If

    Set xmlProcessamento = Nothing

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLProcessamento", 0

End Function

'Enviar itens para processamento
Private Function flProcessar() As String

Dim strXMLRetorno                           As String
Dim xmlProcessamento                        As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlProcessamento.loadXML(flMontarXMLProcessamento)

    If xmlProcessamento.xml <> vbNullString Then
        strXMLRetorno = fgMIUExecutarGenerico("AlterarEmpresaVeiculoLegal", "A6A7A8.clsVeiculoLegal", xmlProcessamento)
    End If

    flProcessar = strXMLRetorno
    Set xmlProcessamento = Nothing

    Exit Function

ErrorHandler:
    Set xmlProcessamento = Nothing
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flProcessar", Me.Caption

End Function

Private Sub cboEmpresaAtual_Click()

    On Error GoTo ErrorHandler
    Call flCarregarListview
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "cboEmpresaAtual_Click", 0

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

    On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(Me.lvwVeiculoLegal, Retorno)
    End Select
    
    Exit Sub
    
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

    fgCursor True
    Call flInicializarFormulario
    fgCursor

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption

End Sub

Private Sub lvwVeiculoLegal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

    Exit Sub
    
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwVeiculoLegal_MouseDown", Me.Caption

End Sub
Private Sub lvwVeiculoLegal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwVeiculoLegal, ColumnHeader.Index)

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwVeiculoLegal_ColumnClick", Me.Caption

End Sub


Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
    
    On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "sair"
            Unload Me
        
        Case "salvar"
            fgCursor True
            strResultadoOperacao = flProcessar
            fgCursor
            
            If strResultadoOperacao <> vbNullString Then
                MsgBox "Atualização de cadastro efetuada com sucesso.", vbInformation, "Cadastro Veículo Legal X Empresa"
                Call flCarregarListview
            End If
    
    End Select
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbCadastro_ButtonClick", Me.Caption

End Sub
