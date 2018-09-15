VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControleFluxoAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parametrização de Atividade (Controle de Fluxo de Operações e Mensagens)"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLocalLiquidacao 
      Height          =   315
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   30
      Top             =   8040
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
            Picture         =   "frmControleFluxoAtividade.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":0112
            Key             =   "AtualizarExcecao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":5922
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":5F56
            Key             =   "Nova"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleFluxoAtividade.frx":6270
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   8580
      TabIndex        =   2
      Top             =   8490
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r    "
            Key             =   "sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3915
      Left            =   60
      TabIndex        =   3
      Top             =   510
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6906
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
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   3915
      Left            =   60
      TabIndex        =   4
      Top             =   4500
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6906
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
      Left            =   8970
      Top             =   90
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin VB.Label lblParamAtividade 
      AutoSize        =   -1  'True
      Caption         =   "Local de Liquidação"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1440
   End
End
Attribute VB_Name = "frmControleFluxoAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Controle de Ativação e Desativação de fluxo de recebimento de operações e mensagens

Option Explicit

Private Const COL_MESG_DESATIVAR            As Integer = 0
Private Const COL_MESG_CO_MESG              As Integer = 1
Private Const COL_MESG_NO_MESG              As Integer = 2

Private Const COL_OPER_DESATIVAR            As Integer = 0
Private Const COL_OPER_NO_TIPO_OPER         As Integer = 1
Private Const COL_OPER_TP_OPER              As Integer = 2
Private Const COL_OPER_TP_MESG_RECB_INTE    As Integer = 3

Private intControleMenuPopUp                As enumTipoConfirmacao

' Carregar combo Local de Liquidacao
Private Sub flCarregarComboLocalLiquidacao()

Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlLeitura, vbNullString, "Repeat_Filtro", vbNullString)
    Call fgAppendNode(xmlLeitura, "Repeat_Filtro", "TP_SEGR", "S")
    Call fgAppendNode(xmlLeitura, "Repeat_Filtro", "TP_VIGE", "S")
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsLocalLiquidacao", xmlLeitura))
    Call fgCarregarCombos(cboLocalLiquidacao, xmlLeitura, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU")
    Set xmlLeitura = Nothing

    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarComboLocalLiquidacao", 0

End Sub

' Carregar lista de mensagens existentes para o Local de Liquidação selecionado
Private Sub flCarregarListas()

Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim objListAux                              As MSComctlLib.ListView
Dim strLocalLiquidacao                      As String
Dim strItemKey                              As String

    On Error GoTo ErrorHandler
    
    strLocalLiquidacao = fgObterCodigoCombo(cboLocalLiquidacao.Text)
    If strLocalLiquidacao = vbNullString Then Exit Sub

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Select Case Val(strLocalLiquidacao)
        Case enumLocalLiquidacao.BMA
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "BMA")
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "LDL")
        
        Case enumLocalLiquidacao.BMC
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "BMC")
        
        Case enumLocalLiquidacao.BMD
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "LDL")
        
        Case enumLocalLiquidacao.CETIP
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "CTP")
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "LDL")
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "LTR")
        
        Case enumLocalLiquidacao.CLBCAcoes
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "LDL")
        
        Case enumLocalLiquidacao.SELIC
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "SEL")
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "RDC")
        
        Case enumLocalLiquidacao.SSTR
            Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_GrupoMensagem", vbNullString)
            Call fgAppendNode(xmlLeitura, "Grupo_GrupoMensagem", "GrupoMensagem", "STR")
        
    End Select
        
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodosTiposMensagens", "A8LQS.clsMensagem", xmlLeitura))
    
    With Me.lvwMensagem
        .ListItems.Clear
        
        For Each objXMLNode In xmlLeitura.selectNodes("Repeat_Mensagem/*")
            strItemKey = "|" & objXMLNode.selectSingleNode("CO_MESG").Text & _
                         "|" & strLocalLiquidacao & _
                         "|0"

            With .ListItems.Add(, strItemKey)
                .SubItems(COL_MESG_CO_MESG) = objXMLNode.selectSingleNode("CO_MESG").Text
                .SubItems(COL_MESG_NO_MESG) = objXMLNode.selectSingleNode("NO_MESG").Text
            End With
        Next
    End With
    
    Set xmlLeitura = Nothing
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_LocalLiquidacao", vbNullString)
    Call fgAppendNode(xmlLeitura, "Grupo_LocalLiquidacao", "LocalLiquidacao", strLocalLiquidacao)
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodosTiposOperacao", "A8LQS.clsOperacao", xmlLeitura))
    
    With Me.lvwOperacao
        .ListItems.Clear
        
        For Each objXMLNode In xmlLeitura.selectNodes("Repeat_TipoOperacao/*")
            strItemKey = "|" & _
                         "|" & strLocalLiquidacao & _
                         "|" & objXMLNode.selectSingleNode("TP_OPER").Text

            With .ListItems.Add(, strItemKey)
                .SubItems(COL_OPER_NO_TIPO_OPER) = objXMLNode.selectSingleNode("NO_TIPO_OPER").Text
                .SubItems(COL_OPER_TP_OPER) = objXMLNode.selectSingleNode("TP_OPER").Text
                .SubItems(COL_OPER_TP_MESG_RECB_INTE) = objXMLNode.selectSingleNode("TP_MESG_RECB_INTE").Text
            End With
        Next
    End With
    
    Set xmlLeitura = Nothing
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, vbNullString, "Grupo_LocalLiquidacao", vbNullString)
    Call fgAppendNode(xmlLeitura, "Grupo_LocalLiquidacao", "LocalLiquidacao", strLocalLiquidacao)
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodosFluxoAtividade", "A8LQS.clsWorkflow", xmlLeitura))
    
    For Each objXMLNode In xmlLeitura.selectNodes("Repeat_FluxoAtividade/*")
        strItemKey = "|" & objXMLNode.selectSingleNode("CO_MESG_SPB").Text & _
                     "|" & strLocalLiquidacao & _
                     "|" & objXMLNode.selectSingleNode("TP_OPER").Text

        Set objListAux = IIf(Split(strItemKey, "|")(1) = vbNullString, _
                             Me.lvwOperacao, _
                             Me.lvwMensagem)
                
        If fgExisteItemLvw(objListAux, strItemKey) Then
            objListAux.ListItems(strItemKey).Checked = True
        End If
    Next
    
    Set xmlLeitura = Nothing

    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListas", 0

End Sub

'Inicializar Grids (ListViews) de mensagens e operações
Private Sub flInicializarListViews()

    With lvwMensagem
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Desativar", 860
        .ColumnHeaders.Add , , "Mensagem", 1395
        .ColumnHeaders.Add , , "Descrição", 7780
    End With

    With lvwOperacao
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Desativar", 860
        .ColumnHeaders.Add , , "Tipo de Operação", 7830
        .ColumnHeaders.Add , , "Código", 690, lvwColumnRight
        .ColumnHeaders.Add , , "Layout", 660, lvwColumnRight
    End With

End Sub

'Monta string XML para processamento em lote
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemProc                             As MSXML2.DOMDocument40
Dim strLocalLiquidacao                      As String

    On Error GoTo ErrorHandler
    
    strLocalLiquidacao = fgObterCodigoCombo(cboLocalLiquidacao.Text)
    If strLocalLiquidacao = vbNullString Then
        With frmMural
            .Display = "Selecione o Local de Liquidação desejado."
            .Show vbModal
        End With
        Exit Function
    End If

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, vbNullString, "Repeat_Processamento", vbNullString)
    
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "Grupo_LocalLiquidacao", vbNullString)
    Call fgAppendNode(xmlProcessamento, "Grupo_LocalLiquidacao", "LocalLiquidacao", strLocalLiquidacao)
    Call fgAppendAttribute(xmlProcessamento, "Grupo_LocalLiquidacao", "Operacao", "ExcluirFluxoAtividade")
    Call fgAppendAttribute(xmlProcessamento, "Grupo_LocalLiquidacao", "Objeto", "A8LQS.clsWorkflow")

    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            If .Checked Then
                
                Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "CO_MESG_SPB", .SubItems(COL_MESG_CO_MESG))
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "CO_LOCA_LIQU", strLocalLiquidacao)
                Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "IncluirFluxoAtividade")
                Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", "A8LQS.clsWorkflow")
                        
                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemProc.xml)
                
                Set xmlItemProc = Nothing
                
            End If
        
        End With
    Next

    For Each objListItem In lvwOperacao.ListItems
        With objListItem
            If .Checked Then
                
                Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "TP_OPER", .SubItems(COL_OPER_TP_OPER))
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "CO_LOCA_LIQU", strLocalLiquidacao)
                Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "IncluirFluxoAtividade")
                Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", "A8LQS.clsWorkflow")
                        
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
    ElseIf xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 1 Then
        fgCursor
        If MsgBox("Confirma a ativação de todos os fluxos de recebimento de Operações e Mensagens para o local de liquidação selecionado ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            flMontarXMLProcessamento = xmlProcessamento.xml
        Else
            flMontarXMLProcessamento = vbNullString
        End If
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
        strXMLRetorno = fgMIUExecutarGenerico("ProcessarEmLote", "A8LQS.clsWorkflow", xmlProcessamento)
    End If

    flProcessar = strXMLRetorno

    Exit Function

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flProcessar", Me.Caption

End Function

Private Sub cboLocalLiquidacao_Click()

    On Error GoTo ErrorHandler
    
    fgCursor True
    Call flCarregarListas
    fgCursor
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "cboLocalLiquidacao_Click", 0

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

    On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(IIf(intControleMenuPopUp = enumTipoConfirmacao.Operacao, lvwOperacao, lvwMensagem), Retorno)
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
    Call flInicializarListViews
    Call flCarregarComboLocalLiquidacao
    fgCursor

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption

End Sub

Private Sub lvwMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.MENSAGEM
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

    Exit Sub
    
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_MouseDown", Me.Caption

End Sub

Private Sub lvwOperacao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.Operacao
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

    Exit Sub
    
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_MouseDown", Me.Caption

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
            
            If strResultadoOperacao <> vbNullString Then
                MsgBox "Atualização de cadastro efetuada com sucesso.", vbInformation, "Fluxo de Atividade"
            End If
    
    End Select
    
    fgCursor
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbCadastro_ButtonClick", Me.Caption

End Sub
