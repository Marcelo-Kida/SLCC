VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReenvioCancelamentoEstornoMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ferramentas - Reenvio, Cancelamento e Estorno de Mensagem SPB"
   ClientHeight    =   7395
   ClientLeft      =   780
   ClientTop       =   2265
   ClientWidth     =   13215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   13215
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGeral 
      Height          =   915
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11805
      Begin VB.Frame Frame2 
         Caption         =   "Local Liquidação"
         Height          =   675
         Left            =   180
         TabIndex        =   8
         Top             =   120
         Width           =   4515
         Begin VB.ComboBox cboLocalLiquidacao 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ação"
         Height          =   675
         Left            =   4800
         TabIndex        =   7
         Top             =   120
         Width           =   4215
         Begin VB.OptionButton optReenvio 
            Caption         =   "Ree&nvio"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   300
            Width           =   1035
         End
         Begin VB.OptionButton optCancelamento 
            Caption         =   "&Cancelamento"
            Height          =   195
            Left            =   1320
            TabIndex        =   3
            Top             =   300
            Width           =   1395
         End
         Begin VB.OptionButton optEstorno 
            Caption         =   "&Estorno"
            Height          =   195
            Left            =   2760
            TabIndex        =   4
            Top             =   300
            Width           =   915
         End
      End
   End
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   10769
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
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Código da Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tipo Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cta Cedente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cta Cessionária"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Deb_Cred"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Título"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Dt.Vencto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Empresa"
         Object.Width           =   882
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   7065
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   582
      ButtonWidth     =   2725
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh &Tela"
            Key             =   "Atualizar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Definir Filtro"
            Key             =   "showfiltro"
            ImageKey        =   "showfiltro"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar         "
            Key             =   "ComandarAcao"
            Object.ToolTipText     =   "Complementar Operação"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r                 "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
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
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReenvioCancelamentoEstornoMsg.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   11760
      Top             =   0
      _ExtentX        =   2990
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmReenvioCancelamentoEstornoMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:35:08
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Reenvio, cancelamento ou estorno
'' de mensagens SPB) à camada controladora de caso de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsOperacao

Option Explicit

Public strHoraAgendamento                   As String

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private strOperacao                         As String
Private blnMensagem                         As Boolean

'Constantes de Configuração de Colunas
Private Const COL_DATA_OPERACAO             As Integer = 0
Private Const COL_NUMERO_COMANDO            As Integer = 1
Private Const COL_COD_MENSAGEM              As Integer = 2
Private Const COL_STATUS                    As Integer = 3
Private Const COL_TIPO_OPERACAO             As Integer = 4
Private Const COL_DATA_LIQUIDACAO           As Integer = 5
Private Const COL_CONTA_CEDENTE             As Integer = 6
Private Const COL_CONTA_CESSIONARIO         As Integer = 7
Private Const COL_DEBITO_CREDITO            As Integer = 8
Private Const COL_TITULO                    As Integer = 9
Private Const COL_VALOR                     As Integer = 10
Private Const COL_DATA_VENCIMENTO           As Integer = 11
Private Const COL_EMPRESA                   As Integer = 12

'Constantes de Configuração de Colunas do ListView de Mensagens

Private Const COL_MESG_COD_MENSAGEM         As Integer = 0
Private Const COL_MESG_DATA_HORA_MESG       As Integer = 1
Private Const COL_MESG_STATUS               As Integer = 2
Private Const COL_MESG_EMPRESA              As Integer = 3
Private Const COL_MESG_VALOR                As Integer = 4

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private Const strFuncionalidade             As String = "frmReenvioCancelamentoEstornoMsg"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private strAcao                             As String

Private lngIndexClassifList                 As Long

'Monta string XML para processamento em lote
Private Function flMontarXMLProcessamento(ByVal intLocalLiquidacao As Integer) As String

Dim objListItem                             As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemProc                             As MSXML2.DOMDocument40
Dim strObjeto                               As String

    On Error GoTo ErrorHandler
    
    If intLocalLiquidacao = enumLocalLiquidacao.SELIC Or _
       intLocalLiquidacao = enumLocalLiquidacao.CCR Then
        strObjeto = "A8LQS.clsOperacao"
    Else
        strObjeto = "A8LQS.clsMensagem"
    End If
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, vbNullString, "Repeat_Processamento", vbNullString)
    
    For Each objListItem In lstOperacao.ListItems
        With objListItem
            If .Checked Then
                
                Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "DH_ULTI_ATLZ", Split(.Tag, "|")(1))
        
                Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "Reenviar")
                Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", strObjeto)
                
                If intLocalLiquidacao = enumLocalLiquidacao.SELIC Or _
                   intLocalLiquidacao = enumLocalLiquidacao.CCR Then
                    Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_SEQU_OPER_ATIV", Mid$(.Key, 2))
                Else
                    Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_CTRL_IF", Mid$(.Key, 2))
                End If
                        
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

'Desmarca a seleção de todos os itens do ListView
Private Sub flMarcarDesmarcarTodas(ByVal lstListView As ListView, ByVal plngTipoSelecao As enumTipoSelecao)

Dim lngLinha                                As Long

    On Error GoTo ErrorHandler

    For lngLinha = 1 To lstListView.ListItems.Count
        lstListView.ListItems(lngLinha).Checked = (plngTipoSelecao = enumTipoSelecao.MarcarTodas)
    Next

    Exit Sub

ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMarcarDesmarcarTodas", 0

End Sub

Private Sub cboLocalLiquidacao_Click()

On Error GoTo ErrorHandler

    flLimpaCampos
    
    If cboLocalLiquidacao.Text = "5 - CETIP" Then
        'Permitir somente o reenvio de mensgens
        optReenvio.value = True
        optCancelamento.value = False
        optCancelamento.Enabled = False
        optEstorno.value = False
        optEstorno.Enabled = False
    ElseIf cboLocalLiquidacao.Text = "17 - BMA" Then
        optCancelamento.Enabled = True
        optEstorno.value = False
        optEstorno.Enabled = False
    ElseIf CLng(fgObterCodigoCombo(cboLocalLiquidacao.Text)) = enumLocalLiquidacao.CCR Then
        optReenvio.value = True
        optReenvio.Enabled = True
        optEstorno.Enabled = False
        optCancelamento.Enabled = False
    Else
        optCancelamento.Enabled = True
        optEstorno.Enabled = True
    End If

Exit Sub
ErrorHandler:

    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeiculoLegal - cboTipoBackOffice_Click", Me.Caption

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

    On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call flMarcarDesmarcarTodas(Me.lstOperacao, Retorno)
    End Select
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons(gstrAtualizar))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmReenvioCancelamentoEstornoMsg - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    Set objFiltro = New frmFiltro
    With objFiltro
        Set .FormOwner = Me
        .TipoFiltro = enumTipoFiltroA8.frmReenvioCancelamentoEstornoMsg
        Load objFiltro
    End With

    flLimpaCampos

    flInicializar

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

Exit Sub
ErrorHandler:
    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmReenvioCancelamentoEstornoMsg - Form_Load", Me.Caption
End Sub

'' Encaminhar a solicitação (Leitura de itens para o preenchimento do listview,
'' conforme o tipo selecionado (Reenvio, Cancelamento ou Estorno)) à camada
'' controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsOperacao.ObterDetalheOperacao
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flCarregarLista(ByVal pstrAcao As String, _
                            ByRef strXMLFiltros As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strSelecaoFiltro                        As String
Dim strAcaoMensagemSPB                      As String
Dim lngLocalLiquidacao                      As Long
Dim lngCont                                 As Long

Dim datDMenos1                              As Date
Dim datDmais1                               As Date
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Call flLimparLista

    lngLocalLiquidacao = CLng(fgObterCodigoCombo(cboLocalLiquidacao.Text))

    blnMensagem = False
    If lngLocalLiquidacao = enumLocalLiquidacao.BMA Then
        If pstrAcao <> "Cancelamento" Then
            If optReenvio.value = True Then
                flCarregarListaMensagem pstrAcao, strXMLFiltros, lngLocalLiquidacao
                blnMensagem = True
                Exit Sub
            End If
        End If
    ElseIf lngLocalLiquidacao = enumLocalLiquidacao.CETIP Then
        If optReenvio.value = True Then
            flCarregarListaMensagem pstrAcao, strXMLFiltros, lngLocalLiquidacao
            blnMensagem = True
            Exit Sub
        End If
    End If

    flFormataListView True

    Select Case pstrAcao
        Case "Cancelamento"
            strSelecaoFiltro = enumStatusOperacao.Pendencia & ";" & _
                               enumStatusOperacao.EmLancamento
            strAcaoMensagemSPB = enumTipoAcao.CancelamentoSolicitado
        Case "Estorno"
            strSelecaoFiltro = enumStatusOperacao.Liquidada
            strAcaoMensagemSPB = enumTipoAcao.EstornoSolicitado
        Case "Reenvio"
            strSelecaoFiltro = enumStatusOperacao.Rejeitada & ";" & _
                               enumStatusOperacao.Expirada & ";" & _
                               enumStatusOperacao.RejeitadaPiloto
    End Select

    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    'xmlDomFiltros.loadXML strXMLFiltros
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")

    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next

    If strAcaoMensagemSPB <> vbNullString Then
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_AcaoMensagemSPB", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_AcaoMensagemSPB", "AcaoMensagemSPB", strAcaoMensagemSPB)
    End If

    If pstrAcao = "Cancelamento" Or pstrAcao = "Estorno" Then
        fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoAcao", ""
        fgAppendNode xmlDomFiltros, "Grupo_TipoAcao", "TipoAcao", "0"
    End If

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", lngLocalLiquidacao)

    'Carlos 09/06/2004 - Apos conversa com Fabiana
    If pstrAcao = "Estorno" Then
        datDMenos1 = fgDataHoraServidor(DataAux)
        datDmais1 = fgDataHoraServidor(DataAux)
    Else
        
        If lngLocalLiquidacao = enumLocalLiquidacao.CCR Then
            datDMenos1 = fgDataHoraServidor(DataAux)
            datDmais1 = fgDataHoraServidor(DataAux)
        Else
            datDMenos1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 2, enumPaginacao.Anterior)
            datDmais1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, enumPaginacao.proximo)
        End If
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(datDMenos1)))
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(datDmais1)))
    '>>> -------------------------------------------------------------------------------------------

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(xmlDomFiltros.xml, _
                                                     vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing

    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarLista")
        End If

        For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheOperacao/*")
            With lstOperacao.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

                'Guarda na propriedade TAG o Status e a data da última atualização
                .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
    
                If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                    .Text = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                End If

                .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_TITULO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)

                If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                End If

                .SubItems(COL_CONTA_CEDENTE) = objDomNode.selectSingleNode("CO_CNTA_CEDT").Text
                .SubItems(COL_CONTA_CESSIONARIO) = objDomNode.selectSingleNode("CO_CNTA_CESS").Text

                'fixo SEL1023 pq BMA só pode cancelar ela - Pedido Mauricio 03/06/2004 - Carlos
                If lngLocalLiquidacao = enumLocalLiquidacao.BMA Then
                    If pstrAcao = "Cancelamento" Then
                        .SubItems(COL_COD_MENSAGEM) = "SEL1023"
                    Else
                        .SubItems(COL_COD_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                    End If
                Else
                    .SubItems(COL_COD_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                End If
                .SubItems(COL_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                .SubItems(COL_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                .SubItems(COL_DEBITO_CREDITO) = IIf(Val(objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito, "Crédito", "Débito")
                .SubItems(COL_EMPRESA) = objDomNode.selectSingleNode("CO_EMPR").Text

            End With
        Next
    End If

    Call fgClassificarListview(Me.lstOperacao, lngIndexClassifList, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0
    
End Sub

'Carrega a lista com as mensagens a serem tratadas
Private Sub flCarregarListaMensagem(ByVal pstrAcao As String, _
                                    ByRef strXMLFiltros As String, _
                                    ByRef plngLocalLiquidacao As Long)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim strSelecaoFiltro                        As String
Dim lngCont                                 As Long

Dim datDMenos1                              As Date
Dim datDmais1                               As Date
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    flLimparLista

    flFormataListView False

    strSelecaoFiltro = enumStatusMensagem.MensagemRejeitada & ";" & _
                       enumStatusMensagem.MensagemExpirada

    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")

'    xmlDomFiltros.loadXML strXMLFiltros
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")

    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", plngLocalLiquidacao)

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CodigoMensagemExcluir", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_CodigoMensagemExcluir", "CodigoMensagemExcluir", "LTR0001")
    Call fgAppendNode(xmlDomFiltros, "Grupo_CodigoMensagemExcluir", "CodigoMensagemExcluir", "LDL0001")

    datDMenos1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, enumPaginacao.Anterior)
    datDmais1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, enumPaginacao.proximo)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(datDMenos1)))
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(datDmais1)))
    '>>> -------------------------------------------------------------------------------------------

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(xmlDomFiltros.xml, _
                                                     vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing

    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaMensagem")
        End If

        'For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheOperacao/*")
        For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheMensagem/*")
            With lstOperacao.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_CTRL_IF").Text)

                'Guarda na propriedade TAG o Status e a data da última atualização
                .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text

                .Text = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                .SubItems(COL_MESG_DATA_HORA_MESG) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text)
                .SubItems(COL_MESG_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                .SubItems(COL_MESG_EMPRESA) = objDomNode.selectSingleNode("CO_EMPR").Text
                'Adrian - 31/01/06 - Para CETIP, o valor da Operação passa a buscar da informação da mensagem de conciliação, não da operação
                If fgVlrXml_To_Interface(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) = 5 Then
                    .SubItems(COL_MESG_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                Else
                    .SubItems(COL_MESG_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                End If

            End With
        Next
    End If

    Call fgClassificarListview(Me.lstOperacao, lngIndexClassifList, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaMensagem", 0
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With lstOperacao
        .Top = fraGeral.Height
        .Left = 0
        .Width = Me.Width - 100
        .Height = (Me.Height - fraGeral.Height) - 1000
    End With
    fraGeral.Width = Me.Width - 100

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objFiltro = Nothing
    Set frmReenvioCancelamentoEstornoMsg = Nothing

End Sub

Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ColumnClick"
End Sub

Private Sub lstOperacao_DblClick()

On Error GoTo ErrorHandler

    If Not lstOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            If blnMensagem Then
                .CodigoEmpresa = lstOperacao.SelectedItem.ListSubItems(COL_MESG_EMPRESA)
                .NumeroControleIF = Mid(lstOperacao.SelectedItem.Key, 2)
            Else
                .CodigoEmpresa = lstOperacao.SelectedItem.ListSubItems(COL_EMPRESA)
                .SequenciaOperacao = Mid(lstOperacao.SelectedItem.Key, 2)
            End If
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_DblClick"

End Sub

Private Sub lstOperacao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        If optReenvio.value Then
            ctlMenu1.ShowMenuMarcarDesmarcar
        End If
    End If

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lstOperacao_MouseDown", Me.Caption

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

On Error GoTo ErrorHandler

    Call flCarregarLista(strAcao, xmlDocFiltros)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - objFiltro_AplicarFiltro"
End Sub

Private Sub optCancelamento_Click()

On Error GoTo ErrorHandler

    If cboLocalLiquidacao.ListIndex = -1 Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Selecione um Local de Liquidação"
        frmMural.Show vbModal
        Exit Sub
    End If

    fgCursor True
    strAcao = "Cancelamento"
    objFiltro.fgCarregarPesquisaAnterior
    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - optCancelamento_Click"
End Sub

Private Sub optEstorno_Click()

On Error GoTo ErrorHandler

    If cboLocalLiquidacao.ListIndex = -1 Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Selecione um Local de Liquidação"
        frmMural.Show vbModal
        Exit Sub
    End If

    fgCursor True
    strAcao = "Estorno"
    objFiltro.fgCarregarPesquisaAnterior
    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - optEstorno_Click"
End Sub

'' Encaminhar a solicitação (Reenviar, cancelar ou estornar) à camada controladora
'' de caso de uso (componente / classe / metodos ) :
''
'' A8MIU.clsOperacao.Reenviar
'' A8MIU.clsOperacao.Cancelar
'' A8MIU.clsOperacao.Estornar
''
'' O método retornará uma String XML para a camada de interface.
''
Private Function flComandarAcao() As Boolean

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim lngCont                                 As Long
Dim lngTipoAcao                             As Long
Dim lngNumeroComando                        As Long
Dim lstItem                                 As MSComctlLib.ListItem
Dim strLabelAcao                            As String

Dim vntSequenciaOperacao                    As Variant
Dim intStatus                               As Integer
Dim strDHUltimaAtualizacao                  As String
Dim intLocalLiquidacao                      As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Const POS_STATUS                            As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1

    On Error GoTo ErrorHandler
    
    intLocalLiquidacao = Val(fgObterCodigoCombo(cboLocalLiquidacao.Text))
    If intLocalLiquidacao = 0 Then
        With frmMural
            .Display = "Selecione o Local de Liquidação desejado."
            .Show vbModal
        End With
        Exit Function
    End If
    
    If optReenvio.value = True Then
        Call flComandarReenvio(intLocalLiquidacao)
        Call flAtualizar
        Exit Function
    End If

    If Me.lstOperacao.SelectedItem Is Nothing Then
        With frmMural
            .Display = "Selecione pelo menos um item para o processamento."
            .Show vbModal
        End With
        Exit Function
    Else
        Set lstItem = Me.lstOperacao.SelectedItem
    End If
    
    vntSequenciaOperacao = Mid(lstItem.Key, 2)
    intStatus = Split(lstItem.Tag, "|")(POS_STATUS)
    strDHUltimaAtualizacao = Split(lstItem.Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO)

    Select Case True
        Case optCancelamento.value
            lngTipoAcao = enumTipoAcao.CancelamentoSolicitado
            strLabelAcao = "Cancelamento"
        Case optEstorno.value
            lngTipoAcao = enumTipoAcao.EstornoSolicitado
            strLabelAcao = "Estorno"
        Case optReenvio.value
            strLabelAcao = "Reenvio"
    End Select
    
    With frmIncluirNumComandoAcao
        .Acao = strLabelAcao
        .NumeroComando = lstItem.SubItems(COL_NUMERO_COMANDO)
        .Show vbModal
        lngNumeroComando = .NumeroComandoAcao
    End With
    Set frmIncluirNumComandoAcao = Nothing
    If lngNumeroComando = 0 Then
        Exit Function
    End If

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    If strLabelAcao = "Estorno" Then
        flComandarAcao = objOperacao.Estornar(vntSequenciaOperacao, _
                                              intStatus, _
                                              lngNumeroComando, _
                                              strDHUltimaAtualizacao, _
                                              vntCodErro, _
                                              vntMensagemErro)
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
    ElseIf strLabelAcao = "Cancelamento" Then
        flComandarAcao = objOperacao.Cancelar(vntSequenciaOperacao, _
                                              intStatus, _
                                              lngNumeroComando, _
                                              strDHUltimaAtualizacao, _
                                              vntCodErro, _
                                              vntMensagemErro)
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
    End If
    
    Set objOperacao = Nothing
    
    Call flAtualizar

    Exit Function

ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flComandarAcao", 0

End Function

'Comanda o reenvio de uma mensagem SPB
Private Function flComandarReenvio(ByVal intLocalLiquidacao As Integer) As Boolean

Dim strXMLRetorno                           As String
Dim xmlProcessamento                        As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlProcessamento.loadXML(flMontarXMLProcessamento(intLocalLiquidacao))

    If xmlProcessamento.xml <> vbNullString Then
        If intLocalLiquidacao = enumLocalLiquidacao.SELIC Or _
           intLocalLiquidacao = enumLocalLiquidacao.CCR Then
            strXMLRetorno = fgMIUExecutarGenerico("ProcessarEmLote", "A8LQS.clsOperacao", xmlProcessamento)
        Else
            strXMLRetorno = fgMIUExecutarGenerico("ProcessarEmLote", "A8LQS.clsMensagem", xmlProcessamento)
        End If
    End If

    Set xmlProcessamento = Nothing

    If strXMLRetorno <> vbNullString Then
        Call fgMostrarResultado(strXMLRetorno, "reenviados")
    End If

'    flProcessar = strXMLRetorno

    Exit Function

ErrorHandler:
    Set xmlProcessamento = Nothing
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flProcessar", Me.Caption

End Function

Private Sub optReenvio_Click()

On Error GoTo ErrorHandler

    If cboLocalLiquidacao.ListIndex = -1 Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Selecione um Local de Liquidação"
        frmMural.Show vbModal
        Exit Sub
    End If

    fgCursor True
    fgLockWindow Me.hwnd
    strAcao = "Reenvio"
    objFiltro.fgCarregarPesquisaAnterior
    fgLockWindow
    fgCursor False

Exit Sub
ErrorHandler:
    fgLockWindow
    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - optReenvio_Click"
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strSelecaoFiltro                        As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    fgCursor True
    
    Select Case Button.Key
        
        Case "ComandarAcao"
            flComandarAcao
        Case "showfiltro"
            fgCursor False
            If optCancelamento.value Or optReenvio.value Or optEstorno.value Then
                objFiltro.Show vbModal
            Else
                frmMural.Caption = Me.Caption
                frmMural.Display = "Selecione uma ação"
                frmMural.Show vbModal
            End If
        Case gstrAtualizar
            flAtualizar
        Case gstrSair
            Unload Me
            
    End Select
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmaçãoOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

'Atualiza o conteúdo dos campos
Private Sub flAtualizar()

On Error GoTo ErrorHandler

    Select Case True
        Case optCancelamento.value
            strAcao = "Cancelamento"
            objFiltro.fgCarregarPesquisaAnterior
        Case optEstorno.value
            strAcao = "Estorno"
            objFiltro.fgCarregarPesquisaAnterior
        Case optReenvio.value
            strAcao = "Reenvio"
            objFiltro.fgCarregarPesquisaAnterior
    End Select

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flAtualizar"
End Sub

'Formata as colunas do listview
Private Sub flFormataListView(ByVal pblnOperacao As Boolean)

On Error GoTo ErrorHandler

    With lstOperacao
        If pblnOperacao Then
            lstOperacao.ColumnHeaders.Add , , "Data Operação"
            lstOperacao.ColumnHeaders.Add , , "Número do Comando"
            lstOperacao.ColumnHeaders.Add , , "Código da Mensagem"
            lstOperacao.ColumnHeaders.Add , , "Status"
            lstOperacao.ColumnHeaders.Add , , "Tipo Operação"
            lstOperacao.ColumnHeaders.Add , , "Data Liquidação"
            lstOperacao.ColumnHeaders.Add , , "Cta Cedente"
            lstOperacao.ColumnHeaders.Add , , "Cta Cessionária"
            lstOperacao.ColumnHeaders.Add , , "Deb_Cred"
            lstOperacao.ColumnHeaders.Add , , "Título"
            lstOperacao.ColumnHeaders.Add , , "Valor"
            lstOperacao.ColumnHeaders.Add , , "Dt. Vencto"
            lstOperacao.ColumnHeaders.Add , , "Empresa"
        Else
            lstOperacao.ColumnHeaders.Add , , "Código da Mensagem"
            lstOperacao.ColumnHeaders.Add , , "Data/Hora Mensagem"
            lstOperacao.ColumnHeaders.Add , , "Status"
            lstOperacao.ColumnHeaders.Add , , "Empresa"
            lstOperacao.ColumnHeaders.Add , , "Valor"
        End If
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormataListView", 0

End Sub

'Inicializa controles e variáveis
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlLocalLiquidacao                      As MSXML2.DOMDocument40
Dim strLocalLiquidacao                      As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")

    strLocalLiquidacao = objMensagem.LerTodosDominioTabela("558", _
                                                           "PJ.TB_LOCAL_LIQUIDACAO", _
                                                           "", _
                                                           "", _
                                                           "", _
                                                           vntCodErro, _
                                                           vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing

    If Trim(strLocalLiquidacao) = vbNullString Then
        Exit Sub
    End If

    Set xmlLocalLiquidacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLocalLiquidacao.loadXML strLocalLiquidacao

    fgCarregarCombos Me.cboLocalLiquidacao, xmlLocalLiquidacao, "DominioTabela", "CODIGO", "SIGLA"

    Set xmlLocalLiquidacao = Nothing

Exit Sub
ErrorHandler:

    Set objMensagem = Nothing
    Set xmlLocalLiquidacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmReenvioCancelamentoEstornoMsg", "flInicializar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()

On Error GoTo ErrorHandler

    flLimparLista

    If cboLocalLiquidacao.ListIndex = -1 Then
        optCancelamento.Enabled = False
        optEstorno.Enabled = False
        optReenvio.Enabled = False
    Else
        optCancelamento.Enabled = True
        optEstorno.Enabled = True
        optReenvio.Enabled = True
        optCancelamento.value = False
        optEstorno.value = False
        optReenvio.value = False
    End If

Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, "frmReenvioCancelamentoEstornoMsg", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Limpa o conteúdo da listagem
Private Sub flLimparLista()

    On Error GoTo ErrorHandler

    With lstOperacao
        .ColumnHeaders.Clear
        .ListItems.Clear
        .CheckBoxes = IIf(optReenvio.value, True, False)
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "frmReenvioCancelamentoEstornoMsg", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub
