VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIntegracaoCCBatch 
   Caption         =   "Conta Corrente - Integração Contabilidade"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   14370
   WindowState     =   2  'Maximized
   Begin VB.Frame fraEmpresa 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   6090
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   5595
      End
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   582
      ButtonWidth     =   1958
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Integrar"
            Key             =   "Integrar"
            ImageKey        =   "Integrar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "refresh"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r      "
            Key             =   "Sair"
            ImageKey        =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   135
      Top             =   5265
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
            Picture         =   "frmIntegracaoCCBatch.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":1286
            Key             =   "Integrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracaoCCBatch.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwContaCorrente 
      Height          =   7005
      Left            =   0
      TabIndex        =   3
      Top             =   900
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   12356
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   26
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Empresa"
         Object.Width           =   5212
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sistema"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Operação"
         Object.Width           =   2357
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Número Comando"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Veiculo Legal"
         Object.Width           =   5980
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Situação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tipo Operação"
         Object.Width           =   5477
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Local Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Banco"
         Object.Width           =   5927
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Agência"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Número C/C"
         Object.Width           =   2194
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Valor Lançamento"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tipo Movto."
         Object.Width           =   1850
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Tipo Lançamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Sub-tipo Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Conta Contábil Débito"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Conta Contábil Crédito"
         Object.Width           =   3281
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Código Histórico Contábil"
         Object.Width           =   3544
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Descrição Histórico Contábil"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Codigo Veiculo Legal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Codigo Situação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Codigo Tipo Operação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Codigo Banco"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Codigo Local Liquidação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Condicão Net Operações"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Canal Venda"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmIntegracaoCCBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Possibilita a integração Batch de Lançamentos em Conta Corrente

Option Explicit

'Constantes de Configuração de Colunas
Private Const COL_EMPRESA                   As Integer = 0
Private Const COL_SISTEMA                   As Integer = 1
Private Const COL_DATA_OPERACAO             As Integer = 2
Private Const COL_NUMERO_COMANDO            As Integer = 3
Private Const COL_VEICULO_LEGAL             As Integer = 4
Private Const COL_SITUACAO                  As Integer = 5
Private Const COL_TIPO_OPERACAO             As Integer = 6
Private Const COL_LOCA_LIQU                 As Integer = 7
Private Const COL_BANCO                     As Integer = 8
Private Const COL_AGENCIA                   As Integer = 9
Private Const COL_CONTA_CORRENTE            As Integer = 10
Private Const COL_VALOR                     As Integer = 11
Private Const COL_TIPO_MOVIMENTO            As Integer = 12
Private Const COL_TIPO_LANCAMENTO           As Integer = 13
Private Const COL_SUB_TIPO_ATIVO            As Integer = 14
Private Const COL_CONTA_CONTABIL_DEB        As Integer = 15
Private Const COL_CONTA_CONTABIL_CRED       As Integer = 16
Private Const COL_COD_HIST_CONTABIL         As Integer = 17
Private Const COL_DES_HIST_CONTABIL         As Integer = 18
Private Const COL_COD_VEIC_LEGA             As Integer = 19
Private Const COL_COD_SITUACAO              As Integer = 20
Private Const COL_COD_TIPO_OPER             As Integer = 21
Private Const COL_COD_BANCO                 As Integer = 22
Private Const COL_COD_LOCA_LIQU             As Integer = 23
Private Const COL_COND_NET_OPERACOES        As Integer = 24
Private Const COL_CANAL_VENDA               As Integer = 25

'Constantes de erros de negócio específicos
Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3043
Private Const COD_ERRO_NEGOCIO_ESTORNO      As Long = 3044

'Constantes de posicionamento de campos na propriedade Key do item do ListView
Private Const POS_NU_SEQU_OPER_ATIV         As Integer = 1
Private Const POS_TP_LANC_ITGR              As Integer = 2
Private Const POS_DH_ULTI_ATLZ              As Integer = 3
Private Const POS_NR_SEQU_LANC              As Integer = 4

Private Const KEY_MSG_CO_LOCA_LIQU          As Integer = 1
Private Const KEY_MSG_CO_ISPB_CNPT          As Integer = 2
Private Const KEY_MSG_CO_CNPJ_CNPT          As Integer = 3
Private Const KEY_MSG_CO_AGEN_COTR          As Integer = 4
Private Const KEY_MSG_NU_CC_COTR            As Integer = 5
Private Const KEY_MSG_CO_ULTI_SITU_PROC     As Integer = 6
Private Const KEY_MSG_TP_IF_CRED_DEBT       As Integer = 7

Private Const KEY_EMPRESA                   As Integer = 1
Private Const KEY_DATA_OPERACAO             As Integer = 2
Private Const KEY_TIPO_OPERACAO             As Integer = 3
Private Const KEY_VEICULO_LEGAL             As Integer = 4
Private Const KEY_LOCA_LIQU                 As Integer = 5
Private Const KEY_BANCO                     As Integer = 6
Private Const KEY_AGENCIA                   As Integer = 7
Private Const KEY_CONTA_CORRENTE            As Integer = 8
Private Const KEY_CO_ULTI_SITU_PROC         As Integer = 9

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmIntegracaoCCBatch"

Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlDomLeitura                       As MSXML2.DOMDocument40
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifList                 As Long

'Carrega o controle com a lista de Empresas
Private Function flCarregarComboEmpresa() As Boolean

Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    For Each xmlNode In xmlMapaNavegacao.selectSingleNode("frmIntegracaoCCBatch/Grupo_Dados/Repeat_Empresa").childNodes
        cboEmpresa.AddItem xmlNode.selectSingleNode("CO_EMPR").Text & " - " & xmlNode.selectSingleNode("NO_REDU_EMPR").Text
        cboEmpresa.ItemData(cboEmpresa.NewIndex) = CLng(xmlNode.selectSingleNode("CO_EMPR").Text)
    Next
        
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboEmpresa", 0
    
End Function

'Carrega a lista de lançamentos
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsContaCorrente
#End If

'Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem

Dim strFiltro                               As String
Dim strListItemKey                          As String
Dim strListItemKey2                         As String
Dim lngCont                                 As Long
Dim strRetLeitura                           As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim blnCorretoras                           As Boolean
Dim dblValorOperacao                        As Double
Dim intDebitoCredito                        As Integer

On Error GoTo ErrorHandler

    lvwContaCorrente.ListItems.Clear
        
    If cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
            
    strFiltro = flMontarXmlFiltro()

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
    strRetLeitura = objOperacao.ObterDetalheLancamento(strFiltro, _
                                                       vntCodErro, _
                                                       vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheLancamento/*")
            'Cesar 10/05/2007 - Conta Corrente Corretoras
            blnCorretoras = False
            If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                strListItemKey = "|" & objDomNode.selectSingleNode("CO_EMPR").Text & _
                                 "|" & objDomNode.selectSingleNode("DT_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("TP_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_BANC").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_AGEN").Text & _
                                 "|" & objDomNode.selectSingleNode("NU_CC").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & objDomNode.selectSingleNode("NR_SEQU_LANC").Text
                
                dblValorOperacao = flNetOperacoesBatch(strListItemKey)
                
                For Each objListItem In lvwContaCorrente.ListItems
                      strListItemKey2 = "|" & fgObterCodigoCombo(objListItem.Text) & _
                                        "|" & fgDt_To_Xml(objListItem.SubItems(COL_DATA_OPERACAO)) & _
                                        "|" & objListItem.SubItems(COL_COD_TIPO_OPER) & _
                                        "|" & objListItem.SubItems(COL_COD_VEIC_LEGA) & _
                                        "|" & objListItem.SubItems(COL_COD_LOCA_LIQU) & _
                                        "|" & objListItem.SubItems(COL_COD_BANCO) & _
                                        "|" & objListItem.SubItems(COL_AGENCIA) & _
                                        "|" & objListItem.SubItems(COL_CONTA_CORRENTE) & _
                                        "|" & objListItem.SubItems(COL_COD_SITUACAO)
                    
                    If strListItemKey = strListItemKey2 Then
                        blnCorretoras = True
                        objListItem.SubItems(COL_COND_NET_OPERACOES) = objListItem.SubItems(COL_COND_NET_OPERACOES) & _
                                                                       "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                        Exit For
                    End If
                Next
                
                If dblValorOperacao > 0 Then
                    intDebitoCredito = enumTipoDebitoCredito.Debito
                Else
                    intDebitoCredito = enumTipoDebitoCredito.Credito
                End If
            End If
            
            If blnCorretoras = False Then
                With lvwContaCorrente.ListItems.Add(, _
                        "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & _
                        "|" & objDomNode.selectSingleNode("TP_LANC_ITGR").Text & _
                        "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                        "|" & objDomNode.selectSingleNode("NR_SEQU_LANC").Text)
                    
                    'Empresa
                    If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                        'Obtem a descrição da Empresa via QUERY XML
                        .Text = _
                                objDomNode.selectSingleNode("CO_EMPR").Text & " - " & xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                                objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                    End If
                    
                    'Sistema
                    .SubItems(COL_SISTEMA) = objDomNode.selectSingleNode("SG_SIST").Text & " - " & objDomNode.selectSingleNode("NO_SIST").Text
                    
                    'Data Operação
                    If objDomNode.selectSingleNode("DT_OPER").Text <> gstrDataVazia Then
                        .SubItems(COL_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER").Text)
                    End If
                    
                    'Número do Comando
                    .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    
                    'Veiculo Legal
                    .SubItems(COL_VEICULO_LEGAL) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_COD_VEIC_LEGA) = objDomNode.selectSingleNode("CO_VEIC_LEGA").Text
                    
                    'Situação
                    .SubItems(COL_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_COD_SITUACAO) = objDomNode.selectSingleNode("CO_SITU_PROC").Text
                    
                    'Tipo de Operação
                    .SubItems(COL_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                    .SubItems(COL_COD_TIPO_OPER) = objDomNode.selectSingleNode("TP_OPER").Text
                    
                    'Local de Liquidação
                    If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                        
                        If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                    objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU") Is Nothing Then
                            
                            'Obtem a descrição do Local de Liquidação via QUERY XML
                            .SubItems(COL_LOCA_LIQU) = _
                                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                    objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU").Text
                    
                        Else
                            
                            vntCodErro = 5
                            vntMensagemErro = "Usuário não possui acesso ao Local de Liquidação " & _
                                              objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "."
                            GoTo ErrorHandler
                            
                        End If
                    
                    End If
                    .SubItems(COL_COD_LOCA_LIQU) = objDomNode.selectSingleNode("CO_LOCA_LIQU").Text
    
                    'Banco
                    .SubItems(COL_BANCO) = objDomNode.selectSingleNode("CO_BANC").Text
                    .SubItems(COL_COD_BANCO) = objDomNode.selectSingleNode("CO_BANC").Text
                    
                    'Agência
                    .SubItems(COL_AGENCIA) = objDomNode.selectSingleNode("CO_AGEN").Text
                    
                    'Número C/C
                    .SubItems(COL_CONTA_CORRENTE) = objDomNode.selectSingleNode("NU_CC").Text
                    
                    If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                        'Valor do Lançamento
                        .SubItems(COL_VALOR) = fgVlrXml_To_Interface(fgVlr_To_Xml(Abs(dblValorOperacao)))
                        
                        'Tipo Movto.
                        .SubItems(COL_TIPO_MOVIMENTO) = IIf(intDebitoCredito = enumTipoDebitoCredito.Debito, "Débito", "Crédito")
                    Else
                        'Valor do Lançamento
                        .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_LANC_CC").Text)
                        
                        'Tipo Movto.
                        .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text
                    End If
                                    
                    'Tipo de Lançamento
                    .SubItems(COL_TIPO_LANCAMENTO) = IIf(Val(objDomNode.selectSingleNode("TP_LANC_ITGR").Text) = enumTipoLancamentoIntegracao.Estorno, "Estorno", "Normal")
                    
                    'Sub-tipo Ativo
                    .SubItems(COL_SUB_TIPO_ATIVO) = objDomNode.selectSingleNode("CO_SUB_TIPO_ATIV").Text
                    
                    'Conta Contábil Débito
                    .SubItems(COL_CONTA_CONTABIL_DEB) = objDomNode.selectSingleNode("CO_CNTA_DEBT").Text
                    
                    'Conta Contábil Crédito
                    .SubItems(COL_CONTA_CONTABIL_CRED) = objDomNode.selectSingleNode("CO_CNTA_CRED").Text
                    
                    'Código Histórico Contábil
                    .SubItems(COL_COD_HIST_CONTABIL) = objDomNode.selectSingleNode("CO_HIST_CNTA_CNTB").Text
                    
                    'Descriçao do Histórico contábil
                    .SubItems(COL_DES_HIST_CONTABIL) = objDomNode.selectSingleNode("DE_HIST_CNTA_CNTB").Text
                    
                    .SubItems(COL_COND_NET_OPERACOES) = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                    
                    'KIDA - SGC
                    .SubItems(COL_CANAL_VENDA) = fgDescricaoCanalVenda(objDomNode.selectSingleNode("TP_CNAL_VEND").Text)
                    
                End With
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lvwContaCorrente, lngIndexClassifList, True)
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing

    Exit Sub

ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
   
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0

End Sub

'Gerencia os lançamentos passíveis de processamento
Private Function flGerenciar() As String

#If EnableSoap = 1 Then
    Dim objContaCorrente                    As MSSOAPLib30.SoapClient30
#Else
    Dim objContaCorrente                    As A8MIU.clsContaCorrente
#End If

Dim xmlLoteLancamentos                      As MSXML2.DOMDocument40
Dim strXMLRetorno                           As String
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long
Dim intIgnoraGradeHorario                   As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set xmlLoteLancamentos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLoteLancamentos, "", "Repeat_Filtros", "")
    
    With lvwContaCorrente.ListItems
        For lngCont = 1 To .Count
            lngItensChecked = lngItensChecked + 1
            
            Call fgAppendNode(xmlLoteLancamentos, "Repeat_Filtros", "Grupo_Lote", "")
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "NU_SEQU_OPER_ATIV", Split(.Item(lngCont).Key, "|")(POS_NU_SEQU_OPER_ATIV), "Repeat_Filtros")

            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "TP_LANC_ITGR", Split(.Item(lngCont).Key, "|")(POS_TP_LANC_ITGR), "Repeat_Filtros")

            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "TipoLancamentoIntegracao", Split(.Item(lngCont).Key, "|")(POS_TP_LANC_ITGR), "Repeat_Filtros")
                       
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "CO_ULTI_SITU_PROC", enumStatusIntegracao.Integrado, "Repeat_Filtros")
                
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "DH_ULTI_ATLZ", Split(.Item(lngCont).Key, "|")(POS_DH_ULTI_ATLZ), "Repeat_Filtros")
            
            'Cesar 08/05/2007 - Conta Corrente Corretoras
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "VA_LANC_CC", fgVlr_To_Xml(.Item(lngCont).SubItems(COL_VALOR)), "Repeat_Filtros")
                
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "IN_LANC_DEBT_CRED", IIf(.Item(lngCont).SubItems(COL_TIPO_MOVIMENTO) = "Débito", _
                                                             enumTipoDebitoCredito.Debito, _
                                                             enumTipoDebitoCredito.Credito), "Repeat_Filtros")
            
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "NetOperacoes", .Item(lngCont).SubItems(COL_COND_NET_OPERACOES), "Repeat_Filtros")
            
            intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed And _
                                        .Item(lngCont).Tag = COD_ERRO_NEGOCIO_GRADE, 1, 0)
            
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "IgnoraGradeHorario", intIgnoraGradeHorario, "Repeat_Filtros")
        
            Call fgAppendNode(xmlLoteLancamentos, _
                      "Grupo_Lote", "NR_SEQU_LANC", Split(.Item(lngCont).Key, "|")(POS_NR_SEQU_LANC), "Repeat_Filtros")

        Next
    End With
    
    If lngItensChecked > 0 Then
        Set objContaCorrente = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
        strXMLRetorno = objContaCorrente.Gerenciar(xmlLoteLancamentos.xml, _
                                                   vntCodErro, _
                                                   vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objContaCorrente = Nothing
        
        'Verifica se o retorno da operação possui erros
        If strXMLRetorno <> vbNullString Then
            '...se sim, carrega o XML de Erros
            Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlRetornoErro.loadXML(strXMLRetorno)
        Else
            '...se não, apenas destrói o objeto
            Set xmlRetornoErro = Nothing
        End If
        
        flGerenciar = strXMLRetorno
    Else
        flGerenciar = vbNullString
    End If
    
    Set xmlLoteLancamentos = Nothing

Exit Function
ErrorHandler:

    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    fgRaiseError App.EXEName, Me.Name, "flGerenciar", 0
    'mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flGerenciar", Me.Caption

End Function

'Verifica se há lançamentos suspensos
Private Function flExisteLancamentoDMenos1Suspenso() As Boolean

#If EnableSoap = 1 Then
    Dim objContaCorrente                    As MSSOAPLib30.SoapClient30
#Else
    Dim objContaCorrente                    As A8MIU.clsContaCorrente
#End If

Dim strRetLeitura                           As String
Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim datDMenos1                              As Date
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltro, "", "Repeat_Filtros", "")
    
    'Filtro Empresa
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", cboEmpresa.ItemData(cboEmpresa.ListIndex))
        
    'Filtro Datas
    datDMenos1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, enumPaginacao.Anterior)
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(datDMenos1)))
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(datDMenos1)))
     
    'Fltro Status
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Status", "Status", enumStatusIntegracao.Suspenso)
      
    Set objContaCorrente = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
    strRetLeitura = objContaCorrente.ObterDetalheLancamento(xmlFiltro.xml, _
                                                            vntCodErro, _
                                                            vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objContaCorrente = Nothing
     
    If strRetLeitura <> vbNullString Then
        flExisteLancamentoDMenos1Suspenso = True
    End If
     
    Set xmlFiltro = Nothing
            
Exit Function
ErrorHandler:
    Set xmlFiltro = Nothing
    Set objContaCorrente = Nothing
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, Me.Name, "flExisteLancamentoDMenos1Suspenso", 0

End Function

'Inicializa controles de tela e variáveis
Private Sub flInicializar()

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
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmIntegracaoCCBatch", "flInicializar")
    End If
    
    flCarregarComboEmpresa
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Exibe de maneira diferenciada os itens que tenham sido rejeitados pela grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            With lvwContaCorrente.ListItems
                For lngCont = 1 To .Count
                    If Split(.Item(lngCont).Key, "|")(POS_NU_SEQU_OPER_ATIV) = objDomNode.selectSingleNode("Operacao").Text Then
                        For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                            .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                        Next
                        
                        .Item(lngCont).Text = "Horário Excedido"
                        .Item(lngCont).ToolTipText = "Horário limite p/envio da mensagem excedido"
                        .Item(lngCont).ForeColor = vbRed
                        .Item(lngCont).Tag = COD_ERRO_NEGOCIO_GRADE
                        
                        Exit For
                    End If
                Next
            End With
        Next
    End If

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0

End Sub

'Exibe de maneira diferenciada os itens que tenham sido rejeitados por processo de estorno
Private Sub flMarcarRejeitadosPorProcessoEstorno()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_ESTORNO & "']")
            With lvwContaCorrente.ListItems
                For lngCont = 1 To .Count
                    If Split(.Item(lngCont).Key, "|")(POS_NU_SEQU_OPER_ATIV) = objDomNode.selectSingleNode("Operacao").Text Then
                        For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                            .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                        Next
                        
                        .Item(lngCont).Text = "Operação Processo Estorno"
                        .Item(lngCont).ToolTipText = "Operação estornada, ou em processo de estorno"
                        .Item(lngCont).ForeColor = vbRed
                        .Item(lngCont).Tag = COD_ERRO_NEGOCIO_ESTORNO
                        
                        Exit For
                    End If
                Next
            End With
        Next
    End If

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorProcessoEstorno", 0

End Sub

'Monta o filtro XML com as condições de pesquisa
Private Function flMontarXmlFiltro() As String

Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim datDMenos1                              As Date
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltro, "", "Repeat_Filtros", "")
    
    'Filtro Empresa
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", cboEmpresa.ItemData(cboEmpresa.ListIndex))
        
    'Filtro Datas
    datDMenos1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, enumPaginacao.Anterior)
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(datDMenos1)))
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
     
    'Fltro Status (Cosidera apenas os registros já integrados com Conta Corrente)
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Status", "Status", enumStatusIntegracao.IntegradoCC)
     
    flMontarXmlFiltro = xmlFiltro.xml
     
    Set xmlFiltro = Nothing
            
Exit Function
ErrorHandler:
    Set xmlFiltro = Nothing
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "frmIntegracaoCCBatch", "flMontarXmlFiltro", 0
    
End Function

'Mostra o resultado do último processamento
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " integrados "
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    If cboEmpresa.ListIndex <> -1 Then
        Call fgCursor(True)
        Call flCarregarLista
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboEmpresa_Click"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("refresh"))
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    Call flInicializar
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load"
    
End Sub

Private Sub Form_Resize()
    
On Error Resume Next
    
    With Me
        .fraEmpresa.Top = 60
        .fraEmpresa.Left = 0
        .fraEmpresa.Width = .Width - 120
        
        .lvwContaCorrente.Top = .fraEmpresa.Top + .fraEmpresa.Height
        .lvwContaCorrente.Left = .fraEmpresa.Left
        .lvwContaCorrente.Height = .Height - .lvwContaCorrente.Top - 720
        .lvwContaCorrente.Width = .Width - 120
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub lvwContaCorrente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwContaCorrente, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwContaCorrente_DblClick", Me.Caption

End Sub

Private Sub lvwContaCorrente_DblClick()
    
On Error GoTo ErrorHandler

    If Not lvwContaCorrente.SelectedItem Is Nothing Then
        If lvwContaCorrente.SelectedItem.ForeColor = vbRed Then Exit Sub
        
        With frmHistLancamentoCC
            .lngCodigoEmpresa = fgObterCodigoCombo(lvwContaCorrente.SelectedItem)
            .vntSequenciaOperacao = Split(lvwContaCorrente.SelectedItem.Key, "|")(POS_NU_SEQU_OPER_ATIV)
            .lngTipoLancamentoITGR = Split(lvwContaCorrente.SelectedItem.Key, "|")(POS_TP_LANC_ITGR)
            .intSequenciaLancamento = Split(lvwContaCorrente.SelectedItem.Key, "|")(POS_NR_SEQU_LANC)
            .strNetOperacoes = lvwContaCorrente.SelectedItem.SubItems(COL_COND_NET_OPERACOES)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwContaCorrente_DblClick", Me.Caption

End Sub

Private Sub lvwContaCorrente_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
    
On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    
    Select Case Button.Key
        Case "Integrar"
            If flExisteLancamentoDMenos1Suspenso() Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Existe(m) Lançamento(s) com status  SUSPENSO com data de D-1."
                frmMural.IconeExibicao = IconInformation
                frmMural.Show vbModal
                GoTo ExitSub
            End If
        
            fgCursor True
            
            strResultadoOperacao = flGerenciar
        
            If strResultadoOperacao <> vbNullString Then
                Call flMostrarResultado(strResultadoOperacao)
                Call flCarregarLista
            End If
            
            Call flMarcarRejeitadosPorGradeHorario
            Call flMarcarRejeitadosPorProcessoEstorno
        
        Case "refresh"
            Call fgCursor(True)
            Call flCarregarLista
        
        Case "Sair"
            Unload Me
            
    End Select
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbFiltro_ButtonClick", Me.Caption

End Sub
'Calcula o Net da operações
Public Function flNetOperacoesBatch(ByVal strItemKey As String)
    
Dim strExpression                   As String
Dim vntValor                        As Variant
    
    vntValor = 0
    
    strExpression = flMontarCalculoNetOperacoesBatch(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlDomLeitura, strExpression))
    
    flNetOperacoesBatch = vntValor

End Function
'Monta uma expressão XPath para a somatória dos valores de operações
Public Function flMontarCalculoNetOperacoesBatch(ByVal strItemKey As String)
                
Dim strDebito                               As String
Dim strCredito                              As String
    
On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                   " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                   " and ../DT_OPER='" & Split(strItemKey, "|")(KEY_DATA_OPERACAO) & "' " & _
                                   " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                   " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                   " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                   " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                   " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                   " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                   " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' "
    
    strCredito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                    " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                    " and ../DT_OPER='" & Split(strItemKey, "|")(KEY_DATA_OPERACAO) & "' " & _
                                    " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                    " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                    " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                    " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                    " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                    " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                    " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' "
    
    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "
    
    flMontarCalculoNetOperacoesBatch = strDebito & strCredito
    
    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCalculoNetOperacoes", 0

End Function



