VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConciliacaoFinanceira 
   Caption         =   "Ferramentas - Conciliação e Liquidação Financeira"
   ClientHeight    =   8595
   ClientLeft      =   765
   ClientTop       =   1335
   ClientWidth     =   14100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   14100
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOptions 
      Caption         =   "Modalidade Liquidação"
      Height          =   570
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   2385
      Begin VB.OptionButton optModalidadeLiqu 
         Caption         =   "LBTR"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optModalidadeLiqu 
         Caption         =   "Bilateral"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Tipo Transferência"
      Height          =   570
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   720
      Width           =   3015
      Begin VB.OptionButton optTipoTransf 
         Caption         =   "Mercado"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   4
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton optTipoTransf 
         Caption         =   "Book Transfer"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   5610
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Natureza Movimento"
      Height          =   570
      Index           =   2
      Left            =   5880
      TabIndex        =   9
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "Pagamento"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "Recebimento"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   6
         Top             =   270
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   8265
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   582
      ButtonWidth     =   2328
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Concordar   "
            Key             =   "concordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Discordar    "
            Key             =   "discordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Rejeitar       "
            Key             =   "retorno"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pagamento "
            Key             =   "pagamento"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pagto. STR "
            Key             =   "pagamentostr"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pg. BACEN "
            Key             =   "pagamentobacen"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pg. Conting."
            Key             =   "pagamentocontingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Regularizar  "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Operações"
            Key             =   "MostrarOperacao"
            Object.ToolTipText     =   "Mostrar Operações"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Mensagens"
            Key             =   "MostrarMensagem"
            Object.ToolTipText     =   "Mostrar Mensagens"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair            "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":16D8
            Key             =   "agendar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceira.frx":19F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   12360
      Top             =   7170
      _extentx        =   2990
      _extenty        =   661
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   2445
      Left            =   60
      TabIndex        =   8
      Top             =   1380
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   4313
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cliente"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contraparte"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número Comando"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor Operação"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Valor Mensagem"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Diferença"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Conf. BT"
         Object.Width           =   1481
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Número Controle LTR"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Identificador Particip. Câmara"
         Object.Width           =   4233
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   5430
      Left            =   60
      TabIndex        =   13
      Top             =   3990
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   9578
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cliente"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID Ativo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Vencimento"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "D/C"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   2990
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "PU"
         Object.Width           =   2194
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Número Comando"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Data Liquidação"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Data Operação"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CNPJ/CPF Comitente"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.Image imgDummyH 
      Height          =   60
      Left            =   60
      MousePointer    =   7  'Size N S
      Top             =   3870
      Width           =   13950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmConciliacaoFinanceira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pela conciliação financeira de operações bruta e bilateral,
'' através da camada de controle de caso de uso MIU.
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlISPBRecebimento                  As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_CLIENTE                   As Integer = 0
Private Const COL_BCO_LIQU_CONTRAPARTE      As Integer = 1
Private Const COL_NUMERO_COMANDO            As Integer = 2
Private Const COL_VALOR_OPERACAO            As Integer = 3
Private Const COL_VALOR_MENSAGEM            As Integer = 4
Private Const COL_DIFERENCA                 As Integer = 5
Private Const COL_STATUS                    As Integer = 6
Private Const COL_CONFIRMACAO_BT            As Integer = 7
Private Const COL_NUMERO_CTRL_LTR           As Integer = 8
Private Const COL_ID_PART_CAMR              As Integer = 9

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const POS_CO_EMPR                   As Integer = 1
Private Const POS_NU_CTRL_LTR               As Integer = 2
Private Const POS_NU_COMD_OPER              As Integer = 3
Private Const POS_ID_PART_CAMR_VEIC_LEGA    As Integer = 4
Private Const POS_ID_PART_CAMR_CNTP         As Integer = 5
Private Const POS_CO_ISPB_BANC_LIQU_CNPT    As Integer = 6
Private Const POS_CO_ULTI_SITU_PROC         As Integer = 7
Private Const POS_CO_CNPJ_VEIC_LEGA         As Integer = 8

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const POS_NU_SEQU_OPER_ATIV         As Integer = 1

'Constantes de posicionamento de campos na propriedade Tag do item do ListView
Private Const POS_NU_CTRL_IF                As Integer = 1
Private Const POS_DH_REGT_MESG_SPB          As Integer = 2
Private Const POS_NU_SEQU_CNTR_REPE         As Integer = 3
Private Const POS_DH_ULTI_ATLZ              As Integer = 4
Private Const POS_CO_MESG_SPB               As Integer = 5
Private Const POS_CO_ISPB_BANC_LIQU_TAG     As Integer = 6

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_CLIENTE                As Integer = 0
Private Const COL_OP_ID_TITULO              As Integer = 1
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 2
Private Const COL_OP_CV                     As Integer = 3
Private Const COL_OP_VALOR                  As Integer = 4
Private Const COL_OP_QUANTIDADE             As Integer = 5
Private Const COL_OP_PU                     As Integer = 6
Private Const COL_OP_CODIGO                 As Integer = 7
Private Const COL_OP_NUMERO_COMANDO         As Integer = 8
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 9
Private Const COL_OP_DATA_OPERACAO          As Integer = 10
Private Const COL_OP_CNPJ_CPJ_COMITENTE     As Integer = 11

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmConciliacaoFinanceira"

'Constantes de erros de negócio específicos
Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3095

'Enums para os Option Buttons
Private Enum enumModalidadeLiquidacao
    Bruta = 0
    Bilateral = 1
End Enum

Private Enum enumTipoTransferencia
    BookTransfer = 0
    Mercado = 1
End Enum

Private lngPerfil                           As Long
Private blnDummyH                           As Boolean

Private intTipoConciliacao                  As enumTipoConciliacao
Private intAcaoConciliacao                  As enumAcaoConciliacao

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Variável criada para armazenar o nome do campo identificador de contraparte.
'Se CO_CNPT_CAMR vier preenchido na mensagem, assume CO_CNPT_CAMR, se não, assume CO_CNPJ_CNPT.
'Regra implementada em 07/11/2006 - Cassiano - Problema identificado em produção.
Private strCampoIdentContraparte            As String
Private blnBilateralBookTransfer            As Boolean
Private blnProcurarNetOperacoes             As Boolean

'Calcular a diferença dos valores das operações e mensagens SPB.
Private Sub flCalcularDiferencasListView()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

On Error GoTo ErrorHandler
    
    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VALOR_OPERACAO)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VALOR_MENSAGEM)))
            
            .SubItems(COL_DIFERENCA) = fgVlrXml_To_Interface(dblValorMensagem - dblValorOperacao)
            
            If dblValorMensagem - dblValorOperacao <> 0 Then
                .ListSubItems(COL_DIFERENCA).ForeColor = vbRed
            Else
                If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                    objListItem.Checked = True
                End If
            End If
        End With
    Next

Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListView", 0)

End Sub

'Controlar a chamada das funcionalidades que irão preencher as informações da Interface
Private Sub flCarregarLista()

Dim strDocFiltros                           As String
    
On Error GoTo ErrorHandler
    
    Call flLimparListas
    
    If Me.cboEmpresa.ListIndex = -1 Or Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    fgCursor True
    
    blnBilateralBookTransfer = False
    blnProcurarNetOperacoes = False
    
    strDocFiltros = flMontarXmlFiltro
    Call flCarregarMensagens(strDocFiltros)
    Call flCarregarNetOperacoes(strDocFiltros)
    
    If blnBilateralBookTransfer Then
        blnBilateralBookTransfer = False
        blnProcurarNetOperacoes = True
        
        strCampoIdentContraparte = "CO_CNPJ_CNPT"
        Call flCarregarNetOperacoes(strDocFiltros)
    
        blnProcurarNetOperacoes = False
    End If
    
    Call flCalcularDiferencasListView

    fgCursor
    
Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarLista", 0)

End Sub

'' Carrega o NET das mensagens SPB e preencher a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método A8MIU.clsMensagem.ObterNetMensagemConciliacao

Private Sub flCarregarMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
Dim strListItemTag                          As String
Dim strValorMensagem                        As String

Dim strEmpresa                              As String
Dim strNumeroCtrlLTR                        As String
Dim strNumeroComando                        As String
Dim strIdentPartCamaraVeicLega              As String
Dim strIdentPartCamaraContraparte           As String
Dim strISPBContraparte                      As String
Dim strCNPJVeiculoLegal                     As String
Dim intStatusConciliacao                    As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

    On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterNetMensagemConciliacao(pstrFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMensagem = Nothing
    strCampoIdentContraparte = "CO_CNPJ_CNPT"
    
    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_NetMensagem/*")
        
            If Mid$(objDomNode.selectSingleNode("NU_COMD_OPER").Text, 1, 10) <> "SFI-CETIP#" Then
            
                'Torna equivalentes os Status de Mensagens e Operações para integrar a chave de conciliação
                Select Case Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text)
                    Case enumStatusMensagem.ConcordanciaBackoffice, enumStatusMensagem.ConcordanciaBackofficeAutomatico
                        intStatusConciliacao = enumStatusOperacao.ConcordanciaBackoffice
                    
                    Case enumStatusMensagem.DiscordanciaBackoffice
                        intStatusConciliacao = enumStatusOperacao.DiscordanciaBackoffice
                    
                    Case enumStatusMensagem.PagamentoBackoffice, enumStatusMensagem.PagamentoBackofficeAutomatico
                        intStatusConciliacao = enumStatusOperacao.PagamentoBackoffice
                        
                    Case enumStatusMensagem.AConciliar
                        intStatusConciliacao = enumStatusOperacao.Registrada
                        
                End Select
                
                ' Montagem da chave do listview
                '=================================================================
                strEmpresa = objDomNode.selectSingleNode("CO_EMPR").Text
                
                If optModalidadeLiqu(enumModalidadeLiquidacao.Bruta).value Then
                    
                    If InStr(1, objDomNode.selectSingleNode("NU_COMD_OPER").Text, "SPR#") <> 0 Or _
                       InStr(1, objDomNode.selectSingleNode("NU_COMD_OPER").Text, "TERMO#") <> 0 Then
                        strNumeroCtrlLTR = vbNullString
                        strNumeroComando = Mid$(objDomNode.selectSingleNode("NU_COMD_OPER").Text, InStr(1, objDomNode.selectSingleNode("NU_COMD_OPER").Text, "#") + 1)
                    Else
                        strNumeroCtrlLTR = objDomNode.selectSingleNode("NU_CTRL_CAMR").Text
                        strNumeroComando = vbNullString
                    End If
                    
                    strIdentPartCamaraVeicLega = vbNullString
                    strIdentPartCamaraContraparte = vbNullString
                    strISPBContraparte = vbNullString
                    strCNPJVeiculoLegal = vbNullString
                    
                Else
    
                    strNumeroCtrlLTR = vbNullString
                    
                    If optTipoTransf(enumTipoTransferencia.Mercado).value Then
                        strCNPJVeiculoLegal = vbNullString
                        strIdentPartCamaraContraparte = vbNullString
                        strIdentPartCamaraVeicLega = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                        If optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                            strISPBContraparte = vbNullString
                        Else
                            strISPBContraparte = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
                        End If
                    Else
                        strCNPJVeiculoLegal = objDomNode.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text
                        strIdentPartCamaraVeicLega = vbNullString
                        strISPBContraparte = vbNullString
                        
                        'RATS 586/001
                        'Problema produção - Cassiano - 07/11/2006
                        If Val(objDomNode.selectSingleNode("CO_CNPT_CAMR").Text) <> 0 Then
                            strCampoIdentContraparte = "CO_CNPT_CAMR"
                            blnBilateralBookTransfer = True
                        End If
                        strIdentPartCamaraContraparte = objDomNode.selectSingleNode(strCampoIdentContraparte).Text
                    End If
                    
                End If
                
                strListItemKey = "|" & strEmpresa & _
                                 "|" & strNumeroCtrlLTR & _
                                 "|" & strNumeroComando & _
                                 "|" & strIdentPartCamaraVeicLega & _
                                 "|" & strIdentPartCamaraContraparte & _
                                 "|" & strISPBContraparte & _
                                 "|" & intStatusConciliacao & _
                                 "|" & strCNPJVeiculoLegal
                '=================================================================
                strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                                 "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                                 "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                                 "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_MESG_SPB").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
    
                If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                
                    With lvwMensagem.ListItems.Add(, strListItemKey)
                        
                        strValorMensagem = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                        
                        If Val(objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                            strValorMensagem = "-" & strValorMensagem
                        End If
                        
                        .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                        .SubItems(COL_VALOR_MENSAGEM) = strValorMensagem
                        .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                        .SubItems(COL_NUMERO_COMANDO) = Replace$(.SubItems(COL_NUMERO_COMANDO), "SPR#", vbNullString)
                        .SubItems(COL_NUMERO_COMANDO) = Replace$(.SubItems(COL_NUMERO_COMANDO), "TERMO#", vbNullString)
                        .SubItems(COL_NUMERO_CTRL_LTR) = objDomNode.selectSingleNode("NU_CTRL_CAMR").Text
                        .SubItems(COL_ID_PART_CAMR) = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                        
                        If optTipoTransf(enumTipoTransferencia.BookTransfer).value Then
                            .SubItems(COL_CONFIRMACAO_BT) = objDomNode.selectSingleNode("IN_CONF_MESG_LTR").Text
                        End If
                        
                        .Tag = strListItemTag
        
                    End With
                    
                Else
                    If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
                       optTipoTransf(enumTipoTransferencia.Mercado).value And _
                       optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                        
                        'Esta condição só pode acontecer na modalidade Bilateral / Mercado / Recebimento.
                        'Neste caso, um veículo legal pode ter mais de uma LTR0005R2, o que resultaria numa duplicação
                        'da chave do ListView.
                        strValorMensagem = objDomNode.selectSingleNode("VA_FINC").Text
                        
                        With lvwMensagem.ListItems(strListItemKey)
                            .SubItems(COL_VALOR_MENSAGEM) = _
                                    fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VALOR_MENSAGEM))) _
                                    + _
                                    fgVlrXml_To_Decimal(strValorMensagem)
                        
                            .SubItems(COL_VALOR_MENSAGEM) = fgVlrXml_To_Interface(fgVlr_To_Xml(.SubItems(COL_VALOR_MENSAGEM)))
                            .Tag = .Tag & strListItemTag
                        End With
                        
                    End If
                End If
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarMensagens", 0)

End Sub

'' Carrega o NET das operações e preencher a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método A8MIU.clsOperacao.ObterNetOperacaoConciliacaoFinanceira
Private Sub flCarregarNetOperacoes(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim xmlISPBAux                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
Dim dblValorOperacao                        As Double
Dim strNomeContraparte                      As String

Dim strEmpresa                              As String
Dim strNumeroCtrlLTR                        As String
Dim strNumeroComando                        As String
Dim strIdentPartCamaraVeicLega              As String
Dim strIdentPartCamaraContraparte           As String
Dim strISPBContraparte                      As String
Dim strCNPJVeiculoLegal                     As String
Dim intStatusConciliacao                    As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim objListItem                             As MSComctlLib.ListItem

    On Error GoTo ErrorHandler

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterNetOperacaoConciliacaoFinanceira(pstrFiltro, _
                                                                      strCampoIdentContraparte, _
                                                                      vntCodErro, _
                                                                      vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlISPBRecebimento = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlISPBRecebimento, "", "Repeat_ISPB", "")
        
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        For Each objDomNode In xmlRetLeitura.selectNodes("/Repeat_NetOperacao/*")
            
            If Not objDomNode.selectSingleNode("NU_COMD_OPER") Is Nothing Then
                strNumeroComando = objDomNode.selectSingleNode("NU_COMD_OPER").Text
            Else
                strNumeroComando = vbNullString
            End If
            
            If Mid$(strNumeroComando, 1, 10) <> "SFI-CETIP#" Then
            
                If optTipoTransf(enumTipoTransferencia.Mercado).value Then
                    If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
                       optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                        strNomeContraparte = vbNullString
                    Else
                        strNomeContraparte = objDomNode.selectSingleNode("NO_ISPB").Text
                    End If
                Else
                    strNomeContraparte = objDomNode.selectSingleNode("NO_CNPT").Text
                End If
                
                If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.RegistradaAutomatica Then
                    intStatusConciliacao = enumStatusOperacao.Registrada
                Else
                    intStatusConciliacao = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
                End If
                
                ' Montagem da chave do listview
                '=================================================================
                strEmpresa = objDomNode.selectSingleNode("CO_EMPR").Text
                
                If optModalidadeLiqu(enumModalidadeLiquidacao.Bruta).value Then
                    
                    If objDomNode.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
                       objDomNode.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
                        strNumeroCtrlLTR = vbNullString
                        strNumeroComando = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    Else
                        strNumeroCtrlLTR = objDomNode.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text
                        strNumeroComando = vbNullString
                    End If
                    
                    strIdentPartCamaraVeicLega = vbNullString
                    strIdentPartCamaraContraparte = vbNullString
                    strISPBContraparte = vbNullString
                    strCNPJVeiculoLegal = vbNullString
                    
                Else
    
                    strNumeroCtrlLTR = vbNullString
                    
                    If optTipoTransf(enumTipoTransferencia.Mercado).value Then
                        strCNPJVeiculoLegal = vbNullString
                        strIdentPartCamaraContraparte = vbNullString
                        strIdentPartCamaraVeicLega = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                        If optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                            strISPBContraparte = vbNullString
                        Else
                            strISPBContraparte = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
                        End If
                    Else
                        strCNPJVeiculoLegal = objDomNode.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text
                        strIdentPartCamaraVeicLega = vbNullString
                        strISPBContraparte = vbNullString
                    
                        'RATS 586/001
                        'Problema produção - Cassiano - 07/11/2006
                        strIdentPartCamaraContraparte = objDomNode.selectSingleNode(strCampoIdentContraparte).Text
                    End If
                    
                End If
                
                strListItemKey = "|" & strEmpresa & _
                                 "|" & strNumeroCtrlLTR & _
                                 "|" & strNumeroComando & _
                                 "|" & strIdentPartCamaraVeicLega & _
                                 "|" & strIdentPartCamaraContraparte & _
                                 "|" & strISPBContraparte & _
                                 "|" & intStatusConciliacao & _
                                 "|" & strCNPJVeiculoLegal
                '=================================================================
                        
                dblValorOperacao = fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                
                'Quando a Conciliação é Bruta, adiciona o sinal "-" ao valor, caso este seja negativo
                'Quando é Bilateral, como o Net é calculado no Select, o sinal já vem no valor caso este seja negativo
                If optModalidadeLiqu(enumModalidadeLiquidacao.Bruta).value Then
                    If Val(objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                        dblValorOperacao = -dblValorOperacao
                    End If
                
                    'Na tebela de Operações, o indicador de Débito e Crédito refere-se aos Títulos, e não ao
                    'valor financeiro, então o sinal do valor é invertido para anular esta disparidade.
                    'Na modalidade Bilateral, e select do Net das operações já está calculado corretamente.
                    dblValorOperacao = -dblValorOperacao
                End If
                
                If blnProcurarNetOperacoes Then
                    For Each objListItem In lvwMensagem.ListItems
                        If objListItem.SubItems(COL_VALOR_OPERACAO) = fgVlrXml_To_Interface(fgVlr_To_Xml(dblValorOperacao)) And _
                           Trim$(objListItem.SubItems(COL_VALOR_MENSAGEM)) <> vbNullString Then
                            GoTo VerificarProximo
                        End If
                    Next
                End If
                
                If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                    
                    'Só inclui a operação acordando com a Natureza do Movimento
                    If (optNaturezaMovimento(enumTipoDebitoCredito.Debito).value And dblValorOperacao < 0) Or _
                       (optNaturezaMovimento(enumTipoDebitoCredito.Credito).value And dblValorOperacao >= 0) Then
                    
                        If blnBilateralBookTransfer Then
                            Exit Sub
                        End If
                        
                        With lvwMensagem.ListItems.Add(, strListItemKey)
                            
                            If Not objDomNode.selectSingleNode("NO_VEIC_LEGA") Is Nothing Then
                                .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                            Else
                                .Text = vbNullString
                            End If
                            
                            .SubItems(COL_BCO_LIQU_CONTRAPARTE) = strNomeContraparte
                            .SubItems(COL_VALOR_OPERACAO) = fgVlrXml_To_Interface(fgVlr_To_Xml(dblValorOperacao))
                            .SubItems(COL_NUMERO_COMANDO) = strNumeroComando
                            .SubItems(COL_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                            
                            If Not objDomNode.selectSingleNode("NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
                                .SubItems(COL_NUMERO_CTRL_LTR) = objDomNode.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text
                            End If
                            
                            If Not objDomNode.selectSingleNode("CO_PARP_CAMR") Is Nothing Then
                                .SubItems(COL_ID_PART_CAMR) = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                            End If
                            
                            If PerfilAcesso = enumPerfilAcesso.AdmArea Then
                                Select Case intStatusConciliacao
                                    Case enumStatusOperacao.ConcordanciaBackoffice, _
                                         enumStatusOperacao.PagamentoBackoffice, _
                                         enumStatusOperacao.ConcordanciaBackofficeAutomatico, _
                                         enumStatusOperacao.PagamentoBackofficeAutomatico
                                         
                                        .ListSubItems(COL_STATUS).ForeColor = vbBlue
                                        
                                End Select
                            End If
                            
                        End With
                        
                        If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
                           optTipoTransf(enumTipoTransferencia.Mercado).value And _
                           optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                           
                            'Armazena os Códigos ISPB para a pesquisa das operações que compuseram o Net.
                            Set xmlISPBAux = CreateObject("MSXML2.DOMDocument.4.0")
                            
                            Call fgAppendNode(xmlISPBAux, "", "Grupo_ISPB", "")
                            
                            Call fgAppendNode(xmlISPBAux, "Grupo_ISPB", "ID_PART_CAMR_CETIP", objDomNode.selectSingleNode("CO_PARP_CAMR").Text)
                            Call fgAppendNode(xmlISPBAux, "Grupo_ISPB", "CO_ISPB_BANC_LIQU_CNPT", objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text)
            
                            Call fgAppendXML(xmlISPBRecebimento, "Repeat_ISPB", xmlISPBAux.xml)
                        
                            Set xmlISPBAux = Nothing
                        End If
    
                    End If
                
                Else
                
                    With lvwMensagem.ListItems(strListItemKey)
                        .SubItems(COL_BCO_LIQU_CONTRAPARTE) = strNomeContraparte
                        .SubItems(COL_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    
                        If PerfilAcesso = enumPerfilAcesso.AdmArea Then
                            Select Case intStatusConciliacao
                                Case enumStatusOperacao.ConcordanciaBackoffice, _
                                     enumStatusOperacao.PagamentoBackoffice, _
                                     enumStatusOperacao.ConcordanciaBackofficeAutomatico, _
                                     enumStatusOperacao.PagamentoBackofficeAutomatico
                                     
                                    .ListSubItems(COL_STATUS).ForeColor = vbBlue
                                    
                            End Select
                        End If
                
                        If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
                           optTipoTransf(enumTipoTransferencia.Mercado).value And _
                           optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                           
                            If dblValorOperacao >= 0 Then
                                
                                .SubItems(COL_VALOR_OPERACAO) = _
                                        fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VALOR_OPERACAO))) _
                                        + _
                                        dblValorOperacao
                            
                                .SubItems(COL_VALOR_OPERACAO) = fgVlrXml_To_Interface(fgVlr_To_Xml(.SubItems(COL_VALOR_OPERACAO)))
                                    
                                'Armazena os Códigos ISPB para a pesquisa das operações que compuseram o Net.
                                Set xmlISPBAux = CreateObject("MSXML2.DOMDocument.4.0")
                                
                                Call fgAppendNode(xmlISPBAux, "", "Grupo_ISPB", "")
                                
                                Call fgAppendNode(xmlISPBAux, "Grupo_ISPB", "ID_PART_CAMR_CETIP", objDomNode.selectSingleNode("CO_PARP_CAMR").Text)
                                Call fgAppendNode(xmlISPBAux, "Grupo_ISPB", "CO_ISPB_BANC_LIQU_CNPT", objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text)
                
                                Call fgAppendXML(xmlISPBRecebimento, "Repeat_ISPB", xmlISPBAux.xml)
                            
                                Set xmlISPBAux = Nothing
                                
                            End If
                        
                        Else
                            'Só inclui a operação acordando com a Natureza do Movimento
                            If (optNaturezaMovimento(enumTipoDebitoCredito.Debito).value And dblValorOperacao < 0) Or _
                               (optNaturezaMovimento(enumTipoDebitoCredito.Credito).value And dblValorOperacao >= 0) Then
                            
                                .SubItems(COL_VALOR_OPERACAO) = fgVlrXml_To_Interface(fgVlr_To_Xml(dblValorOperacao))
                                
                            End If
                        
                        End If
                    End With
    
                End If
            
            End If
            
VerificarProximo:
        Next
    End If
    
    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarNetOperacoes", 0)

End Sub

' Carrega as operações e preencher a interface com os mesmos,
' através da classe controladora de caso de uso MIU, método A8MIU.clsOperacao.ObterDetalheOperacao
Private Sub flCarregarOperacoesPorMensagem()

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
#End If

Dim strRetLeitura           As String
Dim xmlRetLeitura           As MSXML2.DOMDocument40
Dim strDocFiltros           As String
Dim xmlFiltro               As MSXML2.DOMDocument40
Dim objListItem             As ListItem

Dim objDomNode              As MSXML2.IXMLDOMNode
Dim strListItemKey          As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    lvwOperacao.ListItems.Clear

    Set objListItem = lvwMensagem.SelectedItem

    If objListItem Is Nothing Then
        Exit Sub
    End If

    If objListItem.SubItems(COL_VALOR_OPERACAO) = vbNullString Then
        Exit Sub
    End If

    strDocFiltros = flMontarXmlFiltro(objListItem)
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(strDocFiltros, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarOperacoesPorMensagem")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text

            With lvwOperacao.ListItems.Add(, strListItemKey)

                .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                
                .SubItems(COL_OP_CODIGO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
                .SubItems(COL_OP_CV) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_OP_ID_TITULO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
                .SubItems(COL_OP_QUANTIDADE) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("QT_ATIV_MERC").Text)
                .SubItems(COL_OP_PU) = fgVlrXml_To_InterfaceDecimais(objDomNode.selectSingleNode("PU_ATIV_MERC").Text, 8)
                .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                .SubItems(COL_OP_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                
                If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                End If
                
                If objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_OP_DATA_LIQUIDACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
                End If
                
                If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                End If

                If Not objDomNode.selectSingleNode("NR_CNPJ_CPF") Is Nothing Then
                    .SubItems(COL_OP_CNPJ_CPJ_COMITENTE) = objDomNode.selectSingleNode("NR_CNPJ_CPF").Text
                End If

            End With
        Next
    End If

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
    Set xmlRetLeitura = Nothing

Exit Sub
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarOperacoesPorMensagem", 0)

End Sub

'' Executar a conciliação efetuado através da camada controladora de casos de uso
'' MIU, método A8MIU.clsOperacaoMensagem.ConciliarCamaraLote
Private Function flConciliar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLRetorno                           As String
Dim strConciliacaoMsg                       As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    strConciliacaoMsg = flMontarXMLConciliacao
    
    If strConciliacaoMsg <> vbNullString Then
        fgCursor True
        
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.ConciliarCamaraLote(intTipoConciliacao, _
                                                                intAcaoConciliacao, _
                                                                strConciliacaoMsg, _
                                                                vntCodErro, _
                                                                vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        Set objOperacaoMensagem = Nothing
        fgCursor
    End If
    
    'Verifica se o retorno da operação possui erros
    If strXMLRetorno <> vbNullString Then
        '...se sim, carrega o XML de Erros
        Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetornoErro.loadXML(strXMLRetorno)
    Else
        '...se não, apenas destrói o objeto
        Set xmlRetornoErro = Nothing
    End If
    
    flConciliar = strXMLRetorno
    
Exit Function
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flConciliar", Me.Caption

End Function

'Configurar os botões da tela conforme o perfil do usuário liberando ou não utilização

Private Sub flConfigurarBotoesPorFuncionalidade()

On Error GoTo ErrorHandler
    
    With tlbComandos
    
        If (optTipoTransf(enumTipoTransferencia.BookTransfer).value) Or _
           (optTipoTransf(enumTipoTransferencia.Mercado).value And optNaturezaMovimento(enumTipoDebitoCredito.Credito).value) Then
            
            .Buttons("concordancia").Enabled = True
            .Buttons("discordancia").Enabled = IIf(optTipoTransf(enumTipoTransferencia.BookTransfer).value, True, False)
            .Buttons("retorno").Enabled = True
            .Buttons("pagamento").Enabled = False
            .Buttons("pagamentostr").Enabled = False
            .Buttons("pagamentobacen").Enabled = False
            .Buttons("pagamentocontingencia").Enabled = IIf(optTipoTransf(enumTipoTransferencia.BookTransfer).value And _
                                                            optNaturezaMovimento(enumTipoDebitoCredito.Debito).value, True, False)
            .Buttons("regularizacao").Enabled = .Buttons("pagamentocontingencia").Enabled
            
        Else
        
            .Buttons("concordancia").Enabled = True
            .Buttons("discordancia").Enabled = True
            .Buttons("retorno").Enabled = True
            .Buttons("pagamento").Enabled = True
            .Buttons("pagamentostr").Enabled = True
            .Buttons("pagamentobacen").Enabled = True
            .Buttons("pagamentocontingencia").Enabled = True
            .Buttons("regularizacao").Enabled = True
            
        End If
        
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorFuncionalidade", 0

End Sub

'Configurar os botões da tela conforme o perfil do usuário liberando ou não utilização

Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)
    
On Error GoTo ErrorHandler
    
    With tlbComandos
    
        .Buttons("concordancia").Visible = True
        .Buttons("discordancia").Visible = True
        .Buttons("retorno").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("pagamento").Visible = True
        .Buttons("pagamentostr").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("pagamentobacen").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("pagamentocontingencia").Visible = True
        .Buttons("regularizacao").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Refresh
  
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorPerfil", 0

End Sub

' Carrega as propriedades necessárias ao formulário, através da
' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Limpar os list view de mensagem e de operação
Private Sub flLimparListas()
    
    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear

End Sub

'Marcar as mensagens SPB que deveriam ser enviadas mas retornaram rejeitadas por grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            With lvwMensagem.ListItems
                For lngCont = 1 To .Count
                    If Split(.Item(lngCont).Tag, "|")(POS_NU_CTRL_IF) = objDomNode.selectSingleNode("NumeroControleIF").Text And _
                       Split(.Item(lngCont).Tag, "|")(POS_DH_REGT_MESG_SPB) = objDomNode.selectSingleNode("DTRegistroMensagemSPB").Text And _
                       Split(.Item(lngCont).Tag, "|")(POS_NU_SEQU_CNTR_REPE) = objDomNode.selectSingleNode("SequenciaRepeticao").Text Then
                        
                        For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                            .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                        Next
                        
                        .Item(lngCont).Text = "Horário Excedido"
                        .Item(lngCont).ToolTipText = "Horário limite p/envio da mensagem excedido"
                        .Item(lngCont).ForeColor = vbRed
                        
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

'Montar o xml para conciliação das informações apresentadas na interface

Private Function flMontarXMLConciliacao() As String

Dim objListItem                             As ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlConciliacaoMsg                       As MSXML2.DOMDocument40
Dim xmlItemConciliacao                      As MSXML2.DOMDocument40
Dim xmlDadosSTR                             As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer

Dim strValidaConciliacao                    As String
Dim strFinalidadeIF                         As String
Dim strMensagemConf                         As String
Dim strListItemTag                          As String

On Error GoTo ErrorHandler
    
    strValidaConciliacao = flVerificarItensConciliacao
    If strValidaConciliacao <> vbNullString Then
        If InStr(1, strValidaConciliacao, "?") <> 0 Then
            If MsgBox(strValidaConciliacao, vbYesNo + vbQuestion, Me.Caption) = vbNo Then Exit Function
        Else
            frmMural.Display = strValidaConciliacao
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            Exit Function
        End If
    End If
    
    Select Case intAcaoConciliacao
        Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
            strMensagemConf = "Concordância"
            
        Case enumAcaoConciliacao.AdmGeralEnviarDiscordancia
            strMensagemConf = "Discordância"
            
        Case enumAcaoConciliacao.AdmGeralPagamento, enumAcaoConciliacao.BOPagamento
            strMensagemConf = "Pagamento"
            
        Case enumAcaoConciliacao.AdmGeralPagamentoSTR
            strMensagemConf = "Pagamento STR"
            
        Case enumAcaoConciliacao.AdmGeralPagamentoBACEN
            strMensagemConf = "Pagamento BACEN"
            
        Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
            strMensagemConf = "Pagamento Contingência"
            
        Case enumAcaoConciliacao.AdmGeralRegularizar
            strMensagemConf = "Regularização"
            
    End Select
    
    If strMensagemConf <> vbNullString Then
        If MsgBox("Confirma " & strMensagemConf & " do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    
    If intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoSTR Then
        strFinalidadeIF = flObterDominioFinalidadeMsgSTR
    End If
    
    Set xmlDadosSTR = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlConciliacaoMsg = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlConciliacaoMsg, "", "Repeat_Conciliacao", "")
    
    For Each objListItem In Me.lvwMensagem.ListItems
        With objListItem
            If .Checked Then
                
                'Entrada Manual de dados complementares para o Pagamento STR
                If intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoSTR Then
                    With frmComplementoMensagensConciliacao
                        Set .objSelectedItem = objListItem
                        .strComboFinalidade = strFinalidadeIF
                        .Show vbModal
                        Call xmlDadosSTR.loadXML(.xmlComplemento.xml)
                        Set .xmlComplemento = Nothing
                        
                        If xmlDadosSTR.xml = vbNullString Then
                            With frmMural
                                .Display = "Informações complementares para o Pagamento STR não foram registradas. A seleção deste item será desfeita."
                                .IconeExibicao = IconCritical
                                .Show vbModal
                                
                                objListItem.Checked = False
                                GoTo ProximoItem
                            End With
                        End If
                    End With
                End If
            
ProcessarNovaMsg:
                
                strListItemTag = "|" & Split(.Tag, "|")(POS_NU_CTRL_IF) & _
                                 "|" & Split(.Tag, "|")(POS_DH_REGT_MESG_SPB) & _
                                 "|" & Split(.Tag, "|")(POS_NU_SEQU_CNTR_REPE) & _
                                 "|" & Split(.Tag, "|")(POS_DH_ULTI_ATLZ) & _
                                 "|" & Split(.Tag, "|")(POS_CO_MESG_SPB) & _
                                 "|" & Split(.Tag, "|")(POS_CO_ISPB_BANC_LIQU_TAG)

                Set xmlItemConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call fgAppendNode(xmlItemConciliacao, "", "Grupo_Mensagem", "")
                Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                      "NU_CTRL_IF", _
                                                      Split(.Tag, "|")(POS_NU_CTRL_IF))
                Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                      "DH_REGT_MESG_SPB", _
                                                      Split(.Tag, "|")(POS_DH_REGT_MESG_SPB))
                Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                      "NU_SEQU_CNTR_REPE", _
                                                      Split(.Tag, "|")(POS_NU_SEQU_CNTR_REPE))
                Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                      "DH_ULTI_ATLZ", _
                                                      Split(.Tag, "|")(POS_DH_ULTI_ATLZ))
                Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                      "CO_MESG_SPB", _
                                                      Split(.Tag, "|")(POS_CO_MESG_SPB))
                
                intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)
                
                Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                      "IgnoraGradeHorario", _
                                                      intIgnoraGradeHorario)
            
                If xmlDadosSTR.xml <> vbNullString Then
                    For Each objDomNode In xmlDadosSTR.selectNodes("Repeat_Conciliacao/Grupo_Mensagem/*")
                        Call fgAppendXML(xmlItemConciliacao, "Grupo_Mensagem", objDomNode.xml)
                    Next
                End If
                
                For Each objDomNode In xmlISPBRecebimento.selectNodes("Repeat_ISPB/Grupo_ISPB[ID_PART_CAMR_CETIP='" & Split(.Key, "|")(POS_ID_PART_CAMR_VEIC_LEGA) & "']")
                    Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", _
                                                          "CO_ISPB_BANC_LIQU_CNPT", _
                                                          objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text)
                                                          
                Next
                
                .Tag = Replace$(.Tag, strListItemTag, vbNullString)
                If .Tag = vbNullString Then
                    Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", "FIM_MSG", "TRUE")
                Else
                    Call fgAppendNode(xmlItemConciliacao, "Grupo_Mensagem", "FIM_MSG", "FALSE")
                    Call fgAppendXML(xmlConciliacaoMsg, "Repeat_Conciliacao", xmlItemConciliacao.xml)
                    Set xmlItemConciliacao = Nothing
                    
                    GoTo ProcessarNovaMsg
                End If
                                          
                Call fgAppendXML(xmlConciliacaoMsg, "Repeat_Conciliacao", xmlItemConciliacao.xml)
                Set xmlItemConciliacao = Nothing
                
            End If
        End With
        
ProximoItem:
    Next
    
    If xmlConciliacaoMsg.selectNodes("Repeat_Conciliacao/*").length = 0 Then
        flMontarXMLConciliacao = vbNullString
    Else
        flMontarXMLConciliacao = xmlConciliacaoMsg.xml
    End If
    
    Set xmlConciliacaoMsg = Nothing
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLConciliacao", 0

End Function

'Montar o xml de filtro para pesquisa pelo perfil do usuário

Private Function flMontarXmlFiltro(Optional objListItem As ListItem = Nothing) As String
    
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strSelecaoFiltroOper                    As String
Dim strSelecaoFiltroMsg                     As String
Dim lngCont                                 As Long
    
Dim strEmpresa                              As String
Dim strNumeroCtrlLTR                        As String
Dim strNumeroComando                        As String
Dim strIdentPartCamaraVeicLega              As String
Dim strIdentPartCamaraContraparte           As String
Dim strISPBContraparte                      As String
Dim intStatusConciliacao                    As Integer
Dim strCNPJVeiculoLegal                     As String
    
    On Error GoTo ErrorHandler
    
    If PerfilAcesso = BackOffice Then
        strSelecaoFiltroOper = enumStatusOperacao.Registrada & ";" & _
                               enumStatusOperacao.RegistradaAutomatica
        
        strSelecaoFiltroMsg = enumStatusMensagem.AConciliar
    
    ElseIf PerfilAcesso = AdmArea Then
        strSelecaoFiltroOper = enumStatusOperacao.ConcordanciaBackoffice & ";" & _
                               enumStatusOperacao.DiscordanciaBackoffice & ";" & _
                               enumStatusOperacao.PagamentoBackoffice & ";" & _
                               enumStatusOperacao.Registrada & ";" & _
                               enumStatusOperacao.RegistradaAutomatica & ";" & _
                               enumStatusOperacao.ConcordanciaBackofficeAutomatico & ";" & _
                               enumStatusOperacao.PagamentoBackofficeAutomatico
                           
        strSelecaoFiltroMsg = enumStatusMensagem.ConcordanciaBackoffice & ";" & _
                              enumStatusMensagem.DiscordanciaBackoffice & ";" & _
                              enumStatusMensagem.AConciliar & ";" & _
                              enumStatusMensagem.PagamentoBackoffice & ";" & _
                              enumStatusMensagem.ConcordanciaBackofficeAutomatico & ";" & _
                              enumStatusMensagem.PagamentoBackofficeAutomatico
    
    End If
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    
    If objListItem Is Nothing Then
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_StatusOperacao", "")
        For lngCont = LBound(Split(strSelecaoFiltroOper, ";")) To UBound(Split(strSelecaoFiltroOper, ";"))
            Call fgAppendNode(xmlFiltros, "Grupo_StatusOperacao", "Status", Split(strSelecaoFiltroOper, ";")(lngCont))
        Next
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_StatusMensagem", "")
        For lngCont = LBound(Split(strSelecaoFiltroMsg, ";")) To UBound(Split(strSelecaoFiltroMsg, ";"))
            Call fgAppendNode(xmlFiltros, "Grupo_StatusMensagem", "Status", Split(strSelecaoFiltroMsg, ";")(lngCont))
        Next
        
    Else
        
        With objListItem
            strEmpresa = Split(.Key, "|")(POS_CO_EMPR)
            strNumeroCtrlLTR = Split(.Key, "|")(POS_NU_CTRL_LTR)
            strNumeroComando = Split(.Key, "|")(POS_NU_COMD_OPER)
            strIdentPartCamaraVeicLega = Split(.Key, "|")(POS_ID_PART_CAMR_VEIC_LEGA)
            strIdentPartCamaraContraparte = Split(.Key, "|")(POS_ID_PART_CAMR_CNTP)
            strISPBContraparte = Split(.Key, "|")(POS_CO_ISPB_BANC_LIQU_CNPT)
            intStatusConciliacao = Split(.Key, "|")(POS_CO_ULTI_SITU_PROC)
            strCNPJVeiculoLegal = Split(.Key, "|")(POS_CO_CNPJ_VEIC_LEGA)
        End With
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", strEmpresa)
    
        If strNumeroCtrlLTR <> vbNullString Then
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_NumeroControleLTR", "")
            Call fgAppendNode(xmlFiltros, "Grupo_NumeroControleLTR", "NumeroControleLTR", strNumeroCtrlLTR)
        End If

        If strNumeroComando <> vbNullString Then
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_NumeroComando", "")
            Call fgAppendNode(xmlFiltros, "Grupo_NumeroComando", "NumeroComando", strNumeroComando)
        End If

        If strIdentPartCamaraVeicLega <> vbNullString Then
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_IdentificadorPartCamara", "")
            Call fgAppendNode(xmlFiltros, "Grupo_IdentificadorPartCamara", "IdentificadorPartCamara", strIdentPartCamaraVeicLega)
        End If

        If Val(strIdentPartCamaraContraparte) <> 0 Then
            'Problema produção - Cassiano - 07/11/2006
            If Len(strIdentPartCamaraContraparte) > 8 Then
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CNPJContraparte", "")
                Call fgAppendNode(xmlFiltros, "Grupo_CNPJContraparte", "Grupo_CNPJContraparte", strIdentPartCamaraContraparte)
            Else
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_IdentificadorPartCamaraContraparte", "")
                Call fgAppendNode(xmlFiltros, "Grupo_IdentificadorPartCamaraContraparte", "IdentificadorPartCamaraContraparte", strIdentPartCamaraContraparte)
            End If
        End If

        If strCNPJVeiculoLegal <> vbNullString Then
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CNPJ", "")
            Call fgAppendNode(xmlFiltros, "Grupo_CNPJ", "CNPJ", strCNPJVeiculoLegal)
        End If

        If strISPBContraparte <> vbNullString Then
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ISPBContraparte", "")
            Call fgAppendNode(xmlFiltros, "Grupo_ISPBContraparte", "ISPBContraparte", strISPBContraparte)
        Else
            If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
               optTipoTransf(enumTipoTransferencia.Mercado).value And _
               optNaturezaMovimento(enumTipoDebitoCredito.Credito).value Then
                       
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ISPBContraparte", "")
                
                'No caso da Modalidade Bilateral / Mercado / Recebimento, define filtro com os
                'Códigos ISPB armazenados no Array.
                For Each objDomNode In xmlISPBRecebimento.selectNodes("Repeat_ISPB/Grupo_ISPB[ID_PART_CAMR_CETIP='" & strIdentPartCamaraVeicLega & "']")
                    If Trim$(objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text) <> vbNullString Then
                        Call fgAppendNode(xmlFiltros, "Grupo_ISPBContraparte", "ISPBContraparte", objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text)
                    End If
                Next
                    
            End If
        End If

        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", intStatusConciliacao)
        
        If intStatusConciliacao = enumStatusOperacao.Registrada Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica)
        End If
        
    End If
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoContraparte", "")
    
    If optModalidadeLiqu(enumModalidadeLiquidacao.Bruta) And optTipoTransf(enumTipoTransferencia.BookTransfer) Then
        Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bruta)
        Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Interno)
    
    ElseIf optModalidadeLiqu(enumModalidadeLiquidacao.Bruta) And optTipoTransf(enumTipoTransferencia.Mercado) Then
        Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bruta)
        Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo)
    
    ElseIf optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral) And optTipoTransf(enumTipoTransferencia.BookTransfer) Then
        Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bilateral)
        Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Interno)
    
    ElseIf optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral) And optTipoTransf(enumTipoTransferencia.Mercado) Then
        Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bilateral)
        Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo)
    
        If optNaturezaMovimento(enumTipoDebitoCredito.Debito) Then
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ISPBContraparte", "")
            Call fgAppendNode(xmlFiltros, "Grupo_ISPBContraparte", "ISPBContraparte", "NotNull")
        End If
    End If
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_NaturezaMovimento", "")
    If optNaturezaMovimento(enumTipoDebitoCredito.Debito) Then
        Call fgAppendNode(xmlFiltros, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Debito)
    Else
        Call fgAppendNode(xmlFiltros, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Credito)
    End If
    
    ' Somente 'D'efinitiva.
    ' Prévia deixou de existir.
    ' Cassiano - 14/09/2004
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoInformacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoInformacao", "TipoInformacao", "D")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CETIP)
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_AcaoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_AcaoMensagem", "AcaoMensagem", "CONDIÇÃO (TP_ACAO_MESG_SPB_EXEC NOT IN (" & enumTipoAcao.LTR0001ComISPBJaExistente & ") OR TP_ACAO_MESG_SPB_EXEC IS NULL)")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CanalOperacaoInterna", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CanalOperacaoInterna", "CanalOperacaoInterna", "CONDIÇÃO (CO_CNAL_OPER_INTE = 'N' OR CO_CNAL_OPER_INTE IS NULL)")
    '>>> -------------------------------------------------------------------------------------------
    
    flMontarXmlFiltro = xmlFiltros.xml
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltro", 0

End Function

'Montar o resultado do processamento

Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " conciliados "
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'' Obter o dominio da finalidade de mensagens STR,
'' através da classe controladora de caso de uso MIU, método A8MIU.clsOperacaoMensagem.ObterDominioFinalidadeMsgSTR

Private Function flObterDominioFinalidadeMsgSTR() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strComboFinalidade                      As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
    
    fgCursor True
    Call xmlLeitura.loadXML(objOperacaoMensagem.ObterDominioFinalidadeMsgSTR(vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    fgCursor
    
    For Each objDomNode In xmlLeitura.selectNodes("Repeat_DominioAtributo/*")
        strComboFinalidade = strComboFinalidade & _
                             objDomNode.selectSingleNode("CO_DOMI").Text & _
                             " - " & _
                             objDomNode.selectSingleNode("DE_DOMI").Text & _
                             vbTab
    Next
    
    flObterDominioFinalidadeMsgSTR = strComboFinalidade
    
    Set xmlLeitura = Nothing
    Set objOperacaoMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set xmlLeitura = Nothing
    Set objOperacaoMensagem = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flObterDominioFinalidadeMsgSTR", 0

End Function

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString
Private Function flVerificarItensConciliacao() As String

Dim objListItem                             As ListItem
Dim dblValorConsist                         As Double
Dim intStatus                               As Integer

    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flVerificarItensConciliacao = "Selecione pelo menos um item da lista, antes de prosseguir com a operação desejada."
        Exit Function
    End If
    
    If PerfilAcesso = enumPerfilAcesso.BackOffice Then
        For Each objListItem In Me.lvwMensagem.ListItems
            If objListItem.Checked Then
                dblValorConsist = fgVlrXml_To_Decimal(fgVlr_To_Xml(objListItem.SubItems(COL_DIFERENCA)))
                
                If intAcaoConciliacao = enumAcaoConciliacao.BOConcordar Then
                    If dblValorConsist <> 0 Then
                        flVerificarItensConciliacao = "Um ou mais itens estão com divergência de valores. Solicitação de Concordância não permitida."
                        Exit Function
                    End If
                
                ElseIf intAcaoConciliacao = enumAcaoConciliacao.BOPagamento Then
                    If dblValorConsist <> 0 Then
                        flVerificarItensConciliacao = "Um ou mais itens estão com divergência de valores. Solicitação de Pagamento não permitida."
                        Exit Function
                    End If
                
                ElseIf intAcaoConciliacao = enumAcaoConciliacao.BODiscordar Then
                    If dblValorConsist = 0 Then
                        flVerificarItensConciliacao = "Um ou mais itens batidos. Solicitação de Discordância não permitida."
                        Exit Function
                    End If
            
                End If
                
            End If
        Next
    
    Else
        For Each objListItem In Me.lvwMensagem.ListItems
            If objListItem.Checked Then
                dblValorConsist = fgVlrXml_To_Decimal(fgVlr_To_Xml(objListItem.SubItems(COL_DIFERENCA)))
                
                If intAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Then
                    If dblValorConsist <> 0 Then
                        flVerificarItensConciliacao = "Um ou mais itens estão com divergência de valores. Solicitação de Regularização não permitida."
                        Exit Function
                    End If
            
                ElseIf intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Then
                    If dblValorConsist = 0 Then
                        flVerificarItensConciliacao = "Um ou mais itens batidos. Solicitação de Pagamento em Contingência não permitida."
                        Exit Function
                    End If
            
                End If
                
            End If
        Next
    
        For Each objListItem In Me.lvwMensagem.ListItems
            If objListItem.Checked Then
                If objListItem.SubItems(COL_VALOR_OPERACAO) <> vbNullString Then
                    intStatus = Val(Split(objListItem.Key, "|")(POS_CO_ULTI_SITU_PROC))
                Else
                    intStatus = 0
                End If
                
                Select Case intAcaoConciliacao
                    Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                    
                        If intStatus <> enumStatusOperacao.ConcordanciaBackoffice And _
                           intStatus <> enumStatusOperacao.PagamentoBackoffice And _
                           intStatus <> enumStatusOperacao.ConcordanciaBackofficeAutomatico And _
                           intStatus <> enumStatusOperacao.PagamentoBackofficeAutomatico Then
                            flVerificarItensConciliacao = "Só é possível 'Concordar' itens de mensagem nos seguintes status:" & vbNewLine & vbNewLine & _
                                                          "- Concordância Backoffice" & vbNewLine & _
                                                          "- Pagamento Backoffice" & vbNewLine & _
                                                          "- Concordância BO Automática" & vbNewLine & _
                                                          "- Pagamento BO Automático"
                            Exit Function
                        End If
                
                    Case enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                    
                        If intStatus <> enumStatusOperacao.DiscordanciaBackoffice And _
                           intStatus <> 0 Then
                            flVerificarItensConciliacao = "Só é possível 'Discordar' itens de mensagem nos seguintes status:" & vbNewLine & vbNewLine & _
                                                          "- Discordância Backoffice"
                            Exit Function
                        End If
                
                    Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                    
                        If (intStatus <> enumStatusOperacao.DiscordanciaBackoffice And _
                            intStatus <> enumStatusOperacao.Registrada And _
                            intStatus <> enumStatusOperacao.RegistradaAutomatica And _
                            intStatus <> 0) Then
                            flVerificarItensConciliacao = "Só é possível 'Pagar em Contingência' itens de mensagem nos seguintes status:" & vbNewLine & vbNewLine & _
                                                          "- Discordância Backoffice" & vbNewLine & _
                                                          "- Registrada" & vbNewLine & _
                                                          "- Registrada Automático"
                            Exit Function
                        End If
                
                    Case enumAcaoConciliacao.AdmGeralPagamento, _
                         enumAcaoConciliacao.AdmGeralPagamentoBACEN, _
                         enumAcaoConciliacao.AdmGeralPagamentoSTR
                    
                        If intStatus <> enumStatusOperacao.ConcordanciaBackoffice And _
                           intStatus <> enumStatusOperacao.PagamentoBackoffice And _
                           intStatus <> enumStatusOperacao.ConcordanciaBackofficeAutomatico And _
                           intStatus <> enumStatusOperacao.PagamentoBackofficeAutomatico Then
                            flVerificarItensConciliacao = "Só é possível 'Pagar' itens de mensagem nos seguintes status:" & vbNewLine & vbNewLine & _
                                                          "- Concordância Backoffice" & vbNewLine & _
                                                          "- Pagamento Backoffice" & vbNewLine & _
                                                          "- Concordância BO Automática" & vbNewLine & _
                                                          "- Pagamento BO Automático"
                            Exit Function
                        End If
                        
                        If intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoBACEN Or _
                           intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoSTR Then
                            If Split(objListItem.Tag, "|")(POS_CO_ISPB_BANC_LIQU_TAG) = vbNullString Then
                                flVerificarItensConciliacao = "Pagamento através da LTR0003 ou STR não permitido, pois a ISPB IF Creditada não está informada na mensagem recebida da câmara em um ou mais itens."
                                Exit Function
                            End If
                        End If
                
                    Case enumAcaoConciliacao.AdmGeralRegularizar
                    
                        If intStatus <> enumStatusOperacao.ConcordanciaBackoffice And _
                           intStatus <> enumStatusOperacao.PagamentoBackoffice And _
                           intStatus <> enumStatusOperacao.ConcordanciaBackofficeAutomatico And _
                           intStatus <> enumStatusOperacao.PagamentoBackofficeAutomatico Then
                            flVerificarItensConciliacao = "Só é possível 'Regularizar' itens de mensagem nos seguintes status:" & vbNewLine & vbNewLine & _
                                                          "- Concordância Backoffice" & vbNewLine & _
                                                          "- Pagamento Backoffice" & vbNewLine & _
                                                          "- Concordância BO Automática" & vbNewLine & _
                                                          "- Pagamento BO Automático"
                            Exit Function
                        End If
                End Select
            End If
        Next
    End If

End Function

Public Sub RedimensionarForm()
    
    Call Form_Resize

End Sub

Property Get PerfilAcesso() As enumPerfilAcesso
        
    PerfilAcesso = lngPerfil
    
End Property

Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)

    lngPerfil = pPerfil
    
    Select Case pPerfil
        Case enumPerfilAcesso.BackOffice
            Me.Caption = "Conciliação - Liquidação Financeira Bruta / Bilateral"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "Liberação - Liquidação Financeira Bruta / Bilateral"
    End Select
    
    Call flConfigurarBotoesPorPerfil(PerfilAcesso)
    Call flConfigurarBotoesPorFuncionalidade
    Call flLimparListas
    
    If cboEmpresa.ListIndex <> -1 Or cboEmpresa.Text <> vbNullString Then
        Call flCarregarLista
    End If
    
End Property

Private Sub cboEmpresa_Click()
    
On Error GoTo ErrorHandler

    If cboEmpresa.Text <> vbNullString Then
        Call flCarregarLista
    End If
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboEmpresa_Click", Me.Caption

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(lvwMensagem, Retorno)
    End Select
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        fgCursor True
        Call tlbComandos_ButtonClick(tlbComandos.Buttons("refresh"))
        fgCursor
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
    
    fgCursor True
    Call flInicializar
    fgCursor
    
    Set xmlISPBRecebimento = CreateObject("MSXML2.DOMDocument.4.0")
                
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If

        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width

        .lvwMensagem.Top = .fraOptions(0).Top + .fraOptions(0).Height + 120
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwMensagem.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
        .lvwOperacao.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlRetornoErro = Nothing
    Set xmlISPBRecebimento = Nothing
    Set frmConciliacaoFinanceira = Nothing

End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnDummyH = True

End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not blnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If

    Me.imgDummyH.Top = y + imgDummyH.Top

    On Error Resume Next

    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnDummyH = False

End Sub

Private Sub lvwMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwMensagem, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ColumnClick", Me.Caption

End Sub

Private Sub lvwMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
On Error GoTo ErrorHandler
    
    If Item.Checked Then
        If Item.SubItems(COL_VALOR_MENSAGEM) = vbNullString Then
            frmMural.Display = "Seleção do item para conciliação não permitida. Valor de mensagem não encontrado."
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            Item.Checked = False
        End If
    End If
        
    Item.Selected = True
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemCheck", Me.Caption

End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    fgCursor True
    Call flCarregarOperacoesPorMensagem
    fgCursor

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemClick", Me.Caption

End Sub

Private Sub lvwMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_MouseDown", Me.Caption

End Sub

Private Sub lvwOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwOperacao, ColumnHeader.Index)
    lngIndexClassifListOper = ColumnHeader.Index
    
Exit Sub
    
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_ColumnClick", Me.Caption

End Sub

Private Sub lvwOperacao_DblClick()
    
On Error GoTo ErrorHandler

    If Not lvwOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Split(lvwOperacao.SelectedItem.Key, "|")(POS_NU_SEQU_OPER_ATIV)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

End Sub

Private Sub optModalidadeLiqu_Click(Index As Integer)
    
On Error GoTo ErrorHandler
    
    If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
       optTipoTransf(enumTipoTransferencia.Mercado).value Then
        optNaturezaMovimento(enumTipoDebitoCredito.Debito).Enabled = False
        optNaturezaMovimento(enumTipoDebitoCredito.Credito).value = True
    Else
        optNaturezaMovimento(enumTipoDebitoCredito.Debito).Enabled = True
    End If
    
    Call flCarregarLista
    Call flConfigurarBotoesPorFuncionalidade
        
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optModalidadeLiqu_Click", Me.Caption

End Sub

Private Sub optNaturezaMovimento_Click(Index As Integer)
    
On Error GoTo ErrorHandler
    
    Call flCarregarLista
    Call flConfigurarBotoesPorFuncionalidade
        
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optNaturezaMovimento_Click", Me.Caption

End Sub

Private Sub optTipoTransf_Click(Index As Integer)
    
On Error GoTo ErrorHandler
    
    If optModalidadeLiqu(enumModalidadeLiquidacao.Bilateral).value And _
       optTipoTransf(enumTipoTransferencia.Mercado).value Then
        optNaturezaMovimento(enumTipoDebitoCredito.Debito).Enabled = False
        optNaturezaMovimento(enumTipoDebitoCredito.Credito).value = True
    Else
        optNaturezaMovimento(enumTipoDebitoCredito.Debito).Enabled = True
    End If
    
    Call flCarregarLista
    Call flConfigurarBotoesPorFuncionalidade
        
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optTipoTransf_Click", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
Dim strErro                                 As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    intTipoConciliacao = IIf(optModalidadeLiqu(enumModalidadeLiquidacao.Bruta).value, enumTipoConciliacao.Bruta, enumTipoConciliacao.Bilateral)
    intAcaoConciliacao = 0
    
    Select Case Button.Key
        Case "refresh"
            Call flCarregarLista
            
        Case "concordancia"
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoConciliacao = enumAcaoConciliacao.BOConcordar
            Else
                intAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia
            End If
            
        Case "discordancia"
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoConciliacao = enumAcaoConciliacao.BODiscordar
            Else
                intAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia
            End If
            
        Case "retorno"
            intAcaoConciliacao = enumAcaoConciliacao.AdmGeralRejeitar
            
        Case "pagamento"
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoConciliacao = enumAcaoConciliacao.BOPagamento
            Else
                intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento
            End If
            
        Case "pagamentostr"
            intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoSTR
            
        Case "pagamentobacen"
            intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoBACEN
            
        Case "pagamentocontingencia"
            intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia
            
        Case "regularizacao"
            intAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar
            
        Case gstrSair
            Unload Me
            
    End Select
    
    If intAcaoConciliacao <> 0 Then
        strResultadoOperacao = flConciliar
        
        If strResultadoOperacao <> vbNullString Then
            Call flMostrarResultado(strResultadoOperacao)
            Call flCarregarLista
        End If
    
        Call flMarcarRejeitadosPorGradeHorario
    End If
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
