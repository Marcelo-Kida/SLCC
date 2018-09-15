VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiberacaoOperacaoMensagem 
   Caption         =   "Liberação Operação e Mensagem SPB"
   ClientHeight    =   8565
   ClientLeft      =   240
   ClientTop       =   1335
   ClientWidth     =   14235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   4725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   8334
      View            =   3
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
      NumItems        =   32
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selecionar Operação"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Número Operação Câmbio 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Código Associação Câmbio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Tipo Movto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Tipo Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Código Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Veículo Legal(Parte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Contra-Parte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Título"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Data Vencto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Local Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Text            =   "Agendamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Cta. Cedente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Cta. Cessionário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Canal Venda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "CNPJ/CPF Comitente"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Código do Usuário"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Nome Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Moeda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Valor ME"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Taxa Cambial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Código Veículo Legal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Tipo Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Tipo Contrato"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8235
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   582
      ButtonWidth     =   2752
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar Filtro"
            Key             =   "AplicarFiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
            Key             =   "DefinirFiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Liberar        "
            Key             =   "Liberar"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rejeitar       "
            Key             =   "Rejeitar"
            Object.ToolTipText     =   "Rejeitar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Liberar Conting."
            Key             =   "LibConting"
            Object.ToolTipText     =   "LibConting"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Operações  "
            Key             =   "MostrarOperacao"
            Object.ToolTipText     =   "Mostrar Operações"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mensagens "
            Key             =   "MostrarMensagem"
            Object.ToolTipText     =   "Mostrar Mensagens"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair             "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   9720
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
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiberacaoOperacaoMensagem.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMensagem 
      Height          =   4725
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   8334
      View            =   3
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selecionar Mensagem"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Número Registro Operação Câmbio 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Código Associação Câmbio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Código Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "NumeroConciliacao"
         Object.Width           =   0
      EndProperty
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   13200
      Top             =   4620
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   4800
      Width           =   14040
   End
End
Attribute VB_Name = "frmLiberacaoOperacaoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:34:32
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Liberação manual de operações
'' e/ou mensagens que não possuem liberação automática) à camada controladora de
'' caso de uso A8MIU.

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private intControleMenuPopUp                As enumTipoConfirmacao

Private fblnDummyH                          As Boolean
Private blnPrimeiraConsulta                 As Boolean
Private strFiltroXML                        As String

Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3023

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_OPERACAO               As Integer = 1
Private Const COL_OP_TP_ACAO_OPER_ATIV_EXEC As Integer = 2
Private Const COL_OP_DATA_OPERACAO          As Integer = 3
Private Const COL_OP_NUM_COMANDO            As Integer = 4
Private Const COL_OP_NUM_OPER_CAMB2         As Integer = 5
Private Const COL_OP_CD_ASSO_CAMB           As Integer = 6
Private Const COL_OP_VALOR                  As Integer = 7
Private Const COL_OP_TIPO_MOVIMENTO         As Integer = 8
Private Const COL_OP_SITUACAO               As Integer = 9
Private Const COL_OP_TIPO_OPER              As Integer = 10
Private Const COL_OP_CO_MESG_SPB_REGT_OPER  As Integer = 11
Private Const COL_OP_VEICULO_LEGAL_PARTE    As Integer = 12
Private Const COL_OP_CONTRAPARTE            As Integer = 13
Private Const COL_OP_TITULO                 As Integer = 14
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 15
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 16
Private Const COL_OP_LOCAL_LIQUIDACAO       As Integer = 17
Private Const COL_OP_EMPRESA                As Integer = 18
Private Const COL_OP_HORARIO_ENVIO_MSG      As Integer = 19     '<-- Agendamento
Private Const COL_OP_CONTA_CEDENTE          As Integer = 20
Private Const COL_OP_CONTA_CESSIONARIO      As Integer = 21
Private Const COL_OP_CANAL_VENDA            As Integer = 22
Private Const COL_OP_CNPJ_CPF_COMITENTE     As Integer = 23
Private Const COL_OP_CODIGO_USUARIO         As Integer = 24
Private Const COL_OP_NO_CLIE                As Integer = 25
Private Const COL_OP_CD_MOED_ISO            As Integer = 26
Private Const COL_OP_VA_MOED_ESTR           As Integer = 27
Private Const COL_OP_NR_PERC_TAXA_CAMB      As Integer = 28
Private Const COL_OP_CO_VEIC_LEGA           As Integer = 29
Private Const COL_TO_CO_MESG_SPB_REGT_OPER  As Integer = 30
Private Const COL_OP_TP_OPER_CAMB           As Integer = 31

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_DATA_MENSAGEM         As Integer = 1
Private Const COL_MSG_NUMERO_COMANDO        As Integer = 2
Private Const COL_MSG_NUM_OPER_CAMB2        As Integer = 3
Private Const COL_MSG_CD_ASSO_CAMB          As Integer = 4
Private Const COL_MSG_CODIGO_MENSAGEM       As Integer = 5
Private Const COL_MSG_SITUACAO              As Integer = 6
Private Const COL_MSG_EMPRESA               As Integer = 7
Private Const COL_MSG_VALOR                 As Integer = 8
Private Const COL_MSG_CNCL                  As Integer = 9


Private Const strFuncionalidade             As String = "frmLiberacaoOperacaoMensagem"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private lngPerfil                           As Long

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Configura a apresentação das janelas do formulário

Private Sub flArranjarJanelasExibicao(ByVal pstrJanelas As String)
    
On Error GoTo ErrorHandler

    Select Case pstrJanelas
           Case ""
                imgDummyH.Visible = False
                lstOperacao.Visible = False
                lstMensagem.Visible = False
            
           Case "1"
                imgDummyH.Visible = False
                lstOperacao.Visible = True
                lstMensagem.Visible = False
            
           Case "2"
                imgDummyH.Visible = False
                lstOperacao.Visible = False
                lstMensagem.Visible = True
                
           Case "12"
                imgDummyH.Visible = True
                lstOperacao.Visible = True
                lstMensagem.Visible = True
                
    End Select
    
    Call Form_Resize

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flArranjarJanelasExibicao", 0
    
End Sub

'' Encaminhar a solicitação (Leitura de todas as mensagens, para o preenchimento
'' do listview) à camada controladora de caso de uso (componente / classe / metodo
'' ) :
''
'' A8MIU.clsMensagem.ObterDetalheMensagem
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flCarregarListaMensagem(ByVal strFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strSelecaoFiltro                        As String
Dim lngCont                                 As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Call flLimparLista(lstMensagem)

    Select Case lngPerfil
        Case enumPerfilAcesso.Nenhum
            strSelecaoFiltro = enumStatusMensagem.Concordancia & ";" & _
                               enumStatusMensagem.ConcordanciaAutomatica
        Case enumPerfilAcesso.AdmArea
            strSelecaoFiltro = enumStatusMensagem.PendenteLibAlcadaAdmArea
        Case enumPerfilAcesso.AdmGeral
            strSelecaoFiltro = enumStatusMensagem.PendenteLibAlcadaAdmGeral
    End Select

    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlDomFiltros.loadXML(strFiltro) Then
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    End If
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", ""
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    '>>> -------------------------------------------------------------------------------------------
    'Grupo BMC0015 Contingencia
'    If lngPerfil = enumPerfilAcesso.Nenhum Then
'        fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_BMC0015Contingencia", ""
'        fgAppendNode xmlDomFiltros, "Grupo_BMC0015Contingencia", "Status", enumStatusMensagem.AConciliar
'    End If
    
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
        
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            If objDomNode.selectSingleNode("CO_MESG_SPB").Text <> "BMC0112" Then
        
                With lstMensagem.ListItems.Add(, _
                        "k" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                              objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                        
                    .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                           objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                           objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & "|" & _
                           objDomNode.selectSingleNode("NU_PRTC_MESG_LG").Text & "|" & _
                           objDomNode.selectSingleNode("VA_FINC").Text & "|" & _
                           objDomNode.selectSingleNode("TP_OPER").Text
        
                    If objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text <> gstrDataVazia Then
                        .SubItems(COL_MSG_DATA_MENSAGEM) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                    End If
                    
                    .SubItems(COL_MSG_CODIGO_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                    .SubItems(COL_MSG_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    If objDomNode.selectSingleNode("NR_OPER_CAMB_2").Text <> 0 Then
                        .SubItems(COL_MSG_NUM_OPER_CAMB2) = objDomNode.selectSingleNode("NR_OPER_CAMB_2").Text
                    Else
                        .SubItems(COL_MSG_NUM_OPER_CAMB2) = ""
                    End If
                    .SubItems(COL_MSG_CD_ASSO_CAMB) = objDomNode.selectSingleNode("CD_ASSO_CAMB").Text
                    .SubItems(COL_MSG_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_MSG_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    .SubItems(COL_MSG_CNCL) = objDomNode.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text
                                        
                    If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                        'Obtem a descrição da Empresa via QUERY XML
                        .SubItems(COL_MSG_EMPRESA) = _
                            xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                                objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                    End If
                    
                    
                End With
                
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lstMensagem, lngIndexClassifListMesg, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaMensagem", 0

End Sub

'' Encaminhar a solicitação (Leitura de todas as operações, para o preenchimento
'' do listview) à camada controladora de caso de uso (componente / classe / metodo
'' ) :
''
'' A8MIU.clsOperacao.ObterDetalheOperacao
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flCarregarListaOperacao(ByVal strFiltro As String)

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
Dim strSelecaoAcao                          As String
Dim lngCont                                 As Long
Dim dblTotalOperacao                        As Double

Dim strOperacoesBMA                         As String
Dim strOperacoesBMC                         As String
Dim strOperacoesBMD                         As String
Dim strOperacoesCETIP                       As String
Dim strOperacoesSTR                         As String
Dim strOperacoesPAG                         As String
Dim strOperacoesContaCorrente               As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

'KIDA - CCR
Dim strOperacoesCCR                         As String
Dim strOperacoesCAM                         As String

On Error GoTo ErrorHandler

    Call flLimparLista(lstOperacao)
    
    strSelecaoFiltro = enumStatusOperacao.Concordancia & ";" & _
                       enumStatusOperacao.ConcordanciaAutomatica & ";" & _
                       enumStatusOperacao.ConcordanciaBalcao & ";" & _
                       enumStatusOperacao.ConcordanciaBalcaoAutomatica & ";" & _
                       enumStatusOperacao.ConcordanciaReativacao & ";" & _
                       enumStatusOperacao.ConcordanciaReativacaoAutomatica & ";" & _
                       enumStatusOperacao.Conciliada & ";" & _
                       enumStatusOperacao.ConciliadaAutomatica & ";" & _
                       enumStatusOperacao.ConcordanciaBackoffice

    strSelecaoAcao = enumTipoAcao.CancelamentoSolicitado & ";" & _
                     enumTipoAcao.EstornoSolicitado

    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltros.loadXML(strFiltro) Then
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    End If

    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", ""
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoAcao", ""

    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    
    Call fgAppendAttribute(xmlDomFiltros, "Grupo_TipoAcao", "Operador", "OR")
    For lngCont = LBound(Split(strSelecaoAcao, ";")) To UBound(Split(strSelecaoAcao, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoAcao", _
                                         "TipoAcao", Split(strSelecaoAcao, ";")(lngCont))
    Next
    '>>> -------------------------------------------------------------------------------------------
    
    'Coloca filtro por local de liquidacao e tipo de operacao
    
    strOperacoesBMA = enumTipoOperacaoLQS.TransferenciaBMA & ", " & _
                      enumTipoOperacaoLQS.LiquidacaoOperacaoTermoBMA & ", " & _
                      enumTipoOperacaoLQS.LiquidacaoEventosJurosBMA & ", " & _
                      enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA & ", " & _
                      enumTipoOperacaoLQS.OperacaoTermoInternaBMA & ", " & _
                      enumTipoOperacaoLQS.EspecDefinitivaIntermediacao & ", " & _
                      enumTipoOperacaoLQS.EspecDefinitivaCobertura & ", " & _
                      enumTipoOperacaoLQS.EspecTermoIntermediacao & ", " & _
                      enumTipoOperacaoLQS.EspecTermoCobertura & ", " & _
                      enumTipoOperacaoLQS.DepositoBMA & ", " & _
                      enumTipoOperacaoLQS.RetiradaBMA & ", " & _
                      enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA & ", " & _
                      enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA & ", " & _
                      enumTipoOperacaoLQS.CancelamentoEspecificacaoBMA & ", " & _
                      enumTipoOperacaoLQS.EspecCompromissadaIntermediacao & ", " & _
                      enumTipoOperacaoLQS.EspecCompromissadaCobertura & ", " & _
                      enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA & ", "
    
    strOperacoesBMA = strOperacoesBMA & _
                      enumTipoOperacaoLQS.LiquidacaoCompromissadaEspecificaVolta & ", " & _
                      enumTipoOperacaoLQS.LiquidacaoEventosJurosTituloCompro & ", " & _
                      enumTipoOperacaoLQS.OperacaoCompromissadaInternaBMA & ", " & _
                      enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA & ", " & _
                      enumTipoOperacaoLQS.NETEntradaManualMultilateralBMA

    strOperacoesCETIP = enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiro & ", " & _
                        enumTipoOperacaoLQS.RetiradaCustodia & ", " & _
                        enumTipoOperacaoLQS.VincDesvInstrumentoFinanceiro & ", " & _
                        enumTipoOperacaoLQS.TransferenciaCustodia & ", " & _
                        enumTipoOperacaoLQS.ResgateFundoInvestimento & ", " & _
                        enumTipoOperacaoLQS.MovimExercicioDireitosDebentures & ", " & _
                        enumTipoOperacaoLQS.ConversaoPermutaValorMobiliario & ", " & _
                        enumTipoOperacaoLQS.EspecifQtdeCotasFundoInvestimento & ", " & _
                        enumTipoOperacaoLQS.OperacaoDefinitivaCETIP & ", " & _
                        enumTipoOperacaoLQS.OperacaoCompromissadaCETIP & ", " & _
                        enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP & ", " & _
                        enumTipoOperacaoLQS.OperacaoRetornoCompromissadaCETIP & ", " & _
                        enumTipoOperacaoLQS.OperacaoRetencaoIR & ", " & _
                        enumTipoOperacaoLQS.RegistroContratoSWAP & ", " & _
                        enumTipoOperacaoLQS.RegDadosComplemContratoSWAP & ", " & _
                        enumTipoOperacaoLQS.RegistroContratoTermo & ", " & _
                        enumTipoOperacaoLQS.ExercicioOpcaoContratoSWAP & ", " & _
                        enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP & ", " & _
                        enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO & ", " & _
                        enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo & ", " & _
                        enumTipoOperacaoLQS.OperacaoCessaoContratoDerivativo & ", " & _
                        enumTipoOperacaoLQS.OperAnuenciaCessaoContratoDerivativo & ", " & _
                        enumTipoOperacaoLQS.OperIntermediacaoContratoDerivativo & ", "
                        
    strOperacoesCETIP = strOperacoesCETIP & _
                        enumTipoOperacaoLQS.EventosJurosSWAP & ", " & _
                        enumTipoOperacaoLQS.EventosJurosTERMO & ", " & _
                        enumTipoOperacaoLQS.EventosCETIP & ", " & _
                        enumTipoOperacaoLQS.DespesasCETIP & ", " & _
                        enumTipoOperacaoLQS.DepositoBMA & ", " & _
                        enumTipoOperacaoLQS.RetiradaBMA & ", " & _
                        enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA & ", " & _
                        enumTipoOperacaoLQS.TransferenciaBMA & ", " & _
                        enumTipoOperacaoLQS.AplicacaoFundoInvestimentoCETIP & ", " & _
                        enumTipoOperacaoLQS.DepositoFundoInvestimentoCETIP & ", " & _
                        enumTipoOperacaoLQS.NETEntradaManualMultilateralCETIP & ", " & _
                        enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP & ", "
                        
    strOperacoesCETIP = strOperacoesCETIP & _
                        enumTipoOperacaoLQS.RegistroContratoSWAPSemOpcaoBarreira & ", " & _
                        enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira & ", " & _
                        enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP_CETIP21 & ", " & _
                        enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo_CETIP21 & ", " & _
                        enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiroCTP4001

    strOperacoesBMC = enumTipoOperacaoLQS.RegistroOperacoesBMC & ", " & _
                      enumTipoOperacaoLQS.LiquidacaoMultilateralBMC & ", " & _
                      enumTipoOperacaoLQS.DespesasBMC & ", " & _
                      enumTipoOperacaoLQS.TransferenciasBMCDeposito & ", " & _
                      enumTipoOperacaoLQS.TransferenciasBMCRetirada & ", " & _
                      enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC & ", " & _
                      enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao & ", " & _
                      enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica & ", " & _
                      enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaContrArbitParceiroExteriorPaisPropriaIF & ", " & _
                      enumTipoOperacaoLQS.InformaOperacaoArbitragemParceiroPais & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais & ", " & _
                      enumTipoOperacaoLQS.CAMInformaContratacaoInterbancarioViaLeilao & ", " & _
                      enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria & ", " & _
                      enumTipoOperacaoLQS.ConsultaContratosCambioMercadoInterbancario

    strOperacoesBMD = enumTipoOperacaoLQS.DepositoBMA & ", " & _
                      enumTipoOperacaoLQS.RetiradaBMA & ", " & _
                      enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA & ", " & _
                      enumTipoOperacaoLQS.TransferenciaBMA & ", " & _
                      enumTipoOperacaoLQS.NETEntradaManualMultilateralBMD
    
    strOperacoesSTR = enumTipoOperacaoLQS.EnvioPagDespesasBoleto & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasTributos & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib & ", " & _
                      enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega & ", " & _
                      enumTipoOperacaoLQS.CAMInformaContratacaoInterbancarioViaLeilao & ", " & _
                      enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria

    strOperacoesContaCorrente = enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais & ", " & _
                                enumTipoOperacaoLQS.LancamentoCCCashFlow & ", " & _
                                enumTipoOperacaoLQS.LancamentoCCCashFlowStrikeFixo & ", " & _
                                enumTipoOperacaoLQS.LancamentoCCSwapJuros

    'KIDA - CCR
    strOperacoesCCR = enumTipoOperacaoLQS.ConsultaOperacaoCCR & ", " & _
                      enumTipoOperacaoLQS.ConsultaLimitesImportacaoCCR & ", " & _
                      enumTipoOperacaoLQS.EmissaoOperacaoCCR & ", " & _
                      enumTipoOperacaoLQS.NegociacaoOperacaoCCR & ", " & _
                      enumTipoOperacaoLQS.DevolucaoRecolhimentoEstornoReembolsoCCR
    
    strOperacoesCAM = enumTipoOperacaoLQS.ContratacaoMercadoPrimario & ", " & _
                      enumTipoOperacaoLQS.EdicaoContratacaoMercadoPrimario & ", " & _
                      enumTipoOperacaoLQS.ConfirmacaoEdicaoContratacaoMercadoPrimario & ", " & _
                      enumTipoOperacaoLQS.AlteracaoContrato & ", " & _
                      enumTipoOperacaoLQS.EdicaoAlteracaoContrato & ", " & _
                      enumTipoOperacaoLQS.ConfirmacaoEdicaoAlteracaoContrato & ", " & _
                      enumTipoOperacaoLQS.LiquidacaoMercadoPrimario & ", " & _
                      enumTipoOperacaoLQS.BaixaValorLiquidar & ", " & _
                      enumTipoOperacaoLQS.RestabelecimentoBaixa & ", " & _
                      enumTipoOperacaoLQS.CancelamentoValorLiquidar & ", " & _
                      enumTipoOperacaoLQS.EdicaoCancelamentoValorLiquidar & ", " & _
                      enumTipoOperacaoLQS.ConfirmacaoEdicaoCancelamentoValorLiquidar & ", " & _
                      enumTipoOperacaoLQS.VinculacaoContratos & ", " & _
                      enumTipoOperacaoLQS.AnulacaoEvento & ", " & _
                      enumTipoOperacaoLQS.CorretoraRequisitaClausulasEspecificas & ", " & _
                      enumTipoOperacaoLQS.IFInformaClausulasEspecificas & ", " & _
                      enumTipoOperacaoLQS.ManutencaoCadastroAgenciaCentralizadoraCambio & ", " & _
                      enumTipoOperacaoLQS.CredenciamentoDescredenciamentoDispostoRMCCI & ", " & _
                      enumTipoOperacaoLQS.IncorporacaoContratos & ", " & _
                      enumTipoOperacaoLQS.AceiteRejeicaoIncorporacaoContratos & ", " & _
                      enumTipoOperacaoLQS.ConsultaContratosEmSer & ", " & _
                      enumTipoOperacaoLQS.ConsultaEventosUmDia & ", " & _
                      enumTipoOperacaoLQS.ConsultaDetalhamentoContratoInterbancario & ", " & _
                      enumTipoOperacaoLQS.ConsultaEventosContratoMercadoPrimario & ", " & _
                      enumTipoOperacaoLQS.ConsultaEventosContratoIntermediadoMercadoPrimario & ", "
                      
    strOperacoesCAM = strOperacoesCAM & _
                      enumTipoOperacaoLQS.ConsultaHistoricoIncorporacoes & ", " & _
                      enumTipoOperacaoLQS.ConsultaContratosIncorporacao & ", " & _
                      enumTipoOperacaoLQS.ConsultaCadeiaIncorporacoesContrato & ", " & _
                      enumTipoOperacaoLQS.ConsultaPosicaoCambioMoeda & ", " & _
                      enumTipoOperacaoLQS.AtualizaçãoInclusãoInstrucoesPagamento & ", " & _
                      enumTipoOperacaoLQS.ConsultaInstrucoesPagamento & ", " & _
                      enumTipoOperacaoLQS.InformaTIRemContrapartidaaRagadorouRecebedorPaís & ", " & _
                      enumTipoOperacaoLQS.InformaTIRemContrapartidaOutraCDE & ", " & _
                      enumTipoOperacaoLQS.InformaTIRemContrapartidaOperacaoCambialPropria & ", " & _
                      enumTipoOperacaoLQS.RequisitaInclusaoemCadastroCDE & ", " & _
                      enumTipoOperacaoLQS.RequisitaAlteracaoCadastroCDE & ", " & _
                      enumTipoOperacaoLQS.RequisitaExclusaoCadastroCDE & ", " & _
                      enumTipoOperacaoLQS.InformaAnulacaoRegistroTIR & ", " & _
                      enumTipoOperacaoLQS.ConsultaCDE & ", " & _
                      enumTipoOperacaoLQS.ConsultaTIRUmDia & ", " & _
                      enumTipoOperacaoLQS.ConsultaDetalhamentoTIR
                      
                      
    strOperacoesPAG = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega & ", " & _
                      enumTipoOperacaoLQS.CAMInformaContratacaoInterbancarioViaLeilao & ", " & _
                      enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria

    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacaoTipoOperacao", "")
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro1", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro1", "Local", enumLocalLiquidacao.SELIC)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro1", "Tipos", "")
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro2", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro2", "Local", enumLocalLiquidacao.BMA)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro2", "Tipos", strOperacoesBMA)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro3", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro3", "Local", enumLocalLiquidacao.CETIP)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro3", "Tipos", strOperacoesCETIP)

    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro4", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro4", "Local", enumLocalLiquidacao.BMD)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro4", "Tipos", strOperacoesBMD)

    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro5", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro5", "Local", enumLocalLiquidacao.BMC)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro5", "Tipos", strOperacoesBMC)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro6", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro6", "Local", enumLocalLiquidacao.SSTR)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro6", "Tipos", strOperacoesSTR)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro7", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro7", "Local", enumLocalLiquidacao.CONTA_CORRENTE)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro7", "Tipos", strOperacoesContaCorrente)
        
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro8", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro8", "Local", enumLocalLiquidacao.CCR)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro8", "Tipos", strOperacoesCCR)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro9", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro9", "Local", enumLocalLiquidacao.CAM)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro9", "Tipos", strOperacoesCAM)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro10", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro10", "Local", enumLocalLiquidacao.PAG)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro10", "Tipos", strOperacoesPAG)
    
    
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
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaOperacao")
        End If
    
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            
            '=====================================================================================
            'Correção RATS 780:
            'Operações CETIP com Status Concordância Backoffice não podem ser adicionadas à lista,
            'pois possibilitam dupla liberação de mensagens.
            'Cas - 16/07/2008
            
            If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.ConcordanciaBackoffice Or _
               Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> enumLocalLiquidacao.CETIP Then
            '=====================================================================================
            
                With lstOperacao.ListItems.Add(, _
                        "k" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                        
                    'Guarda na propriedade TAG a situação da operação | data da úlmtima atualização
                    .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                           objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                           objDomNode.selectSingleNode("NU_COMD_ACAO_EXEC").Text & "|" & objDomNode.selectSingleNode("TP_OPER").Text
            
                    .SubItems(COL_OP_OPERACAO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            
                    If Not objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = vbNullString And fgVlrXml_To_Decimal(objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) <> 0 Then
                        .SubItems(COL_OP_TP_ACAO_OPER_ATIV_EXEC) = objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text & " - " & fgDescricaoTipoAcao(CLng("0" & objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text))
                    End If
                    
                    If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                    End If
                     
                    .SubItems(COL_OP_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_OP_NUM_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    If objDomNode.selectSingleNode("NR_OPER_CAMB_2").Text <> 0 Then
                        .SubItems(COL_OP_NUM_OPER_CAMB2) = objDomNode.selectSingleNode("NR_OPER_CAMB_2").Text
                    Else
                        .SubItems(COL_OP_NUM_OPER_CAMB2) = ""
                    End If
                    .SubItems(COL_OP_CD_ASSO_CAMB) = objDomNode.selectSingleNode("CD_ASSO_CAMB").Text
                    .SubItems(COL_OP_TIPO_OPER) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                    .SubItems(COL_OP_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                    .SubItems(COL_OP_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_OP_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                    .SubItems(COL_OP_TITULO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                    .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                    'fixo SEL1023 pq BMA só pode cancelar ela - Pedido Mauricio 03/06/2004 - Carlos
                    If CLng(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Then
                        If CLng(objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) = enumTipoAcao.CancelamentoSolicitado Then
                            .SubItems(COL_OP_CO_MESG_SPB_REGT_OPER) = "SEL1023"
                        Else
                            .SubItems(COL_OP_CO_MESG_SPB_REGT_OPER) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                        End If
                    Else
                        .SubItems(COL_OP_CO_MESG_SPB_REGT_OPER) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                    End If
    
                    If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                    End If
                    
                    If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text), "HH:MM")
                    End If
                    
                    If objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_LIQUIDACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
                    End If
                    
                    If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                        
                        If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                               objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU") Is Nothing Then
                            
                            'Obtem a descrição do Local de Liquidação via QUERY XML
                            .SubItems(COL_OP_LOCAL_LIQUIDACAO) = _
                                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                    objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU").Text
                    
                        Else
                            
                            vntCodErro = 5
                            vntMensagemErro = "Usuário não possui acesso ao Local de Liquidação " & _
                                              objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "."
                            GoTo ErrorHandler
                            
                        End If
                    
                    End If
                    
                    If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                        'Obtem a descrição da Empresa via QUERY XML
                        .SubItems(COL_OP_EMPRESA) = _
                            xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                                objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                    End If
                    
                    .SubItems(COL_OP_CONTA_CEDENTE) = objDomNode.selectSingleNode("CO_CNTA_CEDT").Text
                    .SubItems(COL_OP_CONTA_CESSIONARIO) = objDomNode.selectSingleNode("CO_CNTA_CESS").Text
                    
                    'KIDA - SGC
                    .SubItems(COL_OP_CANAL_VENDA) = fgDescricaoCanalVenda(objDomNode.selectSingleNode("TP_CNAL_VEND").Text)
                    
                    If Val("0" & objDomNode.selectSingleNode("NR_CNPJ_CPF").Text) <> 0 Then
                        .SubItems(COL_OP_CNPJ_CPF_COMITENTE) = fgFormataCnpj(objDomNode.selectSingleNode("NR_CNPJ_CPF").Text)
                    End If
                    
                    .SubItems(COL_OP_CODIGO_USUARIO) = objDomNode.selectSingleNode("CO_USUA_CADR_OPER").Text
                
                     'ESTÁ IMPLEMENTAÇÃO ESTÁ EM STAND-BY, AGUARDANDO PRIORIZAÇÃO PARA SER IMPLANTADA
'                    'campos incluídos por solicitação dos usuário do Comex, devido projeto Sisbacen
'                    .SubItems(COL_OP_NO_CLIE) = objDomNode.selectSingleNode("NO_CLIE").Text
'                    .SubItems(COL_OP_CD_MOED_ISO) = objDomNode.selectSingleNode("CD_MOED_ISO").Text
'                    .SubItems(COL_OP_VA_MOED_ESTR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_MOED_ESTR").Text)
'                    .SubItems(COL_OP_NR_PERC_TAXA_CAMB) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("NR_PERC_TAXA_CAMB").Text)
'                    .SubItems(COL_OP_CO_VEIC_LEGA) = objDomNode.selectSingleNode("CO_VEIC_LEGA").Text
'                    .SubItems(COL_TO_CO_MESG_SPB_REGT_OPER) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
'                    If UCase(Trim(objDomNode.selectSingleNode("TP_OPER_CAMB").Text)) = "C" Then
'                        .SubItems(COL_OP_TP_OPER_CAMB) = "Compra"
'                    ElseIf UCase(Trim(objDomNode.selectSingleNode("TP_OPER_CAMB").Text)) = "V" Then
'                        .SubItems(COL_OP_TP_OPER_CAMB) = "Venda"
'                    End If
                
                End With
                
                'Acumula Total de Operações
                dblTotalOperacao = dblTotalOperacao + fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
        
            End If
            
        Next
        
        Call fgClassificarListview(Me.lstOperacao, lngIndexClassifListOper, True)
    
        With lstOperacao.ListItems.Add(, "kTotal")
            .SubItems(COL_OP_NUM_COMANDO) = "Total"
            .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(dblTotalOperacao)
            .Checked = True
            .ListSubItems(COL_OP_NUM_COMANDO).Bold = True
            .ListSubItems(COL_OP_VALOR).Bold = True
        End With
    End If
    
    Set xmlDomLeitura = Nothing

    Exit Sub

ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaOperacao", 0
    
End Sub

'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão) à camada controladora de caso de uso (componente / classe / metodo ) :
''
''
'' A8MIU.clsMiu.ObterMapaNavegacao
''
'' O método retornará uma String XML para a camada de interface.
''
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
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmLiberacaoOperacaoMensagem", "flInicializar")
    End If
    
    Set objMIU = Nothing

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'' Encaminhar a solicitação (Liberação de mensagens e operações) à camada
'' controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsOperacaoMensagem.Liberar
''
'' O método retornará uma String XML para a camada de interface.
''
Private Function flLiberar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim xmlDomLoteOperacaoMensagem              As MSXML2.DOMDocument40
Dim strXMLRetorno                           As String
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long
Dim intIgnoraGradeHorario                   As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Const POS_STATUS                            As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1
Const POS_NU_COMD_ACAO_EXEC                 As Integer = 2

Const POS_NUMERO_CONTROLE_IF                As Integer = 0
Const POS_DATA_REGISTRO_MENSAGEM_SPB        As Integer = 1
Const POS_NU_SEQU_OPER_ATIV_MENSAGEM        As Integer = 2
Const POS_NU_PRTC_MESG_LG                   As Integer = 3
Const POS_VA_FINC                           As Integer = 4
Const QTD_MAX_CONFIRMACAO                   As Integer = 500

Dim lngQuebras                              As Long
Dim lngContQuebras                          As Long
Dim lngContIni                              As Long
Dim lngContFim                              As Long


On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ---------------------------------------------------------------------------
    
    strXMLRetorno = vbNullString
    
    If lstOperacao.ListItems.Count > 0 Then
    
        If lstOperacao.ListItems.Count > QTD_MAX_CONFIRMACAO Then
            lngQuebras = Abs(lstOperacao.ListItems.Count / QTD_MAX_CONFIRMACAO)
            
            lngContIni = 1
            lngContFim = QTD_MAX_CONFIRMACAO
        Else
            lngContIni = 1
            lngContFim = lstOperacao.ListItems.Count
        End If
        
        'Captura o filtro cumulativo OPERAÇÃO
        For lngContQuebras = 0 To lngQuebras
            
            lngItensChecked = 0
            Set xmlDomLoteOperacaoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlDomLoteOperacaoMensagem, "", "Repeat_Filtros", "")
            
            If lngContFim > lstOperacao.ListItems.Count Then lngContFim = lstOperacao.ListItems.Count
            
            'Captura o filtro cumulativo OPERAÇÃO
            With lstOperacao.ListItems
                For lngCont = lngContIni To lngContFim
                    If .Item(lngCont).Checked And .Item(lngCont).Key <> "kTotal" Then
                        lngItensChecked = lngItensChecked + 1
                        
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Repeat_Filtros", "Grupo_Lote", "")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.operacao, "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Operacao", Mid(.Item(lngCont).Key, 2), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "TipoAcao", fgObterCodigoCombo(.Item(lngCont).SubItems(COL_OP_TP_ACAO_OPER_ATIV_EXEC)))
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "NumeroComandoAcao", Split(.Item(lngCont).Tag, "|")(POS_NU_COMD_ACAO_EXEC), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
                        intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed, 1, 0)
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "IgnoraGradeHorario", intIgnoraGradeHorario, "Repeat_Filtros")
                    End If
                Next
            End With
            
            If lngItensChecked > 0 Then
                Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
                strXMLRetorno = strXMLRetorno & objOperacaoMensagem.Liberar(xmlDomLoteOperacaoMensagem.xml, vntCodErro, vntMensagemErro)
                Set objOperacaoMensagem = Nothing
            End If
            
            lngContIni = lngContFim + 1
            lngContFim = lngContFim + QTD_MAX_CONFIRMACAO
            
            Set xmlDomLoteOperacaoMensagem = Nothing
        Next
        
        Set xmlDomLoteOperacaoMensagem = Nothing
    End If
    
    Set xmlDomLoteOperacaoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDomLoteOperacaoMensagem, "", "Repeat_Filtros", "")

    lngItensChecked = 0
        
    'Captura o filtro cumulativo MENSAGEM
    With lstMensagem.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Repeat_Filtros", "Grupo_Lote", "")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.MENSAGEM, "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "NumeroControleIF", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "DTRegistroMensagemSPB", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MENSAGEM_SPB), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "CodigoMensagem", .Item(lngCont).SubItems(COL_MSG_CODIGO_MENSAGEM), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Operacao", Split(.Item(lngCont).Tag, "|")(POS_NU_SEQU_OPER_ATIV_MENSAGEM), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "ProtocoloAlcada", Split(.Item(lngCont).Tag, "|")(POS_NU_PRTC_MESG_LG), "Repeat_Filtros")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "ValorMensagem", Split(.Item(lngCont).Tag, "|")(POS_VA_FINC), "Repeat_Filtros")
                intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed, 1, 0)
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "IgnoraGradeHorario", intIgnoraGradeHorario, "Repeat_Filtros")
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.Liberar(xmlDomLoteOperacaoMensagem.xml, vntCodErro, vntMensagemErro)
        Set objOperacaoMensagem = Nothing
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        'Verifica se o retorno da LIBERAÇÃO possui erros...
        If strXMLRetorno <> vbNullString Then
            '...se sim, carrega o XML de Erros
            Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlRetornoErro.loadXML(strXMLRetorno)
        Else
            '...se não, apenas destrói o objeto
            Set xmlRetornoErro = Nothing
        End If
        
        flLiberar = strXMLRetorno
    End If
    
    If strXMLRetorno <> vbNullString Then
        strXMLRetorno = "<Retorno>" & strXMLRetorno & "</Retorno>"
    Else
        strXMLRetorno = vbNullString
    End If
    
    flLiberar = strXMLRetorno
    Set xmlDomLoteOperacaoMensagem = Nothing
    
    Set xmlDomLoteOperacaoMensagem = Nothing

Exit Function
ErrorHandler:
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - flLiberar", Me.Caption

End Function

Private Function flLiberarContingencia() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim xmlDomLoteOperacaoMensagem              As MSXML2.DOMDocument40
Dim strXMLRetorno                           As String
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long
Dim intIgnoraGradeHorario                   As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Const POS_STATUS                            As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1
Const POS_NU_COMD_ACAO_EXEC                 As Integer = 2

Const POS_NUMERO_CONTROLE_IF                As Integer = 0
Const POS_DATA_REGISTRO_MENSAGEM_SPB        As Integer = 1
Const POS_NU_SEQU_OPER_ATIV_MENSAGEM        As Integer = 2
Const POS_NU_PRTC_MESG_LG                   As Integer = 3
Const POS_VA_FINC                           As Integer = 4

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ---------------------------------------------------------------------------
    Set xmlDomLoteOperacaoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomLoteOperacaoMensagem, "", "Repeat_Filtros", "")
    
    'Captura o filtro cumulativo MENSAGEM
    With lstMensagem.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Repeat_Filtros", "Grupo_Lote", "")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.MENSAGEM, "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "NumeroControleIF", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "DTRegistroMensagemSPB", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MENSAGEM_SPB), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "CodigoMensagem", .Item(lngCont).SubItems(COL_MSG_CODIGO_MENSAGEM), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "Operacao", Split(.Item(lngCont).Tag, "|")(POS_NU_SEQU_OPER_ATIV_MENSAGEM), "Repeat_Filtros")
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "ProtocoloAlcada", Split(.Item(lngCont).Tag, "|")(POS_NU_PRTC_MESG_LG), "Repeat_Filtros")
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "ValorMensagem", Split(.Item(lngCont).Tag, "|")(POS_VA_FINC), "Repeat_Filtros")
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "Contingencia", "1", "Repeat_Filtros")
                          
                intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed, 1, 0)
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "IgnoraGradeHorario", intIgnoraGradeHorario, "Repeat_Filtros")
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.Liberar(xmlDomLoteOperacaoMensagem.xml, _
                                                    vntCodErro, _
                                                    vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objOperacaoMensagem = Nothing
        
        'Verifica se o retorno da LIBERAÇÃO possui erros...
        If strXMLRetorno <> vbNullString Then
            '...se sim, carrega o XML de Erros
            Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlRetornoErro.loadXML(strXMLRetorno)
        Else
            '...se não, apenas destrói o objeto
            Set xmlRetornoErro = Nothing
        End If
        
        flLiberarContingencia = strXMLRetorno
    Else
        flLiberarContingencia = vbNullString
    End If
    
    Set xmlDomLoteOperacaoMensagem = Nothing

Exit Function
ErrorHandler:
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - flLiberarContingencia", Me.Caption

End Function

'Limpa o conteúdo do ListView
Private Sub flLimparLista(ByVal lstListView As ListView)
    lstListView.ListItems.Clear
End Sub

'Desmarca a seleção de todos os itens do ListView
Private Sub flMarcarDesmarcarTodas(ByVal lstListView As ListView, _
                                   ByVal plngTipoSelecao As enumTipoSelecao)

Dim lngLinha                                As Long

On Error GoTo ErrorHandler

    For lngLinha = 1 To lstListView.ListItems.Count
        'Tratamento específico para não desmarcar o CHECK BOX do item TOTAL, no ListView de Operação
        If lstListView.Name = "lstOperacao" Then
            If lstListView.ListItems(lngLinha).Key = "kTotal" Then
                Exit For
            End If
        End If
        
        lstListView.ListItems(lngLinha).Checked = (plngTipoSelecao = enumTipoSelecao.MarcarTodas)
    Next

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMarcarDesmarcarTodas", 0

End Sub

'Exibe de forma diferenciada todos os itens que tenham sido rejeitados por motivo de grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            If Not objDomNode.selectSingleNode("Operacao") Is Nothing Then
                With lstOperacao.ListItems
                    For lngCont = 1 To .Count
                        If UCase(Mid(.Item(lngCont).Key, 2)) = UCase(objDomNode.selectSingleNode("Operacao").Text) Then
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
            Else
                With lstMensagem.ListItems
                    For lngCont = 1 To .Count
                        If UCase(Mid(.Item(lngCont).Key, 2)) = UCase(objDomNode.selectSingleNode("NumeroControleIF").Text) & "|" & _
                                                               UCase(objDomNode.selectSingleNode("DTRegistroMensagemSPB").Text) Then
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
            End If
        Next
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0

End Sub

'Mostra o resultado do último processamento
Private Sub flMostrarResultado(ByVal pstrResultadoLiberacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " liberados "
        .Resultado = pstrResultadoLiberacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'Configura o perfil de acesso do usuário
Property Get PerfilAcesso() As enumPerfilAcesso
    PerfilAcesso = lngPerfil
End Property

'Configura o perfil de acesso do usuário
Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
    
    lngPerfil = pPerfil

    Select Case lngPerfil
        Case enumPerfilAcesso.Nenhum
            Me.Caption = "Liberação - Operação e Mensagem SPB"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "Controle de Alçadas - Liberação Administrador de Área"
        Case enumPerfilAcesso.AdmGeral
            Me.Caption = "Controle de Alçadas - Liberação Administrador Geral"
    End Select
    
    If lngPerfil = enumPerfilAcesso.AdmArea Or lngPerfil = enumPerfilAcesso.AdmGeral Then
        tlbFiltro.Buttons("MostrarOperacao").Visible = False
        tlbFiltro.Buttons("MostrarMensagem").Visible = False
        Call flArranjarJanelasExibicao("2")
    Else
        tlbFiltro.Buttons("MostrarOperacao").Visible = True
        tlbFiltro.Buttons("MostrarMensagem").Visible = True
        tlbFiltro.Buttons("MostrarOperacao").value = tbrPressed
        tlbFiltro.Buttons("MostrarMensagem").value = tbrPressed
        Call flArranjarJanelasExibicao("12")
        Call flLimparLista(lstMensagem)
        Call flLimparLista(lstOperacao)
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("refresh"))

    End If

End Property

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, _
             enumTipoSelecao.DesmarcarTodas
            
            Call flMarcarDesmarcarTodas(IIf(intControleMenuPopUp = enumTipoConfirmacao.operacao, _
                                                lstOperacao, lstMensagem), _
                                        Retorno)
    End Select
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - ctlMenu1_ClickMenu", Me.Caption
    
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
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    If lngPerfil = enumPerfilAcesso.AdmArea Or lngPerfil = enumPerfilAcesso.AdmGeral Then
        tlbFiltro.Buttons("MostrarOperacao").Visible = False
        tlbFiltro.Buttons("MostrarMensagem").Visible = False
        Call flArranjarJanelasExibicao("2")
    End If
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    With Me
        .lstMensagem.ListItems.Clear
        .lstOperacao.ListItems.Clear
    End With
    
    Call fgCursor(True)
    
    Call flInicializar
    
    blnPrimeiraConsulta = True
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmLiberacaoOperacaoMensagem
    Load objFiltro
    objFiltro.fgCarregarPesquisaAnterior
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .tlbFiltro.Top = .ScaleHeight - .tlbFiltro.Height
        
        .imgDummyH.Left = 0
        .imgDummyH.Width = .ScaleWidth
        
        .lstOperacao.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .imgDummyH.Height, _
                                                      .tlbFiltro.Top - 300)
        .lstOperacao.Width = .Width - 100
        
        .lstMensagem.Top = IIf(.imgDummyH.Visible, .imgDummyH.Top + .imgDummyH.Height, 0)
        .lstMensagem.Height = IIf(.imgDummyH.Visible, .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 300, _
                                                      .tlbFiltro.Top - 300)
        .lstMensagem.Width = .Width - 100
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set frmLiberacaoOperacaoMensagem = Nothing
End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = True
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not fblnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyH.Top = y + imgDummyH.Top

    On Error Resume Next
    
    With Me
        If .imgDummyH.Top < 950 Then
            .imgDummyH.Top = 950
        End If
        If .imgDummyH.Top > (.Height - 2100) And (.Height - 2100) > 0 Then
            .imgDummyH.Top = .Height - 2100
        End If
        
        .lstOperacao.Height = .imgDummyH.Top - .imgDummyH.Height
        .lstMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        .lstMensagem.Height = .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 300
    End With
    
    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = False
End Sub

Private Sub lstMensagem_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call fgClassificarListview(Me.lstMensagem, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index
End Sub

Private Sub lstMensagem_DblClick()

Dim strChave                                As String

Const POS_NUMERO_CTRL_IF                    As Integer = 0
Const POS_DATA_REGISTRO_MESG                As Integer = 1
Const POS_NUMERO_SEQUENCIA_OPERACAO         As Integer = 2

On Error GoTo ErrorHandler

    If Not lstMensagem.SelectedItem Is Nothing Then
        strChave = Mid$(lstMensagem.SelectedItem.Key, 2)
        With frmDetalheOperacao
            .SequenciaOperacao = Split(lstMensagem.SelectedItem.Tag, "|")(POS_NUMERO_SEQUENCIA_OPERACAO)
            .NumeroControleIF = Split(strChave, "|")(POS_NUMERO_CTRL_IF)
            .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(strChave, "|")(POS_DATA_REGISTRO_MESG))
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - lstMensagem_DblClick", Me.Caption

End Sub

Private Sub lstMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim objListItem                             As ListItem
    
    If Left(Item.ListSubItems(3).Text, 7) = "BMC0015" Then
        tlbFiltro.Buttons("Liberar").Enabled = False
        tlbFiltro.Buttons("Rejeitar").Enabled = False
        tlbFiltro.Buttons("LibConting").Enabled = True
        
        'Tira o checked de todas as mensagens <> BMC0015
        For Each objListItem In lstMensagem.ListItems
            With objListItem
                If Left(.ListSubItems(3), 7) <> "BMC0015" Then
                    If .Checked Then
                        objListItem.Checked = False
                    End If
                End If
            End With
        Next
        
        'Tira o checked de todas as operacoes
        For Each objListItem In lstOperacao.ListItems
            With objListItem
                If .Checked Then
                    objListItem.Checked = False
                End If
            End With
        Next
        
    Else
        tlbFiltro.Buttons("Liberar").Enabled = True
        tlbFiltro.Buttons("Rejeitar").Enabled = True
        tlbFiltro.Buttons("LibConting").Enabled = False
        
        'Tira o checked de todas as mensagens BMC0015
        For Each objListItem In lstMensagem.ListItems
            With objListItem
                If Left(.ListSubItems(3), 7) = "BMC0015" Then
                    If .Checked Then
                        objListItem.Checked = False
                    End If
                End If
            End With
        Next
        
    End If
    
    Item.Selected = True
    
End Sub

Private Sub lstMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lstOperacao.SelectedItem = Nothing
End Sub

Private Sub lstMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.MENSAGEM
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - lstMensagem_MouseDown", Me.Caption
End Sub

Private Sub lstOperacao_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifListOper = ColumnHeader.Index
End Sub

Private Sub lstOperacao_DblClick()
    
On Error GoTo ErrorHandler

    If Not lstOperacao.SelectedItem Is Nothing Then
        If lstOperacao.SelectedItem.Key <> "kTotal" Then
            With frmDetalheOperacao
                .SequenciaOperacao = Mid(lstOperacao.SelectedItem.Key, 2)
                .Show vbModal
            End With
        End If
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_DblClick"
    
End Sub

Private Sub lstOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim objListItem                             As ListItem

    Item.Selected = True
    
    If Item.Key = "kTotal" Then
        Item.Checked = True
    End If
    
    tlbFiltro.Buttons("Liberar").Enabled = True
    tlbFiltro.Buttons("Rejeitar").Enabled = True
    tlbFiltro.Buttons("LibConting").Enabled = False

    'Tira o checked de todas as mensagens BMC0015
    For Each objListItem In lstMensagem.ListItems
        With objListItem
            If Left(.ListSubItems(3), 7) = "BMC0015" Then
                If .Checked Then
                    objListItem.Checked = False
                End If
            End If
        End With
    Next

End Sub

Private Sub lstOperacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lstMensagem.SelectedItem = Nothing
End Sub

Private Sub lstOperacao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.operacao
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacao - lstOperacao_MouseDown", Me.Caption
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

On Error GoTo ErrorHandler

    strFiltroXML = xmlDocFiltros
    fgCursor True
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            blnPrimeiraConsulta = False
            Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("DefinirFiltro"))
        End If
        
        tlbFiltro.Buttons("AplicarFiltro").value = tbrPressed
        
        If InStr(1, strFiltroXML, "DataIni") = 0 Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA OPERAÇÃO."
            frmMural.Show vbModal
        
            Call flLimparLista(lstOperacao)
            Call flLimparLista(lstMensagem)
        Else
            If lngPerfil = enumPerfilAcesso.Nenhum Then
                Call flCarregarListaOperacao(strFiltroXML)
            End If
            Call flCarregarListaMensagem(strFiltroXML)
        End If
    End If
    
    fgCursor
    
Exit Sub
ErrorHandler:

    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - objFiltro_AplicarFiltro", Me.Caption
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strJanelas                              As String
Dim strSelecaoFiltro                        As String
Dim strResultadoProcessamento                   As String

Dim blnUtilizaFiltro                        As Boolean

Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1
Const POS_NUM_CONTROLE_IF                   As Integer = 0
Const POS_DATA_REGISTRO_MENSAGEM            As Integer = 1

Dim xmlProcessamento                        As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    Call fgCursor(True)
    
    'Verifica se o filtro deve ser aplicado
    blnUtilizaFiltro = (tlbFiltro.Buttons("AplicarFiltro").value = tbrPressed)
    
    If tlbFiltro.Buttons("MostrarOperacao").Visible Then
        If tlbFiltro.Buttons("MostrarOperacao").value = tbrPressed Then
            strJanelas = strJanelas & "1"
        End If
        
        If tlbFiltro.Buttons("MostrarMensagem").value = tbrPressed Then
            strJanelas = strJanelas & "2"
        End If
        
        Call flArranjarJanelasExibicao(strJanelas)
    End If
    
    Select Case Button.Key
        Case "DefinirFiltro"
            blnPrimeiraConsulta = False
            objFiltro.Show vbModal
            
            tlbFiltro.Buttons("Liberar").Enabled = True
            tlbFiltro.Buttons("Rejeitar").Enabled = True
            tlbFiltro.Buttons("LibConting").Enabled = False
            
        Case "refresh"
            If InStr(1, strFiltroXML, "DataIni") = 0 Or Not blnUtilizaFiltro Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA OPERAÇÃO."
                frmMural.Show vbModal
            Else
                objFiltro.fgCarregarPesquisaAnterior
                Call flMarcarRejeitadosPorGradeHorario
            End If
            
            tlbFiltro.Buttons("Liberar").Enabled = True
            tlbFiltro.Buttons("Rejeitar").Enabled = True
            tlbFiltro.Buttons("LibConting").Enabled = False
            
        Case "Liberar"
            strResultadoProcessamento = flLiberar
            If strResultadoProcessamento <> vbNullString Then
                Call flMostrarResultado(strResultadoProcessamento)
                objFiltro.fgCarregarPesquisaAnterior
                Call flMarcarRejeitadosPorGradeHorario
            Else
                frmMural.Display = "Não existem itens selecionados para a liberação."
                frmMural.IconeExibicao = IconExclamation
                frmMural.Show vbModal
            End If
            
        Case "Rejeitar"
            Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlProcessamento.loadXML(flMontarXMLProcessamento)
        
            If xmlProcessamento.xml <> vbNullString Then
                strResultadoProcessamento = fgMIUExecutarGenerico("ProcessarEmLote", "A8LQS.clsMensagem", xmlProcessamento)
            End If
        
            Set xmlProcessamento = Nothing
        
            If strResultadoProcessamento <> vbNullString Then
                Call fgMostrarResultado(strResultadoProcessamento, "rejeitados")
                Call objFiltro.fgCarregarPesquisaAnterior
            End If
            
        Case "LibConting"
            strResultadoProcessamento = flLiberarContingencia
            If strResultadoProcessamento <> vbNullString Then
                Call flMostrarResultado(strResultadoProcessamento)
                objFiltro.fgCarregarPesquisaAnterior
                Call flMarcarRejeitadosPorGradeHorario
            Else
                frmMural.Display = "Não existem itens selecionados para a liberação."
                frmMural.IconeExibicao = IconExclamation
                frmMural.Show vbModal
            End If
        
        Case gstrSair
            Unload Me
            
    End Select
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacaoMensagem - tlbFiltro_ButtonClick", Me.Caption

End Sub

'Monta string XML para processamento em lote
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemProc                             As MSXML2.DOMDocument40
Dim objTipoOperacao                         As A8MIU.clsOperacao

Const POS_NU_CTRL_IF                        As Integer = 0
Const POS_DH_REGT_MESG_SPB                  As Integer = 1
Const POS_NU_SEQU_OPER_ATIV                 As Integer = 2
Const POS_TP_OPER                           As Integer = 5
    
    On Error GoTo ErrorHandler
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, vbNullString, "Repeat_Processamento", vbNullString)
    
    For Each objListItem In lstMensagem.ListItems
        With objListItem
            If .Checked Then
                
                If .Key = "kTotal" Then Exit For
                
                Select Case Val(Split(.Tag, "|")(POS_TP_OPER))
                    
                    Case enumTipoOperacaoLQS.EnvioPAG0105Clientes, enumTipoOperacaoLQS.EnvioPAG0106Clientes, enumTipoOperacaoLQS.EnvioPAG0108Clientes, _
                         enumTipoOperacaoLQS.EnvioPAG0109Clientes, enumTipoOperacaoLQS.EnvioPAG0121Clientes, enumTipoOperacaoLQS.EnvioPAG0134Clientes, _
                         enumTipoOperacaoLQS.EnvioTEDSTR0006Clientes, enumTipoOperacaoLQS.EnvioTEDSTR0007Clientes, enumTipoOperacaoLQS.EnvioTEDSTR0008Clientes, _
                         enumTipoOperacaoLQS.EnvioTEDSTR0009Clientes, enumTipoOperacaoLQS.EnvioTEDSTR0025Clientes, enumTipoOperacaoLQS.EnvioTEDSTR0034Clientes, _
                         enumTipoOperacaoLQS.EmissaoTEDPAG0106FdosFIDC, enumTipoOperacaoLQS.EmissaoTEDPAG0108FdosFIDC, enumTipoOperacaoLQS.EmissaoTEDSTR0007FdosFIDC, _
                         enumTipoOperacaoLQS.EmissaoTEDSTR0008FdosFIDC, enumTipoOperacaoLQS.EnvioSTR0006PagDespesas, enumTipoOperacaoLQS.EnvioSTR0007PagDespesas, _
                         enumTipoOperacaoLQS.EnvioSTR0008PagDespesas, enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta, enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib, _
                         enumTipoOperacaoLQS.EnvioSTR0007PagDespesas, enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta, enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta, _
                         enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib, enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoCBLC, enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoCBLC, _
                         enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoCBLC, enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoCBLC, enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoBMF, _
                         enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoBMF, enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoBMF, enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoBMF, _
                         enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoSTR, enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoSTR, enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoSTR, _
                         enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoSTR
                         
                            Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                    
                            Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "RejeitarTED")
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", "A8LQS.clsMensagem")
                            
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_CTRL_IF", Mid$(Split(.Key, "|")(POS_NU_CTRL_IF), 2))
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "DH_REGT_MESG_SPB", Split(.Key, "|")(POS_DH_REGT_MESG_SPB))
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_SEQU_OPER_ATIV", Split(.Tag, "|")(POS_NU_SEQU_OPER_ATIV))
                            
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_SEQU_CNCL_OPER_ATIV_MESG", .SubItems(COL_MSG_CNCL))
                                    
                            Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemProc.xml)
                            
                            Set xmlItemProc = Nothing
                         
                    Case Else
                    
                        frmMural.Display = "Rejeição permitida apenas para mensagens: " & vbNewLine & vbNewLine & _
                                       " - Emissão Operação CCR;" & vbNewLine & _
                                       " - Negociação Operação CCR;" & vbNewLine & _
                                       " - Devolução RecolhimentoEstorno Reembolso CCR;" & vbNewLine & _
                                       " - Consulta Limite Importação CCR;" & vbNewLine & _
                                       " - Consulta Operações CCR;" & vbNewLine & _
                                       " - STR e PAG - Envio de TED a Clientes;" & vbNewLine & _
                                       " - STR - Pagamento de Despesas."
                                       
                        frmMural.Show vbModal
                        
                        Exit Function
                
                End Select
                
            End If
        
        End With
    Next

    For Each objListItem In lstOperacao.ListItems
        With objListItem
            If .Checked Then
                
                If .Key = "kTotal" Then Exit For
                
                Select Case Val(Split(.Tag, "|")(3))
                    
                    Case enumTipoOperacaoLQS.EnvioPagDespesasBoleto, enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta, enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib, _
                         enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente, enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta, enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib, _
                         enumTipoOperacaoLQS.EnvioPagDespesasTributos, enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta, enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib, _
                         enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais, enumTipoOperacaoLQS.LancamentoCCCashFlow, enumTipoOperacaoLQS.LancamentoCCCashFlowStrikeFixo, _
                         enumTipoOperacaoLQS.LancamentoCCSwapJuros, enumTipoOperacaoLQS.ConsultaOperacaoCCR, enumTipoOperacaoLQS.EmissaoOperacaoCCR, _
                         enumTipoOperacaoLQS.NegociacaoOperacaoCCR, enumTipoOperacaoLQS.DevolucaoRecolhimentoEstornoReembolsoCCR, enumTipoOperacaoLQS.ConsultaLimitesImportacaoCCR, _
                         enumTipoOperacaoLQS.ContratacaoMercadoPrimario, enumTipoOperacaoLQS.EdicaoContratacaoMercadoPrimario, _
                         enumTipoOperacaoLQS.ConfirmacaoEdicaoContratacaoMercadoPrimario, enumTipoOperacaoLQS.AlteracaoContrato, _
                         enumTipoOperacaoLQS.EdicaoAlteracaoContrato, enumTipoOperacaoLQS.ConfirmacaoEdicaoAlteracaoContrato, _
                         enumTipoOperacaoLQS.LiquidacaoMercadoPrimario, enumTipoOperacaoLQS.BaixaValorLiquidar, enumTipoOperacaoLQS.RestabelecimentoBaixa, _
                         enumTipoOperacaoLQS.CancelamentoValorLiquidar, enumTipoOperacaoLQS.EdicaoCancelamentoValorLiquidar, _
                         enumTipoOperacaoLQS.ConfirmacaoEdicaoCancelamentoValorLiquidar, enumTipoOperacaoLQS.VinculacaoContratos, enumTipoOperacaoLQS.AnulacaoEvento, _
                         enumTipoOperacaoLQS.CorretoraRequisitaClausulasEspecificas, enumTipoOperacaoLQS.IFInformaClausulasEspecificas, _
                         enumTipoOperacaoLQS.ManutencaoCadastroAgenciaCentralizadoraCambio, enumTipoOperacaoLQS.CredenciamentoDescredenciamentoDispostoRMCCI, _
                         enumTipoOperacaoLQS.IncorporacaoContratos, enumTipoOperacaoLQS.AceiteRejeicaoIncorporacaoContratos, _
                         enumTipoOperacaoLQS.ConsultaContratosEmSer, enumTipoOperacaoLQS.ConsultaEventosUmDia, enumTipoOperacaoLQS.ConsultaDetalhamentoContratoInterbancario, _
                         enumTipoOperacaoLQS.ConsultaEventosContratoMercadoPrimario, enumTipoOperacaoLQS.ConsultaEventosContratoIntermediadoMercadoPrimario, _
                         enumTipoOperacaoLQS.ConsultaHistoricoIncorporacoes, enumTipoOperacaoLQS.ConsultaContratosIncorporacao, enumTipoOperacaoLQS.ConsultaCadeiaIncorporacoesContrato, _
                         enumTipoOperacaoLQS.ConsultaPosicaoCambioMoeda, enumTipoOperacaoLQS.AtualizaçãoInclusãoInstrucoesPagamento, enumTipoOperacaoLQS.ConsultaInstrucoesPagamento, _
                         enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega, enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega, _
                         enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara, enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara, _
                         enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega, enumTipoOperacaoLQS.InformaContrArbitParceiroExteriorPaisPropriaIF, _
                         enumTipoOperacaoLQS.InformaOperacaoArbitragemParceiroPais, enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais, _
                         enumTipoOperacaoLQS.CAMInformaContratacaoInterbancarioViaLeilao, enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria, _
                         enumTipoOperacaoLQS.ConsultaContratosCambioMercadoInterbancario
                         
                            Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                    
                            Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "RejeitarOperacao")
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", "A8LQS.clsOperacao")
                    
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_SEQU_OPER_ATIV", Mid$(Split(.Key, "|")(POS_NU_CTRL_IF), 2))
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "DH_REGT_MESG_SPB", Split(.Tag, "|")(POS_DH_REGT_MESG_SPB))
                            
                            Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemProc.xml)
                    
                            Set xmlItemProc = Nothing
                    
                    Case Else
                    
                        frmMural.Display = "Rejeição permitida apenas para operações: " & vbNewLine & vbNewLine & _
                                       " - Emissão Operação CCR;" & vbNewLine & _
                                       " - Negociação Operação CCR;" & vbNewLine & _
                                       " - Devolução RecolhimentoEstorno Reembolso CCR;" & vbNewLine & _
                                       " - Consulta Limite Importação CCR;" & vbNewLine & _
                                       " - Consulta operações CCR;" & vbNewLine & _
                                       " - Pagamento de Despesas;" & vbNewLine & _
                                       " - Lançamento em Conta Corrente - Operações Manuais;" & vbNewLine & _
                                       " - Operações CAM."
                                       
                        frmMural.Show vbModal
                        
                        Exit Function
                
                End Select
                
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
