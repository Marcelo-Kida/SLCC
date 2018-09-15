VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfirmacaoOperacao 
   Caption         =   "Ferramentas - Confirmação Operação/Mensagem SPB"
   ClientHeight    =   8970
   ClientLeft      =   240
   ClientTop       =   1335
   ClientWidth     =   14880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin A8.ctlMenu ctlMenu1 
      Left            =   2280
      Top             =   8040
      _ExtentX        =   2990
      _ExtentY        =   661
   End
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
      NumItems        =   29
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
         Text            =   "Data Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Veículo Legal (Parte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contra-Parte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo Movto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Título"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Data Vencto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Tipo Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Local Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Tipo Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Text            =   "Agendamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Cta. Cedente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Cta. Cessionário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Canal Venda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "CNPJ/CPF Comitente"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Código do Usuário"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Nome Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Moeda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Valor ME"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Taxa Cambial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Código Veículo Legal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Tipo Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Tipo Contrato"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8640
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   582
      ButtonWidth     =   2725
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar Filtro  "
            Key             =   "AplicarFiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro "
            Key             =   "DefinirFiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela "
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confirmar      "
            Key             =   "Confirmar"
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rejeitar        "
            Key             =   "Rejeitar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agendamento"
            Key             =   "Agendamento"
            Object.ToolTipText     =   "Alterar Agendamento"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Operações    "
            Key             =   "MostrarOperacao"
            Object.ToolTipText     =   "Mostrar Operações"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mensagens   "
            Key             =   "MostrarMensagem"
            Object.ToolTipText     =   "Mostrar Mensagens"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
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
            Picture         =   "frmConfirmacaoOperacao.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfirmacaoOperacao.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMensagem 
      Height          =   2565
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   4524
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
      NumItems        =   7
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
         Text            =   "Código Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Agendamento"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   4800
      Width           =   14040
   End
End
Attribute VB_Name = "frmConfirmacaoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:46
'-------------------------------------------------
'' Objeto responsável pela confirmação de mensagens e operações através de
'' interação com a camada de controle de caso de uso MIU.
''
'' Classes especificamente consideradas de destino:
''   A8MIU.clsOperacao
''   A8MIU.clsOperacaoMensagem
''   A8MIU.clsMensagem
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1
Private strFiltroXML                        As String

Private fblnDummyH                          As Boolean
Private blnPrimeiraConsulta                 As Boolean

Private intControleMenuPopUp                As enumTipoConfirmacao

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_OPERACAO               As Integer = 1
Private Const COL_OP_DATA_OPERACAO          As Integer = 2
Private Const COL_OP_NUMERO_COMANDO         As Integer = 3
Private Const COL_OP_VEICULO_LEGAL_PARTE    As Integer = 4
Private Const COL_OP_CONTRAPARTE            As Integer = 5
Private Const COL_OP_SITUACAO               As Integer = 6
Private Const COL_OP_TIPO_MOVIMENTO         As Integer = 7
Private Const COL_OP_TITULO                 As Integer = 8
Private Const COL_OP_VALOR                  As Integer = 9
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 10
Private Const COL_OP_TIPO_OPERACAO          As Integer = 11
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 12
Private Const COL_OP_LOCAL_LIQUIDACAO       As Integer = 13
Private Const COL_OP_TIPO_LIQUIDACAO        As Integer = 14
Private Const COL_OP_EMPRESA                As Integer = 15
Private Const COL_OP_HORARIO_ENVIO_MSG      As Integer = 16     '<-- Agendamento
Private Const COL_OP_CONTA_CEDENTE          As Integer = 17
Private Const COL_OP_CONTA_CESSIONARIO      As Integer = 18
Private Const COL_OP_CANAL_VENDA            As Integer = 19
Private Const COL_OP_CNPJ_CPF_COMITENTE     As Integer = 20
Private Const COL_OP_CODIGO_USUARIO         As Integer = 21
Private Const COL_OP_NO_CLIE                As Integer = 22
Private Const COL_OP_CD_MOED_ISO            As Integer = 23
Private Const COL_OP_VA_MOED_ESTR           As Integer = 24
Private Const COL_OP_NR_PERC_TAXA_CAMB      As Integer = 25
Private Const COL_OP_CO_VEIC_LEGA           As Integer = 26
Private Const COL_TO_CO_MESG_SPB_REGT_OPER  As Integer = 27
Private Const COL_OP_TP_OPER_CAMB           As Integer = 28

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_DATA_MENSAGEM         As Integer = 1
Private Const COL_MSG_NUMERO_COMANDO        As Integer = 2
Private Const COL_MSG_CODIGO_MENSAGEM       As Integer = 3
Private Const COL_MSG_SITUACAO              As Integer = 4
Private Const COL_MSG_EMPRESA               As Integer = 5
Private Const COL_MSG_HORARIO_ENVIO_MSG     As Integer = 6      '<-- Agendamento

Private Const strFuncionalidade             As String = "frmConfirmacaoOperacao"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'' Confirma o lote das operações e/ou mensagens selecionados no objeto, através de
'' interação com a camada de controle de caso de uso MIU, método A8MIU.
'' clsOperacaoMensagem.Confirmar,  e retorna uma string contendo o resultado da
'' operação de lote.
Private Function flConfirmar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim xmlDomLoteOperacaoMensagem              As MSXML2.DOMDocument40
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim strXMLRetorno                           As String
Dim lngQuebras                              As Long
Dim lngContQuebras                          As Long
Dim lngContIni                              As Long
Dim lngContFim                              As Long

Const POS_STATUS                            As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1
Const POS_NUMERO_CONTROLE_IF                As Integer = 0
Const POS_DATA_REGISTRO_MENSAGEM_SPB        As Integer = 1
Const POS_CODIGO_MENSAGEM                   As Integer = 2
Const POS_PROTOCOLO_OPERACAO_LG             As Integer = 4
Const QTD_MAX_CONFIRMACAO                   As Integer = 500

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
            
            With lstOperacao.ListItems
                For lngCont = lngContIni To lngContFim
                    If .Item(lngCont).Checked And .Item(lngCont).SubItems(COL_OP_OPERACAO) <> vbNullString Then
                        lngItensChecked = lngItensChecked + 1
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Repeat_Filtros", "Grupo_Lote", "")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.operacao, "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Operacao", Mid(.Item(lngCont).Key, 2), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "CodigoMensagem", Split(.Item(lngCont).Tag, "|")(POS_CODIGO_MENSAGEM), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
                        Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Grupo_Lote", "Protocolo", Split(.Item(lngCont).Tag, "|")(POS_PROTOCOLO_OPERACAO_LG), "Repeat_Filtros")
                    End If
                Next
            End With
            
            If lngItensChecked > 0 Then
                Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
                strXMLRetorno = strXMLRetorno & objOperacaoMensagem.Confirmar(xmlDomLoteOperacaoMensagem.xml, vntCodErro, vntMensagemErro)
                Set objOperacaoMensagem = Nothing
            End If
            
            lngContIni = lngContFim + 1
            lngContFim = lngContFim + QTD_MAX_CONFIRMACAO
            
            Set xmlDomLoteOperacaoMensagem = Nothing
        Next
    End If
    
    Set xmlDomLoteOperacaoMensagem = Nothing
    
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
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = strXMLRetorno & objOperacaoMensagem.Confirmar(xmlDomLoteOperacaoMensagem.xml, vntCodErro, vntMensagemErro)
        Set objOperacaoMensagem = Nothing
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    End If
    
    If strXMLRetorno <> vbNullString Then
        strXMLRetorno = "<Retorno>" & strXMLRetorno & "</Retorno>"
    Else
        strXMLRetorno = vbNullString
    End If
    
    flConfirmar = strXMLRetorno
    Set xmlDomLoteOperacaoMensagem = Nothing

Exit Function
ErrorHandler:
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flConfirmar", 0

End Function

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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    
    Call flInicializar
    
    blnPrimeiraConsulta = True
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmConfirmacaoOperacao
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    
    Call flSelecionaLista(enumTipoConfirmacao.operacao)
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - Form_Load", Me.Caption
    
End Sub

'' Carrega a lista das operações passíveis de confirmação, através de interação
'' com a camada de controle de caso de uso MIU, método A8MIU.clsOperacao.
'' ObterDetalheOperacao, e preenche o listview de operações com as mesmas
Private Sub flCarregarListaOperacao(Optional ByVal pstrxmlFiltro As String = vbNullString)

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
Dim lngCont                                 As Long
Dim dblTotalOperacao                        As Double

Dim strOperacoesBMA                         As String
Dim strOperacoesCETIP                       As String
Dim strOperacoesBMD                         As String
Dim strOperacoesBMC                         As String
Dim strOperacoesCBLC                        As String

Dim strOperacoesSTR                         As String
Dim strOperacoesPAG                         As String
Dim strOperacoesContaCorrente               As String

'KIDA - CCR
Dim strOperacoesCCR                         As String
Dim strOperacoesCAM                         As String

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Call flLimparLista(lstOperacao)
    
    strSelecaoFiltro = enumStatusOperacao.EmSer & ";" & _
                       enumStatusOperacao.ManualEmSer & ";" & _
                       enumStatusOperacao.Inconsistencia & ";" & _
                       enumStatusOperacao.ReativacaoSolicitada
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Verifica se o filtro foi informado...
    If pstrxmlFiltro <> vbNullString Then
        '...se sim, lê o filtro existente e adiciona o filtro de STATUS
        Call xmlDomFiltros.loadXML(pstrxmlFiltro)
    Else
        '...se não, cria um novo filtro apenas para o envio do filtro de STATUS
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    '>>> -------------------------------------------------------------------------------------------
    
    strOperacoesBMA = enumTipoOperacaoLQS.TransferenciaBMA & ", " & _
                      enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA & ", " & _
                      enumTipoOperacaoLQS.OperacaoTermoInternaBMA & ", " & _
                      enumTipoOperacaoLQS.EspecDefinitivaIntermediacao & ", " & _
                      enumTipoOperacaoLQS.EspecDefinitivaCobertura & ", " & _
                      enumTipoOperacaoLQS.EspecTermoIntermediacao & ", " & _
                      enumTipoOperacaoLQS.EspecTermoCobertura & ", " & _
                      enumTipoOperacaoLQS.DepositoBMA & ", " & _
                      enumTipoOperacaoLQS.RetiradaBMA & ", " & _
                      enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA & ", " & _
                      enumTipoOperacaoLQS.CancelamentoEspecificacaoBMA & ", " & _
                      enumTipoOperacaoLQS.EspecCompromissadaIntermediacao & ", " & _
                      enumTipoOperacaoLQS.EspecCompromissadaCobertura & ", " & _
                      enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA & ", " & _
                      enumTipoOperacaoLQS.OperacaoCompromissadaInternaBMA & ", " & _
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

    strOperacoesBMC = enumTipoOperacaoLQS.TransferenciasBMCDeposito & ", " & _
                      enumTipoOperacaoLQS.TransferenciasBMCRetirada & ", " & _
                      enumTipoOperacaoLQS.DespesasBMC & ", " & _
                      enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC & ", " & _
                      enumTipoOperacaoLQS.RegistroOperacoesBMC & ", " & _
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
                      

    strOperacoesBMD = enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF & ", " & _
                      enumTipoOperacaoLQS.NETEntradaManualMultilateralBMD

    strOperacoesCBLC = enumTipoOperacaoLQS.RegistroLiquidacaoEventoCBLC & ", " & _
                       enumTipoOperacaoLQS.NETEntradaManualMultilateralCBLC

    strOperacoesSTR = enumTipoOperacaoLQS.EnvioTEDSTR0006Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioTEDSTR0007Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioTEDSTR0008Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioTEDSTR0009Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioTEDSTR0025Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioTEDSTR0034Clientes & ", " & _
                      enumTipoOperacaoLQS.EmissaoTEDSTR0007FdosFIDC & ", " & _
                      enumTipoOperacaoLQS.EmissaoTEDSTR0008FdosFIDC & ", " & _
                      enumTipoOperacaoLQS.EnvioSTR0006PagDespesas & ", " & _
                      enumTipoOperacaoLQS.EnvioSTR0008PagDespesas & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasBoleto & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasTributos & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib & ", " & _
                      enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib & ", " & _
                      enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta & ", " & _
                      enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib & ", "

    strOperacoesSTR = strOperacoesSTR & _
                      enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega & ", " & _
                      enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara & ", " & _
                      enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega & ", " & _
                      enumTipoOperacaoLQS.CAMInformaContratacaoInterbancarioViaLeilao & ", " & _
                      enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria

    strOperacoesPAG = enumTipoOperacaoLQS.EnvioPAG0105Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioPAG0106Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioPAG0108Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioPAG0109Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioPAG0121Clientes & ", " & _
                      enumTipoOperacaoLQS.EnvioPAG0134Clientes & ", " & _
                      enumTipoOperacaoLQS.EmissaoTEDPAG0106FdosFIDC & ", " & _
                      enumTipoOperacaoLQS.EmissaoTEDPAG0108FdosFIDC & ", " & _
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
    Call fgAppendAttribute(xmlDomFiltros, "Filtro6", "Local", enumLocalLiquidacao.CLBCAcoes)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro6", "Tipos", strOperacoesCBLC)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro7", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro7", "Local", enumLocalLiquidacao.SSTR)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro7", "Tipos", strOperacoesSTR)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro8", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro8", "Local", enumLocalLiquidacao.PAG)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro8", "Tipos", strOperacoesPAG)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro9", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro9", "Local", enumLocalLiquidacao.CONTA_CORRENTE)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro9", "Tipos", strOperacoesContaCorrente)
    
    'KIDA - CCR
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro10", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro10", "Local", enumLocalLiquidacao.CCR)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro10", "Tipos", strOperacoesCCR)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacaoTipoOperacao", "Filtro11", "", "Repeat_Filtros")
    Call fgAppendAttribute(xmlDomFiltros, "Filtro11", "Local", enumLocalLiquidacao.CAM)
    Call fgAppendAttribute(xmlDomFiltros, "Filtro11", "Tipos", strOperacoesCAM)
    
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
            With lstOperacao.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                    
                'Guarda na propriedade TAG a situação da operação |
                '                            data da úlmtima atualização |
                '                            código da mensagem SPB |
                '                            código do local de liquidação |
                '                            número do protocolo da operação LG
                .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                       objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text & "|" & _
                       objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "|" & _
                       objDomNode.selectSingleNode("NU_PRTC_OPER_LG").Text & "|" & _
                       objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text
    
                .SubItems(COL_OP_OPERACAO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
                
                If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                End If
                
                .SubItems(COL_OP_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_OP_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_OP_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                .SubItems(COL_OP_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                .SubItems(COL_OP_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_OP_TITULO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                
                .SubItems(COL_OP_CANAL_VENDA) = fgDescricaoCanalVenda(objDomNode.selectSingleNode("TP_CNAL_VEND").Text)
                
                If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                End If
                
                .SubItems(COL_OP_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                
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
                
                .SubItems(COL_OP_TIPO_LIQUIDACAO) = objDomNode.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text
                
                If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                   Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                    'Obtem a descrição da Empresa via QUERY XML
                    .SubItems(COL_OP_EMPRESA) = _
                        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                            objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                End If
                
                If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_OP_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text), "HH:MM")
                End If
                
                .SubItems(COL_OP_CONTA_CEDENTE) = objDomNode.selectSingleNode("CO_CNTA_CEDT").Text
                .SubItems(COL_OP_CONTA_CESSIONARIO) = objDomNode.selectSingleNode("CO_CNTA_CESS").Text
                
                .SubItems(COL_OP_CNPJ_CPF_COMITENTE) = objDomNode.selectSingleNode("NR_CNPJ_CPF").Text
                .SubItems(COL_OP_CODIGO_USUARIO) = objDomNode.selectSingleNode("CO_USUA_CADR_OPER").Text
                
                 'ESTÁ IMPLEMENTAÇÃO ESTÁ EM STAND-BY, AGUARDANDO PRIORIZAÇÃO PARA SER IMPLANTADA
'                'campos incluídos por solicitação dos usuário do Comex, devido projeto Sisbacen
'                .SubItems(COL_OP_NO_CLIE) = objDomNode.selectSingleNode("NO_CLIE").Text
'                .SubItems(COL_OP_CD_MOED_ISO) = objDomNode.selectSingleNode("CD_MOED_ISO").Text
'                .SubItems(COL_OP_VA_MOED_ESTR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_MOED_ESTR").Text)
'                .SubItems(COL_OP_NR_PERC_TAXA_CAMB) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("NR_PERC_TAXA_CAMB").Text)
'                .SubItems(COL_OP_CO_VEIC_LEGA) = objDomNode.selectSingleNode("CO_VEIC_LEGA").Text
'                .SubItems(COL_TO_CO_MESG_SPB_REGT_OPER) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
'                 If UCase(Trim(objDomNode.selectSingleNode("TP_OPER_CAMB").Text)) = "C" Then
'                    .SubItems(COL_OP_TP_OPER_CAMB) = "Compra"
'                 ElseIf UCase(Trim(objDomNode.selectSingleNode("TP_OPER_CAMB").Text)) = "V" Then
'                    .SubItems(COL_OP_TP_OPER_CAMB) = "Venda"
'                 End If

            End With
            
            'Acumula Total de Operações
            dblTotalOperacao = dblTotalOperacao + fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
        Next
        
        Call fgClassificarListview(Me.lstOperacao, lngIndexClassifListOper, True)
    
        With lstOperacao.ListItems.Add(, "kTotal")
            .SubItems(COL_OP_TITULO) = "Total"
            .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(dblTotalOperacao)
            .Checked = True
            .ListSubItems(COL_OP_TITULO).Bold = True
            .ListSubItems(COL_OP_VALOR).Bold = True
        End With
    End If
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Set objOperacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaOperacao", 0
    
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
    Set frmConfirmacaoOperacao = Nothing
End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = True
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

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

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - imgDummyH_MouseMove"

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = False
End Sub

Private Sub lstMensagem_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index

Exit Sub
ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstMensagem_ColumnClick"

End Sub

Private Sub lstMensagem_DblClick()

Dim strChave                                As String

Const POS_NUMERO_CTRL_IF                    As Integer = 0
Const POS_DATA_REGISTRO_MESG                As Integer = 1
Const POS_NUMERO_SEQUENCIA_OPERACAO         As Integer = 3

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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - lstMensagem_DblClick", Me.Caption
    
End Sub

Private Sub lstMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lstMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Call flSelecionaLista(enumTipoConfirmacao.MENSAGEM)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstMensagem_ItemClick"
End Sub

Private Sub lstMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.MENSAGEM
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - lstMensagem_MouseDown", Me.Caption

End Sub

Private Sub lstOperacao_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifListOper = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ColumnClick"
End Sub

Private Sub lstOperacao_DblClick()
    
On Error GoTo ErrorHandler

    If Not lstOperacao.SelectedItem Is Nothing Then
        If lstOperacao.SelectedItem.SubItems(COL_OP_OPERACAO) <> vbNullString Then
            With frmDetalheOperacao
                .SequenciaOperacao = Mid$(lstOperacao.SelectedItem.Key, 2) 'lstOperacao.SelectedItem.SubItems(COL_OP_OPERACAO)
                .Show vbModal
            End With
        End If
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_DblClick"
    
End Sub

Private Sub lstOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Item.Selected = True
    
    If Item.Key = "kTotal" Then
        Item.Checked = True
    End If
    
End Sub

Private Sub lstOperacao_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Call flSelecionaLista(enumTipoConfirmacao.operacao)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ItemClick"
End Sub

Private Sub lstOperacao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.operacao
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - lstOperacao_MouseDown", Me.Caption

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
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
        Else
            Call flCarregarListaOperacao(strFiltroXML)
            Call flCarregarListaMensagem(strFiltroXML)
        End If
    End If
    
    fgCursor
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - objFiltro_AplicarFiltro", Me.Caption

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strJanelas                              As String
Dim strResultadoProcessamento               As String

Dim blnUtilizaFiltro                        As Boolean

Const POS_STATUS                            As Integer = 0
Const POS_NUM_CONTROLE_IF                   As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1
Const POS_DATA_REGISTRO_MENSAGEM_SPB        As Integer = 1
Const POS_CODOGO_MENSAGEM_XML               As Integer = 2
Const POS_CODIGO_MENSAGEM                   As Integer = 2
Const POS_CODIGO_LOCAL_LIQUIDACAO           As Integer = 3

Dim xmlProcessamento                        As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    Call fgCursor(True)
    
    'Verifica se o filtro deve ser aplicado
    blnUtilizaFiltro = (tlbFiltro.Buttons("AplicarFiltro").value = tbrPressed)
    
    '>>> Trata a visualização de janelas (Operação e Mensagem) ----------------------------
    If tlbFiltro.Buttons("MostrarOperacao").value = tbrPressed Then
        strJanelas = strJanelas & "1"
    End If
    
    If tlbFiltro.Buttons("MostrarMensagem").value = tbrPressed Then
        strJanelas = strJanelas & "2"
    End If
    
    Call flArranjarJanelasExibicao(strJanelas)
    '>>> ----------------------------------------------------------------------------------
    
    Select Case Button.Key
        Case "DefinirFiltro"
            blnPrimeiraConsulta = False
            
            Set objFiltro = New frmFiltro
            Set objFiltro.FormOwner = Me
            objFiltro.TipoFiltro = enumTipoFiltroA8.frmConfirmacaoOperacao
            objFiltro.Show vbModal
            
        Case "refresh"
            If InStr(1, strFiltroXML, "DataIni") = 0 Or Not blnUtilizaFiltro Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
            Else
                Call flCarregarListaOperacao(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                Call flCarregarListaMensagem(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                Call flSelecionaLista(enumTipoConfirmacao.operacao)
            End If
            
        Case "Confirmar"
            strResultadoProcessamento = flConfirmar
            
            If strResultadoProcessamento <> vbNullString Then
                Call flMostrarResultado(strResultadoProcessamento)
                Call flCarregarListaOperacao(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                Call flCarregarListaMensagem(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                
                Call flSelecionaLista(enumTipoConfirmacao.operacao)
            End If
            
        Case "Rejeitar"
            Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlProcessamento.loadXML(flMontarXMLProcessamento)
        
            If xmlProcessamento.xml <> vbNullString Then
                strResultadoProcessamento = fgMIUExecutarGenerico("ProcessarEmLote", "A8LQS.clsOperacao", xmlProcessamento)
            End If
        
            Set xmlProcessamento = Nothing
        
            If strResultadoProcessamento <> vbNullString Then
                Call fgMostrarResultado(strResultadoProcessamento, "rejeitados")
                Call flCarregarListaOperacao(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                Call flCarregarListaMensagem(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                
                Call flSelecionaLista(enumTipoConfirmacao.operacao)
            End If
        
        Case "Agendamento"
            With frmAlteracaoAgendamento
                If Not lstOperacao.SelectedItem Is Nothing Then
                    If Not lstOperacao.SelectedItem.Key = "kTotal" Then
                        .TipoAgendamento = enumTipoAgendamento.operacao
                        
                        .SequenciaOperacao = Mid$(lstOperacao.SelectedItem.Key, 2) 'lstOperacao.SelectedItem.SubItems(COL_OP_OPERACAO)
                        
                        .StatusOperacao = Split(lstOperacao.SelectedItem.Tag, "|")(POS_STATUS)
                        .DHUltimaAtualizacao = Split(lstOperacao.SelectedItem.Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO)
                        .Comando = lstOperacao.SelectedItem.SubItems(COL_OP_NUMERO_COMANDO)
                        .DataOperacaoMensagem = lstOperacao.SelectedItem.SubItems(COL_OP_DATA_OPERACAO)
                        .HoraAgendamento = lstOperacao.SelectedItem.SubItems(COL_OP_HORARIO_ENVIO_MSG)
                        .CodigoMensagem = Split(lstOperacao.SelectedItem.Tag, "|")(POS_CODIGO_MENSAGEM)
                        .LocalLiquidacao = Split(lstOperacao.SelectedItem.Tag, "|")(POS_CODIGO_LOCAL_LIQUIDACAO)

                        If Trim(Split(lstOperacao.SelectedItem.Tag, "|")(POS_CODIGO_MENSAGEM)) = vbNullString Then
                            MsgBox "Operação não permite agendamento.", vbCritical, Me.Caption
                            Call fgCursor(False)
                        Else
                            .Show vbModal
                        End If
                        
                        Call flCarregarListaOperacao(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                        
                        GoTo ExitSub
                    End If
                End If
                
                If Not lstMensagem.SelectedItem Is Nothing Then
                    .TipoAgendamento = enumTipoAgendamento.MENSAGEM
                    
                    .NumeroControleIF = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_NUM_CONTROLE_IF)
                    .DTRegistroMensagemSPB = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_DATA_REGISTRO_MENSAGEM_SPB)
                    
                    .StatusMensagem = Split(lstMensagem.SelectedItem.Tag, "|")(POS_STATUS)
                    .DHUltimaAtualizacao = Split(lstMensagem.SelectedItem.Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO)
                    .Comando = lstMensagem.SelectedItem.SubItems(COL_MSG_NUMERO_COMANDO)
                    .DataOperacaoMensagem = lstMensagem.SelectedItem.SubItems(COL_MSG_DATA_MENSAGEM)
                    .HoraAgendamento = lstMensagem.SelectedItem.SubItems(COL_MSG_HORARIO_ENVIO_MSG)
                    .CodigoMensagem = lstMensagem.SelectedItem.SubItems(COL_MSG_CODIGO_MENSAGEM)
                    .CodigoMensagemXML = Split(lstMensagem.SelectedItem.Tag, "|")(POS_CODOGO_MENSAGEM_XML)
                    
                    .Show vbModal
                    
                    Call flCarregarListaMensagem(IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
                End If
            End With
        
        Case gstrSair
            Unload Me
       
    End Select
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

Private Sub flLimparLista(ByVal lstListView As ListView)
    lstListView.ListItems.Clear
End Sub

'Montar o resultado do processamento

Private Sub flMostrarResultado(ByVal pstrResultadoConfirmacao As String)

    With frmResultOperacaoLote
        .strDescricaoOperacao = " confirmados "
        .Resultado = pstrResultadoConfirmacao
        .Show vbModal
    End With

End Sub

' Carrega as propriedades necessárias a interface frmCompromissadaGenerica, através da
' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

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
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConfirmacaoOperacao", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Sub

'' Carrega a lista das mensagens passíveis de confirmação, através de interação
'' com a camada de controle de caso de uso MIU, método A8MIU.clsMensagem.
'' ObterDetalheMensagem, e preenche o listview de mensagens com as mesmas
Private Sub flCarregarListaMensagem(Optional ByVal pstrxmlFiltro As String = vbNullString)

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim objDomNode              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strRetLeitura           As String
Dim xmlDomLeitura           As MSXML2.DOMDocument40
Dim strSelecaoFiltro        As String
Dim lngCont                 As Long
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Call flLimparLista(lstMensagem)
    
    strSelecaoFiltro = enumStatusMensagem.ManualEmSer
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Verifica se o filtro foi informado...
    If pstrxmlFiltro <> vbNullString Then
        '...se sim, lê o existente e adiciona o filtro de STATUS
        Call xmlDomFiltros.loadXML(pstrxmlFiltro)
    Else
        '...se não, cria um novo apenas para o envio do filtro de STATUS
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
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
        
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            With lstMensagem.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                          objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                
                'Guarda na propriedade TAG a situação da mensagem |
                '                            data da úlmtima atualização |
                '                            Código do ID XML correspondente |
                '                            ID da Operação associada
                .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                       objDomNode.selectSingleNode("CO_TEXT_XML").Text & "|" & _
                       objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
                If objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_MENSAGEM) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                End If
                
                .SubItems(COL_MSG_CODIGO_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                .SubItems(COL_MSG_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_MSG_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                
                If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                   Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                    'Obtem a descrição da Empresa via QUERY XML
                    .SubItems(COL_MSG_EMPRESA) = _
                        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                            objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                End If
                
                If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text), "HH:MM")
                End If
            End With
        Next
    End If
    
    Call fgClassificarListview(Me.lstMensagem, lngIndexClassifListMesg, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Set objMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaMensagem", 0

End Sub

'' Oculta ou exibe as listagens de acordo com as preferências do usuário
Private Sub flArranjarJanelasExibicao(ByVal pstrJanelas As String)
    
On Error GoTo ErrorHandler

    Select Case pstrJanelas
           Case ""
                imgDummyH.Visible = False
                lstOperacao.Visible = False
                lstMensagem.Visible = False
                Call flSelecionaLista(0)
                
           Case "1"
                imgDummyH.Visible = False
                lstOperacao.Visible = True
                lstMensagem.Visible = False
                
                Call flSelecionaLista(enumTipoConfirmacao.operacao)
                
           Case "2"
                imgDummyH.Visible = False
                lstOperacao.Visible = False
                lstMensagem.Visible = True
                
                Call flSelecionaLista(enumTipoConfirmacao.MENSAGEM)
                
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

'' Marca/Desmarca todos os itens de uma listagem
Private Sub flMarcarDesmarcarTodas(ByVal lstListView As ListView, _
                                   ByVal plngTipoSelecao As enumTipoSelecao)

Dim lngLinha                                As Long

On Error GoTo ErrorHandler

    For lngLinha = 1 To lstListView.ListItems.Count
        'Tratamento específico para não desmarcar o CHECK BOX do item TOTAL, no ListView de Operação
        If lstListView.Name = "lstOperacao" Then
            If lstListView.ListItems(lngLinha).SubItems(COL_OP_OPERACAO) = vbNullString Then
                Exit For
            End If
        End If
        
        lstListView.ListItems(lngLinha).Checked = (plngTipoSelecao = enumTipoSelecao.MarcarTodas)
    Next

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMarcarDesmarcarTodas", 0

End Sub

'' Rotina necessária para o controle do agendamento.
Private Sub flSelecionaLista(ByVal intTipoVisao As enumTipoConfirmacao)

On Error Resume Next
    
    '>>> Rotina necessária para o controle de visualização da tela de agendamento
    Select Case intTipoVisao
        Case enumTipoConfirmacao.operacao
            Set lstMensagem.SelectedItem = Nothing
            lstOperacao.SetFocus
            
            If lstOperacao.ListItems.Count > 0 Then
                If lstOperacao.SelectedItem Is Nothing Then
                    lstOperacao.ListItems(1).Selected = True
                End If
            End If
        
        Case enumTipoConfirmacao.MENSAGEM
            Set lstOperacao.SelectedItem = Nothing
            lstMensagem.SetFocus
            
            If lstMensagem.ListItems.Count > 0 Then
                If lstMensagem.SelectedItem Is Nothing Then
                    lstMensagem.ListItems(1).Selected = True
                End If
            End If
            
        Case Else
            Set lstOperacao.SelectedItem = Nothing
            Set lstMensagem.SelectedItem = Nothing
            
    End Select

End Sub

'Monta string XML para processamento em lote
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemProc                             As MSXML2.DOMDocument40

Const POS_CODIGO_LAYOUT                     As Integer = 5
    
    On Error GoTo ErrorHandler
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, vbNullString, "Repeat_Processamento", vbNullString)
    
    For Each objListItem In lstOperacao.ListItems
        With objListItem
            If .Checked Then
                
                If .Key = "kTotal" Then Exit For
                
                Select Case Val(Split(.Tag, "|")(POS_CODIGO_LAYOUT))
                    Case enumTipoMensagemBUS.EnvioTEDClientes, enumTipoMensagemBUS.EnvioPagDespesas, enumTipoMensagemBUS.LancamentoContaCorrenteBG, _
                         enumTipoMensagemBUS.ConsultaOperacaoCCR, enumTipoMensagemBUS.EmissaoOperacaoCCR, enumTipoMensagemBUS.NegociacaoOperacaoCCR, _
                         enumTipoMensagemBUS.DevolucaoRecolhimentoEstornoReembolsoCCR, enumTipoMensagemBUS.ConsultaLimitesImportacaoCCR, _
                         enumTipoMensagemBUS.ContratacaoMercadoPrimario, enumTipoMensagemBUS.EdicaoContratacaoMercadoPrimario, _
                         enumTipoMensagemBUS.ConfirmacaoEdicaoContratacaoMercadoPrimario, enumTipoMensagemBUS.AlteracaoContrato, _
                         enumTipoMensagemBUS.EdicaoAlteracaoContrato, enumTipoMensagemBUS.ConfirmacaoEdicaoAlteracaoContrato, _
                         enumTipoMensagemBUS.LiquidacaoMercadoPrimario, enumTipoMensagemBUS.BaixaValorLiquidar, _
                         enumTipoMensagemBUS.RestabelecimentoBaixa, enumTipoMensagemBUS.CancelamentoValorLiquidar, _
                         enumTipoMensagemBUS.EdicaoCancelamentoValorLiquidar, enumTipoMensagemBUS.ConfirmacaoEdicaoCancelamentoValorLiquidar, _
                         enumTipoMensagemBUS.VinculacaoContratos, enumTipoMensagemBUS.AnulacaoEvento, enumTipoMensagemBUS.CorretoraRequisitaClausulasEspecificas, _
                         enumTipoMensagemBUS.IFInformaClausulasEspecificas, enumTipoMensagemBUS.ManutencaoCadastroAgenciaCentralizadoraCambio, _
                         enumTipoMensagemBUS.CredenciamentoDescredenciamentoDispostoRMCCI, enumTipoMensagemBUS.IncorporacaoContratos, _
                         enumTipoMensagemBUS.AceiteRejeicaoIncorporacaoContratos, enumTipoMensagemBUS.ConsultaContratosEmSer, _
                         enumTipoMensagemBUS.ConsultaEventosUmDia, enumTipoMensagemBUS.ConsultaDetalhamentoContratoInterbancario, _
                         enumTipoMensagemBUS.ConsultaEventosContratoMercadoPrimario, enumTipoMensagemBUS.ConsultaEventosContratoIntermediadoMercadoPrimario, _
                         enumTipoMensagemBUS.ConsultaHistoricoIncorporacoes, enumTipoMensagemBUS.ConsultaContratosIncorporacao, _
                         enumTipoMensagemBUS.ConsultaCadeiaIncorporacoesContrato, enumTipoMensagemBUS.ConsultaPosicaoCambioMoeda, _
                         enumTipoMensagemBUS.AtualizaçãoInclusãoInstrucoesPagamento, enumTipoMensagemBUS.ConsultaInstrucoesPagamento, _
                         enumTipoMensagemBUS.ComplementoInformacoesContratacaoInterbancarioViaLeilao, _
                         enumTipoMensagemBUS.IFCamaraConsultaContratosCambioMercadoInterbancario
                         
                            Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                
                            Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "DH_ULTI_ATLZ", Split(.Tag, "|")(1))
                    
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "RejeitarTED")
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", "A8LQS.clsOperacao")
                            
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_SEQU_OPER_ATIV", Mid$(.Key, 2))
                                    
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "TP_OPER", Val(Split(.Tag, "|")(POS_CODIGO_LAYOUT)))
                                    
                            Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemProc.xml)
                            
                            Set xmlItemProc = Nothing
                         
                    Case enumTipoMensagemBUS.RegistroOperacaoInterbancaria, _
                         enumTipoMensagemBUS.RegistroOperacaoArbitragem, _
                         enumTipoMensagemBUS.IFInformaLiquidacaoInterbancaria
                            
                            Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
                
                            Call fgAppendNode(xmlItemProc, vbNullString, "Grupo_ItemProc", vbNullString)
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "DH_ULTI_ATLZ", Split(.Tag, "|")(1))
                    
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Operacao", "RejeitarOperacao")
                            Call fgAppendAttribute(xmlItemProc, "Grupo_ItemProc", "Objeto", "A8LQS.clsOperacao")
                            
                            Call fgAppendNode(xmlItemProc, "Grupo_ItemProc", "NU_SEQU_OPER_ATIV", Mid$(.Key, 2))
                                    
                            Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemProc.xml)
                            
                            Set xmlItemProc = Nothing

                    Case Else
                        frmMural.Display = "Rejeição permitida apenas para operações: " & vbNewLine & vbNewLine & _
                                       " - Emissão Operação CCR;" & vbNewLine & _
                                       " - Negociação Operação CCR;" & vbNewLine & _
                                       " - Devolução RecolhimentoEstorno Reembolso CCR;" & vbNewLine & _
                                       " - Consulta Limite Importação CCR;" & vbNewLine & _
                                       " - Consulta operações CCR;" & vbNewLine & _
                                       " - Envio de TED a Clientes;" & vbNewLine & _
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
