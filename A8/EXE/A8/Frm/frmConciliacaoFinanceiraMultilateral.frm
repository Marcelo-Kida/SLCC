VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmConciliacaoFinanceiraMultilateral 
   Caption         =   "Ferramentas - Conciliação e Liquidação Financeira Multilateral"
   ClientHeight    =   7185
   ClientLeft      =   2085
   ClientTop       =   1785
   ClientWidth     =   12765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   12765
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAux 
      Caption         =   "fraAux"
      Height          =   615
      Left            =   8100
      TabIndex        =   10
      Top             =   120
      Width           =   5595
      Begin VB.Frame Frame1 
         Caption         =   "&Natureza Movimento"
         Height          =   570
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2895
         Begin VB.OptionButton optNaturezaMovimento 
            Caption         =   "Pa&gamento"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   255
            Width           =   1215
         End
         Begin VB.OptionButton optNaturezaMovimento 
            Caption         =   "&Recebimento"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   15
            Top             =   255
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame fraTipoMensagem 
         Caption         =   "&Tipo de Mensagem"
         Height          =   570
         Left            =   3060
         TabIndex        =   11
         Top             =   0
         Width           =   2295
         Begin VB.OptionButton optTipoMsg 
            Caption         =   "&Prévia"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   13
            Top             =   255
            Width           =   915
         End
         Begin VB.OptionButton optTipoMsg 
            Caption         =   "&Definitiva"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   12
            Top             =   255
            Value           =   -1  'True
            Width           =   975
         End
      End
   End
   Begin VB.ComboBox cboCodigoMensagem 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   1950
   End
   Begin VB.ListBox lstCNPJ 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox cboLocalLiquidacao 
      Height          =   315
      Left            =   4020
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1830
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3750
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   990
      Left            =   0
      TabIndex        =   9
      Top             =   6195
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   1746
      ButtonWidth     =   3334
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela "
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Concordar            "
            Key             =   "Concordar"
            Object.ToolTipText     =   "Concodar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Discordar              "
            Key             =   "Discordar"
            Object.ToolTipText     =   "Discordar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Ajustar Valor       "
            Key             =   "Ajustar"
            Object.ToolTipText     =   "Ajustar Valor da Operação"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confirmar (LDL0003)"
            Key             =   "Confirmar"
            Object.ToolTipText     =   "Confirmar (LDL0003)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Liberar                "
            Key             =   "Liberar"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rejeitar             "
            Key             =   "Rejeitar"
            Object.ToolTipText     =   "Rejeitar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagamento           "
            Key             =   "Pgto"
            Object.ToolTipText     =   "Pgto"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pgto Contingência"
            Key             =   "PgtoCont"
            Object.ToolTipText     =   "PgtoCont"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Regularizar            "
            Key             =   "Regularizar"
            Object.ToolTipText     =   "Regularizar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   1860
      Left            =   120
      TabIndex        =   8
      Top             =   4260
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   3281
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Conciliar Operação"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Número Comando"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número Comando Original"
         Object.Width           =   3792
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qtde Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Qtde a Conciliar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Qtde Conciliada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CNPJ/CNPF Comitente"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   9000
      Top             =   4800
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
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":0112
            Key             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":0564
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":0676
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":09C8
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":0D1A
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":106C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":13BE
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":16D8
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraMultilateral.frx":1B2A
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMensagem 
      Height          =   3285
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5794
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblCodigoMensagem 
      AutoSize        =   -1  'True
      Caption         =   "&Código Mensagem SPB"
      Height          =   195
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Local de Liquidação"
      Height          =   195
      Left            =   4020
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Empresa"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   120
      MousePointer    =   7  'Size N S
      Top             =   4140
      Width           =   7920
   End
End
Attribute VB_Name = "frmConciliacaoFinanceiraMultilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'' Efetua a Conciliação Financeira Multilateral das câmaras.
''
'' Montas os NET´s de Operação e Mensagem por CNPJ, batendo as operações com as
'' mensagens LDL0001, LDL0005R2, LDL0009R2, LDL0026R1, BMC0101, BMC0103
''
Option Explicit

Private lngItemCheckedMensagem              As Long
Private intTipoOperacao                     As enumTipoOperacaoLQS
Private lngPerfil                           As Long
Private strTipoBackOffice                   As String
Private strValorZero                        As String               '<-- para verificar quais mensagens estão zeradas, pois o display varia de acordo com "Regional Settings"

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
'Private xmlDomOperacao                     As MSXML2.DOMDocument40     '<- Utilizado para conciliação de DESPESAS
Private xmlOperacoes                        As MSXML2.DOMDocument40     '<- Lista de operacoes das mensagens
Private Const strFuncionalidade             As String = "frmConciliacaoFinanceiraMultilateral"
Private fblnDummyH                          As Boolean
Private lngListItemLocalLiquidacao          As Long

Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3023
Private Const STR_HORARIO_EXCEDIDO          As String = "Horário excedido"

Private Const MSG_IDX_POS_TIPO_LINHA            As Integer = 0      'O=somente operacao,   M=mensagem (com ou sem vlr de operacs)
Private Const MSG_IDX_POS_NU_CTRL_IF            As Integer = 1
Private Const MSG_IDX_POS_DH_REGT_MESG_SPB      As Integer = 2
Private Const MSG_IDX_POS_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const MSG_IDX_POS_DH_ULTI_ATLZ          As Integer = 4
Private Const MSG_IDX_POS_CNPJ                  As Integer = 5
Private Const MSG_IDX_POS_VALOR                 As Integer = 6
Private Const MSG_IDX_POS_CO_ULTI_SITU_PROC     As Integer = 7
Private Const MSG_IDX_POS_CO_MESG_SPB           As Integer = 8
Private Const MSG_IDX_POS_TP_INFO_LDL           As Integer = 9
Private Const MSG_IDX_POS_IN_OPER_DEBT_CRED     As Integer = 10
Private Const MSG_IDX_POS_CO_LOCA_LIQU          As Integer = 11
Private Const MSG_IDX_POS_LISTA_CNPJ            As Integer = 12  'lista de CNPJ´s das repeticoes (somente para mensagem mãe)

Private Const OP_IDX_POS_NU_SEQU_OPER_ATIV      As Integer = 0
Private Const OP_IDX_POS_DH_ULTI_ATLZ           As Integer = 1

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_AREA                  As Integer = 1
Private Const COL_MSG_VEICULO_LEGAL         As Integer = 2
Private Const COL_MSG_CNPJ                  As Integer = 3
Private Const COL_MSG_MOEDA                 As Integer = 4
Private Const COL_MSG_VALOR_OPERACOES       As Integer = 5
Private Const COL_MSG_VALOR_OPERACOES_2     As Integer = 6          'para guardar o valor sem formatacao
Private Const COL_MSG_VALOR_CAMARA          As Integer = 7
Private Const COL_MSG_VALOR_CAMARA_2        As Integer = 8          '
Private Const COL_MSG_VALOR_DIFERENCA       As Integer = 9
Private Const COL_MSG_VALOR_DIFERENCA_ORIG  As Integer = 10          'diferenca entre a msg e o valor original das operacoes (sem ajuste)
Private Const COL_MSG_STATUS                As Integer = 11
Private Const COL_MSG_HORA_MENSAGEM         As Integer = 12

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_CLIENTE                As Integer = 0
Private Const COL_OP_ID_TITULO              As Integer = 1
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 2
Private Const COL_OP_CV                     As Integer = 3
Private Const COL_OP_IN_ENTR_SAID_RECU_FINC As Integer = 4
Private Const COL_OP_MOEDA                  As Integer = 5
Private Const COL_OP_VALOR                  As Integer = 6
Private Const COL_OP_VALOR_ME               As Integer = 7
Private Const COL_OP_TAXA                   As Integer = 8
Private Const COL_OP_QUANTIDADE             As Integer = 9
Private Const COL_OP_PU                     As Integer = 10
Private Const COL_OP_CODIGO                 As Integer = 11
Private Const COL_OP_NUMERO_COMANDO         As Integer = 12
Private Const COL_OP_CONTRAPARTE            As Integer = 13
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 14
Private Const COL_OP_DATA_OPERACAO          As Integer = 15
Private Const COL_OP_VALOR_ORIGINAL         As Integer = 16
Private Const COL_OP_EMPRESA                As Integer = 17
Private Const COL_OP_STATUS                 As Integer = 18
Private Const COL_OP_CNPJ_CPF_COMITENTE     As Integer = 19

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngQtdeRepeticoesMensagem           As Long
Private lngQtdeCNPJSemMensagem              As Long
Private lngQtdeMensagensMae                 As Long

Private blnConsultaAtivada                  As Boolean

Private arrCNPJ()
'conteúdo das colunas (primeira dimensão, pois é um array COLUNA/LINHA)
Private Const ARR_CNPJ = 0
Private Const ARR_VALOR_MSG = 1
Private Const ARR_VALOR_OP = 2

Private blnCNPJVazio                        As Boolean

Private Const ARR_MAX_COLS = 2
Private Enum enumNaturezaMovimento
    Pagamento = 1
    Recebimento = 2
End Enum

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

Private Sub cboCodigoMensagem_Click()

    blnConsultaAtivada = False
    Call flConfiguraControles
    Call flConfigurarBotoesPorPerfil
    blnConsultaAtivada = True
    Call flMontaTela

End Sub

Private Sub cboEmpresa_Click()

    blnConsultaAtivada = False
    Call flConfigurarBotoesPorPerfil
    blnConsultaAtivada = True
    Call flMontaTela

End Sub

'' Exibe mensagem e operações ACONCILIAR na tela
Private Sub flMontaTela()

Dim lngEmpresa                              As Long
Dim lngLocalLiquidacao                      As Long
Dim strNaturezaMovimento                    As String
Dim strTipoInformacaoMensagem               As String
Dim strCNPJ                                 As String

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    If cboEmpresa.ListIndex <> -1 And cboLocalLiquidacao.ListIndex <> -1 And cboCodigoMensagem.ListIndex <> -1 And blnConsultaAtivada Then
        lngEmpresa = fgObterCodigoCombo(Me.cboEmpresa)
        lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)
        strNaturezaMovimento = IIf(optNaturezaMovimento(0).value, enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
        strTipoInformacaoMensagem = IIf(optTipoMsg(0).value, "D", "P")
        strCNPJ = fgObterCampoExtraCombo(cboEmpresa, lstCNPJ)
    
        Call flFormatarListas
        DoEvents
        
        lstOperacao.ListItems.Clear
        
        Call flCarregarListaMensagem(lngEmpresa, _
                                     lngLocalLiquidacao, _
                                     strNaturezaMovimento, _
                                     strTipoInformacaoMensagem, _
                                     strCNPJ)
        flConfigurarBotoesPorPerfil
                                     
        If lstMensagem.ListItems.Count > 0 Then
            If PerfilAcesso = AdmGeral Then
                lstMensagem.ListItems(1).Selected = False
            Else
                'Já seleciona o primeiro e mostra as mensagens
                lstMensagem.ListItems(1).Selected = True
                flMostraOperacoes
                        
            End If
        End If
        
                                     
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - flMontaTela", Me.Caption
    Exit Sub
    Resume

End Sub

Private Sub cboLocalLiquidacao_Click()

Dim lngLocalLiquidacao As Long
    
    If cboLocalLiquidacao.ListIndex <> -1 Then
        lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)
        
        'Verifica se escolheu as câmaras tratadas
        If lngLocalLiquidacao <> enumLocalLiquidacao.BMA _
            And lngLocalLiquidacao <> enumLocalLiquidacao.CETIP _
            And lngLocalLiquidacao <> enumLocalLiquidacao.BMC Then
            
            MsgBox "Câmara inválida: dever ser BMA, CETIP ou BMC.", vbExclamation
            cboLocalLiquidacao.ListIndex = lngListItemLocalLiquidacao
        Else
            If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
                optTipoMsg(1).Enabled = False
                optTipoMsg(1).value = False
                optTipoMsg(0).value = True
            Else
                optTipoMsg(1).Enabled = True
            End If
            
            lngListItemLocalLiquidacao = cboLocalLiquidacao.ListIndex
            blnConsultaAtivada = False
            Call flConfigurarBotoesPorPerfil
            Call flCarregaMensagensPorCamara
            blnConsultaAtivada = True
            Call flMontaTela
        End If
    End If
    
End Sub

Private Sub Form_Load()

#If EnableSoap = 1 Then
    Dim objControleAcesso   As MSSOAPLib30.SoapClient30
#Else
    Dim objControleAcesso   As A8MIU.clsControleAcessDado
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    ' Verifica o tipo de BackOffice do usuário
    
    flCNPJ_Reset
    
    blnConsultaAtivada = True
    
    tlbFiltro.Wrappable = False
    fraAux.BorderStyle = vbBSNone

    Set objControleAcesso = fgCriarObjetoMIU("A8MIU.clsControleAcessDado")
    strTipoBackOffice = objControleAcesso.ObterTipoBackOfficeUsuario(vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcesso = Nothing

    Set xmlOperacoes = CreateObject("MSXML2.DOMDocument.4.0")
    xmlOperacoes.loadXML ""

    strValorZero = fgVlrXml_To_Interface(0)
    
    Call fgCenterMe(Me)
    Call fgCursor(True)
    Set Me.Icon = mdiLQS.Icon
    Call flInicializar
    
    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR", , lstCNPJ, "NU_CNPJ")
    Call fgCarregarCombos(Me.cboLocalLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "SG_LOCA_LIQU")
    
    Call flFormatarListas
    flConfiguraControles
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - Form_Load", Me.Caption
    Exit Sub
    Resume

End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = True
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not fblnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If
    flPosicionaControles x, y

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = False
End Sub

Private Sub lstMensagem_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    If PerfilAcesso <> AdmGeral Then
        Call fgClassificarListview(Me.lstMensagem, ColumnHeader.Index)
        lngIndexClassifListMesg = ColumnHeader.Index
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - lstMensagem_ColumnClick", Me.Caption

End Sub

Private Sub lstMensagem_DblClick()

    If Not lstMensagem.SelectedItem Is Nothing Then
        If Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_TIPO_LINHA) = "M" Then
            With frmDetalheOperacao
                .NumeroControleIF = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_NU_CTRL_IF)
                .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_DH_REGT_MESG_SPB))
                .NumeroSequenciaRepeticao = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)
                '.SequenciaOperacao = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_CNPJ)
                .Show vbModal
            End With
        Else
            frmMural.Caption = Me.Caption
            frmMural.Display = "Esta linha apresenta valor das operações que não estão associadas a um item da mensagem."
            frmMural.Show vbModal
        End If
    End If
    
End Sub

Private Sub lstMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim lngLocal                                As Long

On Error GoTo ErrorHandler
    
    lngLocal = "0" & fgObterCodigoCombo(Me.cboLocalLiquidacao)

    If Split(Mid(Item.Key, 2), "|")(MSG_IDX_POS_TIPO_LINHA) <> "M" Then
        Item.Checked = False
    End If
    
    If PerfilAcesso = AdmGeral And Val(Split(Mid(Item.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) = 1 Then
        Item.Checked = False

    End If

    If lngLocal = enumLocalLiquidacao.BMC Then
        If PerfilAcesso = AdmArea And _
                Not fgIN(Val(Split(Mid(Item.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)), _
                            enumStatusMensagem.ConcordanciaBackoffice, _
                            enumStatusMensagem.AConciliar) Then
                            
            Item.Checked = False
        End If
    Else
        If PerfilAcesso = AdmArea And _
                Not fgIN(Val(Split(Mid(Item.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)), _
                            enumStatusMensagem.ConcordanciaBackoffice, _
                            enumStatusMensagem.ConcordanciaBackofficePrevia, _
                            enumStatusMensagem.DiscordanciaBackoffice) Then
                            
            Item.Checked = False
        End If
    End If
    
    Call flHabilitaBotoesPorSelecao

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstMensagem_ItemCheck"

End Sub

Private Sub lstMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

    'Mostra as operações desta mensagem
    flMostraOperacoes

End Sub

Private Sub lstOperacao_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    If PerfilAcesso <> AdmGeral Then
        Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
        lngIndexClassifListOper = ColumnHeader.Index
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - lstOperacao_ColumnClick", Me.Caption

End Sub

Private Sub lstOperacao_DblClick()
    
    If Not lstOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .CodigoEmpresa = lstOperacao.SelectedItem.ListSubItems(COL_OP_EMPRESA)
            .SequenciaOperacao = Split(Mid(lstOperacao.SelectedItem.Key, 2), "|")(OP_IDX_POS_NU_SEQU_OPER_ATIV)
            .Show vbModal
        End With
    End If
    
End Sub

Private Sub lstOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Item.Selected = True
    fgCursor True
    'Call flMostrarDiferencaConciliacao(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)
    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ItemCheck"

End Sub

Private Sub optNaturezaMovimento_Click(Index As Integer)

    blnConsultaAtivada = False
    Call flConfiguraControles
    blnConsultaAtivada = True
    Call flMontaTela
    
End Sub

Private Sub optTipoMsg_Click(Index As Integer)

    blnConsultaAtivada = False
    Call flConfiguraControles
    blnConsultaAtivada = True
    Call flMontaTela
    
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strRetorno                              As String
Dim strErro                                 As String
Dim strMensagemConfirmacao                  As String
Dim lngLocal                                As Long
Dim strCodigoMensagem                       As String

Dim intAcao                                 As enumAcaoConciliacao
Dim intAcaoAlternativa                      As Long

On Error GoTo ErrorHandler

    Button.Enabled = False: DoEvents
    
    fgCursor True
    lngLocal = "0" & fgObterCodigoCombo(Me.cboLocalLiquidacao)
    strCodigoMensagem = Me.cboCodigoMensagem.Text
    
    If lngLocal = 0 Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Local de Liquidação não selecionado."
        frmMural.Show vbModal
        GoTo ExitSub
    End If

    If strCodigoMensagem = vbNullString Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Mensagem não selecionada."
        frmMural.Show vbModal
        GoTo ExitSub
    End If

    intAcaoAlternativa = 0

    Select Case Button.Key
        Case "Refresh"
            Call flMontaTela                                    '<-- Recarrega as Listas (Operação e Mensagem)
        
        Case "Concordar", "Discordar", "Liberar", "Confirmar", "Rejeitar", "Pgto", "PgtoCont", "Regularizar"
            
            If lngLocal = enumLocalLiquidacao.BMC Then
                intAcao = fgDECODE(Button.Key, _
                    "Concordar", enumAcaoConciliacao.BOConcordar, _
                    "Discordar", enumAcaoConciliacao.BODiscordar, _
                    "Confirmar", enumAcaoConciliacao.BOEnviarConcordancia, _
                    "Rejeitar", enumAcaoConciliacao.AdmAreaRejeitar, _
                    "Pgto", enumAcaoConciliacao.AdmAreaLiberar, _
                    "PgtoCont", enumAcaoConciliacao.AdmAreaPagamentoContingencia, _
                    "Regularizar", IIf(strCodigoMensagem = "BMC0101", enumAcaoConciliacao.BORegularizar, enumAcaoConciliacao.AdmAreaRegularizar))
            Else
                intAcao = fgDECODE(Button.Key, _
                    "Concordar", IIf(PerfilAcesso = BackOffice, enumAcaoConciliacao.BOConcordar, enumAcaoConciliacao.AdmGeralEnviarConcordancia), _
                    "Discordar", IIf(PerfilAcesso = BackOffice, enumAcaoConciliacao.BODiscordar, enumAcaoConciliacao.AdmGeralEnviarDiscordancia), _
                    "Liberar", enumAcaoConciliacao.AdmAreaLiberar, _
                    "Rejeitar", IIf(PerfilAcesso = AdmArea, enumAcaoConciliacao.AdmAreaRejeitar, enumAcaoConciliacao.AdmGeralRejeitar), _
                    "Pgto", enumAcaoConciliacao.AdmGeralPagamento, _
                    "PgtoCont", enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                    "Regularizar", enumAcaoConciliacao.AdmGeralRegularizar)
            End If
            
            If lngQtdeMensagensMae > 1 Then
                If fgIN(intAcao, _
                    enumAcaoConciliacao.BOEnviarConcordancia, _
                    enumAcaoConciliacao.BOEnviarConcordanciaContingencia, _
                    enumAcaoConciliacao.AdmAreaLiberar, _
                    enumAcaoConciliacao.AdmAreaPagamentoContingencia, _
                    enumAcaoConciliacao.AdmAreaRegularizar, _
                    enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                    enumAcaoConciliacao.AdmGeralEnviarDiscordancia, _
                    enumAcaoConciliacao.AdmGeralPagamento, _
                    enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                    enumAcaoConciliacao.AdmGeralRegularizar) Then
                        
                    If MsgBox("ATENÇÃO" & vbCrLf & vbCrLf & "Há mais de uma mensagem sendo exibida." & vbCrLf & "Esta ação será realizada somente com a primeira mensagem." & vbCrLf & vbCrLf & "Deseja continuar?" _
                            , vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação") = vbNo Then
                        GoTo ExitSub
                    End If
                End If
            End If
            
            strRetorno = flValidarCampos(intAcao, lngLocal)
            If strRetorno <> "" Then
                frmMural.Caption = Me.Caption
                frmMural.Display = strRetorno
                frmMural.Show vbModal
                GoTo ExitSub
            End If

            strMensagemConfirmacao = ""
            strRetorno = flMontarXMLConciliacao(intAcao, strErro, strMensagemConfirmacao, intAcaoAlternativa)
                    
            If strErro <> vbNullString Then
                
                frmMural.Caption = Me.Caption
                frmMural.Display = strErro
                frmMural.Show vbModal
                GoTo ExitSub
            
            ElseIf strMensagemConfirmacao <> vbNullString Then
                
                fgCursor
                If MsgBox(strMensagemConfirmacao, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação") = vbNo Then
                    GoTo ExitSub
                End If
                If intAcaoAlternativa <> 0 Then
                    'Mensagem de Confirmacao mudou a acao escolhida inicialmente
                    intAcao = intAcaoAlternativa
                End If
                strMensagemConfirmacao = ""
                fgCursor False
            
            End If
            
            If strRetorno <> vbNullString Then
                If strMensagemConfirmacao = vbNullString Then
                    If lngLocal = enumLocalLiquidacao.BMC Then
                        strMensagemConfirmacao = fgDECODE(intAcao, _
                            enumAcaoConciliacao.AdmAreaRejeitar, "Deseja REJEITAR os NET´s dos CNPJ´s selecionados ?", _
                            enumAcaoConciliacao.BODiscordar, "Deseja DISCORDAR dos NET´s dos CNPJ´s selecionados ?", _
                            enumAcaoConciliacao.BOEnviarConcordancia, "Deseja enviar a MENSAGEM DE CONCORDÂNCIA ?", _
                            enumAcaoConciliacao.BOEnviarConcordanciaContingencia, "Deseja enviar a MENSAGEM DE CONCORDÂNCIA ?", _
                            enumAcaoConciliacao.AdmAreaLiberar, "Deseja efetuar o PAGAMENTO ?", _
                            enumAcaoConciliacao.AdmAreaPagamentoContingencia, "Deseja efetuar o PAGAMENTO EM CONTINGÊNCIA ?", _
                            enumAcaoConciliacao.AdmGeralRegularizar, "Deseja REGULARIZAR o PAGAMENTO EM CONTINGÊNCIA ?", _
                            "")
                    Else
                        strMensagemConfirmacao = fgDECODE(intAcao, _
                            enumAcaoConciliacao.AdmAreaRejeitar, "Deseja REJEITAR os NET´s dos CNPJ´s selecionados ?", _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, "Deseja enviar a MENSAGEM DE CONCORDÂNCIA ?", _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, "Deseja enviar a MENSAGEM DE DISCORDÂNCIA ?", _
                            enumAcaoConciliacao.AdmGeralPagamento, "Deseja efetuar o PAGAMENTO ?", _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, "Deseja efetuar o PAGAMENTO EM CONTINGÊNCIA ?", _
                            enumAcaoConciliacao.AdmGeralRegularizar, "Deseja REGULARIZAR o PAGAMENTO EM CONTINGÊNCIA ?", _
                            enumAcaoConciliacao.AdmGeralRejeitar, "Deseja REJEITAR os NET´s dos CNPJ´s selecionados ?", _
                            "")
                    End If
                End If
                
                If strMensagemConfirmacao <> vbNullString Then
                    fgCursor
                    If MsgBox(strMensagemConfirmacao, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação") = vbNo Then
                        GoTo ExitSub
                    End If
                End If
                
                fgCursor False
                
                strRetorno = flLiquidar(PerfilAcesso, intAcao, strRetorno)
            End If
            
            If strRetorno <> vbNullString Then
                Call flMostrarResultado(strRetorno)
            End If
            Call flMontaTela
            If strRetorno <> vbNullString Then
                Call flMarcarRejeitadosPorGradeHorario(strRetorno)
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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - tlbFiltro_ButtonClick", Me.Caption

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("Refresh"))
        Call fgCursor(False)
    End If
    
    If fgDesenv Then
        'configura perfis para teste em desenvolvimento
        If KeyCode = vbKeyB And ((Shift And vbCtrlMask) > 0) And ((Shift And vbShiftMask) > 0) Then
            PerfilAcesso = BackOffice
        ElseIf KeyCode = vbKeyA And ((Shift And vbCtrlMask) > 0) And ((Shift And vbShiftMask) > 0) Then
            PerfilAcesso = AdmArea
        End If
    End If
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    flPosicionaControles

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

'' Efetua o posicionamento dos controles
Private Sub flPosicionaControles(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)

On Error Resume Next
    
    Me.imgDummyH.Top = y + imgDummyH.Top
    
    With Me
        If imgDummyH.Top < 2000 Then
            imgDummyH.Top = 2000
        End If
        If imgDummyH.Top > (.Height - 3500) And (.Height - 3500) > 0 Then
            imgDummyH.Top = .Height - 3500
        End If
    End With
    
    With Me
        tlbFiltro.Top = .ScaleHeight - tlbFiltro.Height
        
        imgDummyH.Left = 0
        imgDummyH.Width = .ScaleWidth
        
        lstOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        lstOperacao.Height = .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 600
        lstOperacao.Width = .ScaleWidth - (lstOperacao.Left * 2)
        
        'Configuração por perfil de usuário
        If (PerfilAcesso = AdmArea) Or (PerfilAcesso = BackOffice) Then
            lstOperacao.Visible = True
            lstMensagem.Height = .imgDummyH.Top - .imgDummyH.Height - 800
        ElseIf PerfilAcesso = AdmGeral Then
            lstOperacao.Visible = False
            lstMensagem.Height = (lstOperacao.Height + lstOperacao.Top) - lstMensagem.Top
        End If
        lstMensagem.Width = .ScaleWidth - (lstMensagem.Left * 2)
    
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConciliacaoTitulos = Nothing
End Sub

'' Inicializa os controles e grids
Public Function flInicializar() As Boolean

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
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConciliacaoFinanceiraMultilateral", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
Exit Function
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Function

'' Valida o preenchimento dos campos em tela
Private Function flValidarCampos(pintAcao As enumAcaoConciliacao, plngLocalLiquidacao As Long) As String

Dim strRetorno                              As String
Dim objLI                                   As ListItem

    'Verifica se as seleções estão OK
    
    If fgIN(PerfilAcesso, _
            BackOffice, AdmArea) Then
        If fgItemsCheckedListView(lstMensagem) = 0 Then
            strRetorno = "Selecione um item para a conciliação."
        End If
        
        If plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
            If fgIN(pintAcao, enumAcaoConciliacao.BOConcordar) Then
                For Each objLI In lstMensagem.ListItems
                    If objLI.Checked Then
                        If Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) > 1 _
                            And Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)) <> enumStatusMensagem.AConciliar Then
                            strRetorno = "Esta ação só é permitida quando o status do NET estiver 'A CONCILIAR'."
                            GoTo Saida
                        End If
                    End If
                Next
            End If
            
            If fgIN(pintAcao, enumAcaoConciliacao.BODiscordar, enumAcaoConciliacao.BOEnviarConcordancia, enumAcaoConciliacao.AdmAreaRejeitar, enumAcaoConciliacao.AdmAreaLiberar) Then
                For Each objLI In lstMensagem.ListItems
                    If objLI.Checked Then
                        If Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) > 1 _
                            And Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)) <> enumStatusMensagem.ConcordanciaBackoffice Then
                            strRetorno = "Esta ação só é permitida quando o status do NET estiver 'CONCORDANCIA BACKOFICE'."
                            GoTo Saida
                        End If
                    End If
                Next
            End If
        End If
    Else        'ADM GERAL
        If fgIN(pintAcao, _
                    enumAcaoConciliacao.AdmGeralRejeitar) Then
            'Para estes itens, somentes itens de repeticao (NU_SEQU_CNTR_REPE > 1).
            If fgItemsCheckedListView(lstMensagem) = 0 Then
                strRetorno = "Selecione um item para a 'Rejeição'."
                GoTo Saida
            Else
                For Each objLI In lstMensagem.ListItems
                    If objLI.Checked Then
                        If Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) = 1 Then
                            strRetorno = "Esta operação não é permitida para a mensagem principal."
                            GoTo Saida
                        ElseIf Not fgIN(Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)), enumStatusMensagem.ConcordanciaAdmArea, enumStatusMensagem.DiscordanciaAdmArea) Then
                            strRetorno = "Só é possível 'Rejeitar' itens de mensagem nos seguintes status:" & vbCrLf & vbCrLf & "Concordância Adm. Área" & vbCrLf & "Discordância Adm. Área"
                            GoTo Saida
                        End If
                    End If
                Next
            End If
        End If
        If fgIN(pintAcao, _
                    enumAcaoConciliacao.AdmGeralPagamento, _
                    enumAcaoConciliacao.AdmGeralRegularizar, _
                    enumAcaoConciliacao.AdmGeralEnviarConcordancia) Then
            
            For Each objLI In lstMensagem.ListItems
                If Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) > 1 _
                    And Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)) <> enumStatusMensagem.ConcordanciaAdmArea Then
                    strRetorno = "Esta ação só é permitida quando o status de todos os CNPJ´s estiverem em 'Concordância Adm. Área'."
                    GoTo Saida
                End If
            Next
        End If
        If fgIN(pintAcao, _
                    enumAcaoConciliacao.AdmGeralPagamento, _
                    enumAcaoConciliacao.AdmGeralRegularizar, _
                    enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                    enumAcaoConciliacao.AdmGeralEnviarDiscordancia) And lngQtdeRepeticoesMensagem = 0 Then
            
            strRetorno = "Esta ação só é permitida após a chegada da mensagem."
            GoTo Saida
        End If
        If fgIN(pintAcao, _
                    enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                    enumAcaoConciliacao.AdmGeralEnviarDiscordancia) Then
            'se todos os itens estiverem como 'Concordancia Adm Area', não pode estas acoes
            Dim n                                   As Long
            For Each objLI In lstMensagem.ListItems
                If Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) > 1 _
                    And Val(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC)) <> enumStatusMensagem.ConcordanciaAdmArea Then
                    n = n + 1
                    GoTo Saida
                End If
            Next
            If n = 0 Then
                strRetorno = "Esta ação não é permitida quando o status de todos os CNPJ´s estiverem em 'Concordância Adm. Área'."
            End If
        End If
        If fgIN(pintAcao, _
                    enumAcaoConciliacao.AdmGeralPagamento, _
                    enumAcaoConciliacao.AdmGeralEnviarConcordancia) _
            And lngQtdeCNPJSemMensagem > 0 Then

            strRetorno = "Esta ação não é permitida quando algum CNPJ possui NET´s em operações que não estão na mensagem."
            GoTo Saida
        End If
        If fgIN(pintAcao, _
                    enumAcaoConciliacao.AdmAreaRegularizar) _
            And plngLocalLiquidacao = enumLocalLiquidacao.BMC Then

            strRetorno = "Esta ação deve ser comandada pelo Administrador de Área."
            GoTo Saida
        End If
        
        'If PerfilAcesso = AdmGeral And fgItemsCheckedListView(lstMensagem) > 1 Then
        '    strRetorno = "Selecione somente uma mensagem para a operação."
        'End If
    End If
    
Saida:
    
    flValidarCampos = strRetorno
    
End Function

'' Apresentas quaisquer erros que tenham ocorrido durante o processo conciliatório
Private Sub flApresentarErrosNegocio(ByVal pstrRetorno As String)

Dim xmlErroNegocio                          As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRetorno                              As String
Dim blnRestricaoImpeditiva                  As Boolean
Dim blnInformarJustificativa                As Boolean

    Set xmlErroNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlErroNegocio.loadXML(pstrRetorno)
    
    '1º apresenta as restrições impeditivas, se houver
    For Each objDomNode In xmlErroNegocio.selectNodes("//@EN_IMP")
        If Not blnRestricaoImpeditiva Then
            strRetorno = strRetorno & "RESTRIÇÃO IMPEDITIVA:" & vbNewLine
            
            blnRestricaoImpeditiva = True
        End If
        
        strRetorno = strRetorno & objDomNode.Text & vbNewLine & vbNewLine
    Next
    
    '2º apresenta as solicitações de justificativa, se houver
    For Each objDomNode In xmlErroNegocio.selectNodes("//@EN")
        If Not blnInformarJustificativa Then
            strRetorno = strRetorno & "INFORME JUSTIFICATIVA E COMENTÁRIO:" & vbNewLine
            
            blnInformarJustificativa = True
        End If
        
        strRetorno = strRetorno & objDomNode.Text & vbNewLine & vbNewLine
    Next
    
    Set xmlErroNegocio = Nothing
    
    If strRetorno <> vbNullString Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
    End If

End Sub

'' Colocar os nomes nas colunas dos grids
Private Sub flFormatarListas()

Dim lngLocalLiquidacao As Long
Dim strCodigoMensagem  As String
    
    lngLocalLiquidacao = fgDECODE(fgObterCodigoCombo(Me.cboLocalLiquidacao), vbNullString, 0, fgObterCodigoCombo(Me.cboLocalLiquidacao))
    strCodigoMensagem = Me.cboCodigoMensagem.Text

    If PerfilAcesso = BackOffice Or PerfilAcesso = AdmArea Then
        lstMensagem.CheckBoxes = True
    End If
    lstMensagem.AllowColumnReorder = False
    With lstMensagem.ColumnHeaders
        .Clear

        .Add , , "Conciliar", 900
        .Add , , "Área", IIf(PerfilAcesso = AdmGeral, 1200, 0)
        .Add , , "Veículo Legal", 3000
        .Add , , "CNPJ", 1940, lvwColumnRight
        .Add , , "Moeda", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocalLiquidacao), lvwColumnRight
        .Add , , "Valor Operações", 1695, lvwColumnRight
        .Add , , "Valor Operações (sem formatacao)", 0 'deixa escondido
        .Add , , "Valor Câmara", 1695, lvwColumnRight
        .Add , , "Valor Câmara (sem formatacao)", 0 'deixa escondido
        .Add , , "Diferença", 1695, lvwColumnRight
        .Add , , "Diferença Valor Original das Opers", 0 'deixa escondido
        .Add , , "Status", 2000
        .Add , , "Hora Mensagem", 2400
    
    End With
    If Not fgIN(strCodigoMensagem, "BMC0101", "BMC0103") Then
        lstMensagem.ColumnHeaders(COL_MSG_MOEDA + 1).Width = 0
    End If
    
    lstOperacao.CheckBoxes = False
    lstOperacao.AllowColumnReorder = False
    With lstOperacao.ColumnHeaders
        .Clear
    
        .Add , , "Cliente", 2700
        .Add , , "ID Ativo", 2000
        .Add , , "Data Vencimento", 1275
        .Add , , "D/C", 800
        .Add , , "Entrada/Saída em Moeda Nacional", 1000
        .Add , , "Moeda", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocalLiquidacao), lvwColumnRight
        .Add , , "Valor", 1695, lvwColumnRight
        .Add , , "Valor Moeda Estrangeira", flLarguraColunaCamara(1695, enumLocalLiquidacao.BMC, lngLocalLiquidacao), lvwColumnRight
        .Add , , "Taxa Câmbio", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocalLiquidacao), lvwColumnRight
        .Add , , "Quantidade", flLarguraColunaCamara(1440, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocalLiquidacao), lvwColumnRight
        .Add , , "PU", flLarguraColunaCamara(1344, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocalLiquidacao), lvwColumnRight
        .Add , , "Operação", 2000
        .Add , , "Número Comando", 1440
        .Add , , "Contraparte", flLarguraColunaCamara(1500, enumLocalLiquidacao.BMC, lngLocalLiquidacao)
        .Add , , "Data Liquidação", 1275
        .Add , , "Data Operação", 1275
        .Add , , "Valor Original (sem ajuste)", flLarguraColunaCamara(1695, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocalLiquidacao), lvwColumnRight
        .Add , , "Empresa", 0
        .Add , , "Situação", 2000
        .Add , , "CNPJ/CPF Comitente", 2100
        
    End With
    If PerfilAcesso <> AdmGeral Then
        'lstMensagem.ColumnHeaders(COL_MSG_AREA).Width = 0
    End If
    
    
End Sub

'' Monta um XML com os dados de operações e mensagens selecionados para a
'' conciliação
Private Function flMontarXMLConciliacao(ByVal pintAcao As enumAcaoConciliacao, _
                                        ByRef pstrErro As String, _
                                        Optional ByRef pstrMensagemConfirmacao, _
                                        Optional ByRef plngAcaoAlternativa As Long = 0) As String

'Se OK, retorna pstrErro = vbNullString

Dim xmlDom                                  As MSXML2.DOMDocument40
Dim xmlBranchOperacoes                      As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim lngLocalLiquidacao                      As Long
Dim intIgnoraGradeHorario                   As Integer
Dim objLI                                   As ListItem
Dim objLIFilha                              As ListItem
Dim strOperacoes                            As String
Dim blnAdmGeral                             As Boolean
Dim blnMensagemMae                          As Boolean
Dim strCodigoMensagem                       As String

Dim strCNPJs                                As String           'lista de cnpjs adicionados ao xml de conciliacao
Dim strCNPJ                                 As String
Dim strNU_CTRL_IF                           As String

    lngLocalLiquidacao = Val(fgObterCodigoCombo(Me.cboLocalLiquidacao))

    Set xmlDom = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlBranchOperacoes = CreateObject("MSXML2.DOMDocument.4.0")
    
    '//////////////////////////////////////////////////////////////////
    '//
    '// Para BackOffice e Adm Area manda assim:
    '//
    '// <XML>
    '//     <ITEMMENSAGEM> (sempre c/ NU_SEQU_CNTR_REPE > 1)
    '//         <OP1>
    '//         <OP2>
    '//         <OP3>
    '//     </ITEMMENSAGEM>
    '//     <ITEMMENSAGEM> (sempre c/ NU_SEQU_CNTR_REPE > 1)
    '//         <OP1>
    '//     </ITEMMENSAGEM>
    '// </XML>
    '//
    '// Para Admnistrador Geral, manda assim
    '// <XML>
    '//     <MENSAGEM_MAE> (NU_SEQU_CNTR_REPE = 1)
    '//     </MENSAGEM_MAE>
    '//     <ITEMMENSAGEM> (sempre c/ NU_SEQU_CNTR_REPE > 1)
    '//         <OP1>
    '//         <OP2>
    '//     </ITEMMENSAGEM>
    '//     <ITEMMENSAGEM> (sempre c/ NU_SEQU_CNTR_REPE > 1)
    '//         <OP1>
    '//         <OP2>
    '//     </ITEMMENSAGEM>
    '// </XML>
    '//
    '//////////////////////////////////////////////////////////////////

    blnAdmGeral = (PerfilAcesso = AdmGeral)
    
    xmlDom.loadXML ""
    
    strCNPJs = ";"
    For Each objLI In lstMensagem.ListItems
        strCodigoMensagem = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CO_MESG_SPB)
        blnMensagemMae = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE) = "1" And Left(strCodigoMensagem, 3) = "LDL"
        strCNPJ = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_CNPJ)
        strNU_CTRL_IF = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_CTRL_IF)
        If objLI.Checked Or (blnMensagemMae And pintAcao <> AdmGeralRejeitar) Then
            If xmlDom.xml = vbNullString Then
                Call fgAppendNode(xmlDom, "", "Repeat_Conciliacao", "")
            End If
            
            'Validar linha
            If pintAcao = enumAcaoConciliacao.BOConcordar Then
                If objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).Text <> strValorZero Then
                    If cboCodigoMensagem.Text = "LDL0026R1" Then
                        pstrMensagemConfirmacao = "Um ou mais itens estão com divergência de valores." & vbCrLf & vbCrLf & "Deseja continuar mesmo assim?"
                    Else
                        If strCodigoMensagem = "BMC0101" Then
                            pstrMensagemConfirmacao = "Um ou mais itens estão com divergência de valores." & vbCrLf & vbCrLf & "Deseja continuar mesmo assim?"
                            plngAcaoAlternativa = enumAcaoConciliacao.BOEnviarConcordanciaContingencia
                        Else
                            pstrErro = "Para esta mensagem, só é permitido 'Concordar' com itens que não possuam diferença de valores."
                            Exit Function
                        End If
                    End If
                End If
            ElseIf pintAcao = enumAcaoConciliacao.BODiscordar Then
                If objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).Text = strValorZero And lngLocalLiquidacao <> enumLocalLiquidacao.BMC Then
                    pstrErro = "Só é permitido 'Discordar' de itens que possuam diferença de valores."
                    Exit Function
                End If
            ElseIf pintAcao = enumAcaoConciliacao.BORegularizar Then
                If objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).Text <> strValorZero Then
                    pstrErro = "Um ou mais itens estão com divergência de valores."
                    Exit Function
                End If
            ElseIf fgIN(pintAcao, enumAcaoConciliacao.AdmAreaLiberar, enumAcaoConciliacao.AdmAreaRegularizar) Then
                If objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).Text <> strValorZero Then
                    pstrErro = "Um ou mais itens estão com divergência de valores."
                    Exit Function
                End If
            ElseIf pintAcao = enumAcaoConciliacao.AdmAreaPagamentoContingencia Then
                If objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).Text = strValorZero Then
                    pstrErro = "Não há divergência de valores. Efetue o pagamento normalmente."
                    Exit Function
                End If
            
            End If
            
            If blnMensagemMae Then
                'Mensagem mae (adm geral)
                'Insere a mae ...
                pstrErro = flMontarXMLConciliacao_InsereMensagemEOperacoes(xmlDom, objLI, False)
                If pstrErro <> vbNullString Then
                    Exit For
                End If
                '...e as filhas no XML (operacoes somente insere pras filhas)
                For Each objLIFilha In lstMensagem.ListItems
                    
                    If Split(Mid(objLIFilha.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE) <> "1" _
                    And Split(Mid(objLIFilha.Key, 2), "|")(MSG_IDX_POS_NU_CTRL_IF) = strNU_CTRL_IF Then
                        
                        If _
                                      Split(Mid(objLIFilha.Key, 2), "|")(MSG_IDX_POS_NU_CTRL_IF) _
                                    = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_NU_CTRL_IF) _
                            And _
                                      Split(Mid(objLIFilha.Key, 2), "|")(MSG_IDX_POS_DH_REGT_MESG_SPB) _
                                    = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_DH_REGT_MESG_SPB) Then
                        
                            pstrErro = flMontarXMLConciliacao_InsereMensagemEOperacoes(xmlDom, objLIFilha, True)
                            If pstrErro <> vbNullString Then
                                Exit For
                            End If
                        End If
                    End If
                Next
                'Só considera a primeira mensagem mãe
                GoTo Saida
            Else
                'Item de repeticao da LDL
                            
                If InStr(1, strCNPJs, (";" & strCNPJ & ";"), vbTextCompare) > 0 Then
                    'Não deixa adicionar o mesmo CNPJ duas vezes no XML de conciliacao
                    pstrErro = "Não é permitido selecionar o mesmo CNPJ mais de uma vez."
                    Exit Function
                End If
                strCNPJs = strCNPJs & strCNPJ & ";"

                pstrErro = flMontarXMLConciliacao_InsereMensagemEOperacoes(xmlDom, objLI, True)
                
                If pstrErro <> vbNullString Then
                    Exit For
                End If
            End If

            'intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed, 1, 0)
            'Call fgAppendNode(xmlDom, "Grupo_Operacao", "IgnoraGradeHorario", _
                                    intIgnoraGradeHorario, "Repeat_Conciliacao")
            
        End If
    Next
    
Saida:
    
    If pstrErro = vbNullString Then
        flMontarXMLConciliacao = xmlDom.xml
    Else
        flMontarXMLConciliacao = ""
    End If
    
    Set xmlDom = Nothing

End Function

'' Habilita/Desabilita botões de acordo com o perfil do usuário
Private Sub flConfigurarBotoesPorPerfil()
'Mostra somente os botões permitidos ao usuário de acordo com o seu perfil de acesso

Dim lngLocal                        As Long
Dim strCodigoMensagem               As String
   
    lngLocal = fgDECODE(fgObterCodigoCombo(Me.cboLocalLiquidacao), vbNullString, 0, fgObterCodigoCombo(Me.cboLocalLiquidacao))
    strCodigoMensagem = Me.cboCodigoMensagem.Text

    With tlbFiltro
    
        .Buttons("Concordar").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice) _
                                     Or (PerfilAcesso = enumPerfilAcesso.AdmGeral)
        .Buttons("Concordar").ToolTipText = IIf((PerfilAcesso = enumPerfilAcesso.BackOffice), "Concordar", _
                                            IIf((PerfilAcesso = enumPerfilAcesso.AdmGeral), "Enviar Mensagem de Concordância (LDL0003)", ""))
        
        .Buttons("Discordar").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice) _
                                     Or (PerfilAcesso = enumPerfilAcesso.AdmGeral)
        .Buttons("Discordar").ToolTipText = IIf((PerfilAcesso = enumPerfilAcesso.BackOffice), "Discordar", _
                                            IIf((PerfilAcesso = enumPerfilAcesso.AdmGeral), "Enviar Mensagem de Discordância (LDL0003)", ""))
        
        .Buttons("Liberar").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea) And lngLocal <> enumLocalLiquidacao.BMC
        .Buttons("Liberar").ToolTipText = "Liberar o status atual"
        
        .Buttons("Confirmar").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice And lngLocal = enumLocalLiquidacao.BMC And strCodigoMensagem = "LDL0001")
        
        .Buttons("Rejeitar").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea) Or (PerfilAcesso = enumPerfilAcesso.AdmGeral)
        .Buttons("Rejeitar").ToolTipText = "Rejeitar o status atual"
        
        .Buttons("Pgto").Visible = (PerfilAcesso = enumPerfilAcesso.AdmGeral) Or _
                                  ((PerfilAcesso = enumPerfilAcesso.AdmArea) And lngLocal = enumLocalLiquidacao.BMC)
        .Buttons("Pgto").ToolTipText = "Enviar mensagem de Pagamento"
        
        .Buttons("PgtoCont").Visible = (PerfilAcesso = enumPerfilAcesso.AdmGeral) Or _
                                      ((PerfilAcesso = enumPerfilAcesso.AdmArea) And lngLocal = enumLocalLiquidacao.BMC)
        .Buttons("PgtoCont").ToolTipText = "Enviar mensagem de Pagamento em Contingência"
        
        .Buttons("Regularizar").Visible = (PerfilAcesso = enumPerfilAcesso.AdmGeral) Or _
                                          (PerfilAcesso = enumPerfilAcesso.AdmArea And lngLocal = enumLocalLiquidacao.BMC) Or _
                                          (PerfilAcesso = enumPerfilAcesso.BackOffice And lngLocal = enumLocalLiquidacao.BMC And strCodigoMensagem = "BMC0101")
        .Buttons("Regularizar").ToolTipText = "Enviar mensagem de Regularização (se houve Pgto. em Contingência)"

        Dim i
        For i = 2 To tlbFiltro.Buttons.Count
            If tlbFiltro.Buttons(i).Style = tbrSeparator Then
                tlbFiltro.Buttons(i).Visible = _
                    tlbFiltro.Buttons(i - 1).Visible
            End If
        Next
  
        .Refresh
    
    End With

End Sub

'' Configura qual o perfil do usuário que está utilizando a tela
'' - Backoffice
'' - Administrador Backoffice
'' - Administrador Geral
Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
'Controla o perfil de acesso do usuário

    lngPerfil = pPerfil
    
    Select Case pPerfil
        Case enumPerfilAcesso.BackOffice
            fraTipoMensagem.Visible = True
            Me.Caption = "Conciliação e Liquidação Financeira Multilateral"
        Case enumPerfilAcesso.AdmArea
            fraTipoMensagem.Visible = True
            Me.Caption = "Liberação - Conciliação e Liquidação Financeira Multilateral"
        Case enumPerfilAcesso.AdmGeral
            fraTipoMensagem.Visible = False
            Me.Caption = "Liberação - Liquidação Financeira Multilateral"
    End Select
                       
    flConfigurarBotoesPorPerfil
    flConfiguraControles
    flPosicionaControles
    flCarregaMensagensPorCamara
    
'    flMontaTela
    
End Property

Property Get PerfilAcesso() As enumPerfilAcesso
'Retorna o perfil de acesso do usuário
        
    PerfilAcesso = lngPerfil
    
End Property

'' Busca operações que estão para ser conciliadas
Private Function flObterDetalheOperacoes(ByVal plngEmpresa As Long, _
                                         ByVal plngLocalLiquidacao As Long, _
                                         ByVal pstrCNPJ As String, _
                                         ByVal pstrTipoInformacaoMensagem As String, _
                                         ByVal pstrCodigoMensagem As String, _
                                         ByVal pblnExisteMensagem As Boolean) As String
    'Retorna um XMLString com as operacoes deste item de mensagem
    
    'pstrCNPJ: um CNPJ, ou lista de CNPJs separados por pipe -> "|"

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
#End If

Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strRetLeitura           As String
Dim xmlDomLeitura           As MSXML2.DOMDocument40
Dim arrCNPJ                 As Variant
Dim i                       As Long

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    'Filtro STATUS
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Select Case pstrCodigoMensagem
        Case "LDL0001"
            If plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
                Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
                If pblnExisteMensagem Then
                    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.DebitoMoedaNacionalLiquidado)
                End If
            Else
                If pstrTipoInformacaoMensagem = "P" Then
                    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
                Else
                    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.Registrada)
                    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica)
                End If
            End If
        Case "LDL0005R2"
            If plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado)
            Else
                Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.Registrada)
                Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica)
            End If
        Case "LDL0009R2"
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
        Case "LDL0026R1"
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
        Case "BMC0101"
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
            If pblnExisteMensagem Then
                Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado)
            End If
        Case "BMC0103"
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.DebitoMoedaNacionalLiquidado)

    End Select
    
    If PerfilAcesso = AdmArea Or PerfilAcesso = AdmGeral Then
        'também vê estes status
        If pstrTipoInformacaoMensagem = "P" Then
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficePrevia)
        Else
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.DiscordanciaBackoffice)
        End If
    End If
    
    If PerfilAcesso = AdmGeral Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.DiscordanciaAdmArea)
    End If
    
    'Filtro por tipo de operação ("Ver")
    Select Case pstrCodigoMensagem
        Case "LDL0001"
            If pstrTipoInformacaoMensagem = "P" Then
                'Previa
                '------
                If plngLocalLiquidacao = enumLocalLiquidacao.BMA Then
                    
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LiquidacaoOperacaoTermoBMA, "Repeat_Filtros")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LiquidacaoCompromissadaEspecificaVolta, "Repeat_Filtros")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LiquidacaoEventosJurosTituloCompro, "Repeat_Filtros")
                
                End If
            Else
                'Definitiva
                '----------
                If plngLocalLiquidacao = enumLocalLiquidacao.BMA Then
                    
                    'Todas, exceto Liquidacao Fisica
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacaoExc", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacaoExc", "Tipo", enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA, "Repeat_Filtros")
                
                ElseIf plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
                    
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LiquidacaoMultilateralBMC, "Repeat_Filtros")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC, "Repeat_Filtros")
                
                End If
            
            End If
        Case "LDL0005R2"
                If plngLocalLiquidacao = enumLocalLiquidacao.BMA Then
    
                    'Todas, exceto Liquidacao Fisica
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacaoExc", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacaoExc", "Tipo", enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA, "Repeat_Filtros")
                
                ElseIf plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
                    
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LiquidacaoMultilateralBMC, "Repeat_Filtros")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC, "Repeat_Filtros")
                
                End If
        Case "LDL0009R2"
            
            'Somente BMA
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LiquidacaoEventosJurosBMA, "Repeat_Filtros")

        Case "LDL0026R1"
            
            'Somente CETIP
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.EventosJurosSWAP, "Repeat_Filtros")
            Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.EventosJurosTERMO, "Repeat_Filtros")
            Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.EventosCETIP, "Repeat_Filtros")
            Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.NETEntradaManualMultilateralCETIP, "Repeat_Filtros")

    End Select
    
    'Não segrega BackOffice para o Administrador Geral
    If PerfilAcesso = AdmGeral Then
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_SegregaBackOffice", "Segrega", "False", "Repeat_Filtros")
    End If
    
    If fgIN(PerfilAcesso, _
                BackOffice, AdmArea) Then
            
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CNPJ", "")

            arrCNPJ = Split(pstrCNPJ, "|")
            For i = LBound(arrCNPJ) To UBound(arrCNPJ)
                Call fgAppendNode(xmlDomFiltros, "Grupo_CNPJ", "CNPJ", arrCNPJ(i))
            Next
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", plngLocalLiquidacao)

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoLiquidacao", "TipoLiqu", enumTipoLiquidacao.Multilateral)

    'somente do dia
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    If Not fgDesenv Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux) + 1, "YYYYMMDD") & "000000"))
    Else
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
    End If
    '>>> --------------------------------------------------------------------------------------------------
    'fgDumpXML xmlDomFiltros
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing

    flObterDetalheOperacoes = strRetLeitura

Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Function

'' Mostra operações na tela
Private Sub flCarregarListaOperacao(ByVal pxmlOperacao As MSXML2.DOMDocument40, _
                                    ByVal pstrCNPJ As String)

Dim objDomNode                              As MSXML2.IXMLDOMNode

    lstOperacao.ListItems.Clear
   
    For Each objDomNode In pxmlOperacao.selectNodes("//Grupo_DetalheOperacao[./CO_CNPJ_VEIC_LEGA='" & pstrCNPJ & "']")
        With lstOperacao.ListItems.Add(, _
                "k" & fgSelectSingleNode(objDomNode, "NU_SEQU_OPER_ATIV").Text & "|" & _
                      fgSelectSingleNode(objDomNode, "DH_ULTI_ATLZ").Text)
                
            .Tag = fgSelectSingleNode(objDomNode, "DH_ULTI_ATLZ").Text
            
            .Text = fgSelectSingleNode(objDomNode, "NO_VEIC_LEGA").Text
            
            .SubItems(COL_OP_CODIGO) = fgSelectSingleNode(objDomNode, "CO_OPER_ATIV").Text
            
            .SubItems(COL_OP_CV) = fgSelectSingleNode(objDomNode, "IN_OPER_DEBT_CRED").Text
            .SubItems(COL_OP_IN_ENTR_SAID_RECU_FINC) = fgSelectSingleNode(objDomNode, "IN_ENTR_SAID_RECU_FINC").Text

            .SubItems(COL_OP_ID_TITULO) = fgSelectSingleNode(objDomNode, "NU_ATIV_MERC").Text
            
            If fgSelectSingleNode(objDomNode, "DT_VENC_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_VENC_ATIV").Text)
            End If
            
            .SubItems(COL_OP_QUANTIDADE) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "QT_ATIV_MERC").Text)
            .SubItems(COL_OP_PU) = fgVlrXml_To_InterfaceDecimais(fgSelectSingleNode(objDomNode, "PU_ATIV_MERC").Text, 8)
            .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_OPER_ATIV").Text)
            .SubItems(COL_OP_VALOR_ORIGINAL) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_OPER_ATIV_ORIG").Text)
            .SubItems(COL_OP_EMPRESA) = fgSelectSingleNodeText(objDomNode, "CO_EMPR")
            
            .SubItems(COL_OP_NUMERO_COMANDO) = fgSelectSingleNode(objDomNode, "NU_COMD_OPER").Text
            
            If fgSelectSingleNode(objDomNode, "VA_OPER_ATIV").Text <> fgSelectSingleNode(objDomNode, "VA_OPER_ATIV_ORIG").Text Then
                .ListSubItems(COL_OP_VALOR).ForeColor = RGB(0, 125, 0)
            End If
            
            If fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_LIQUIDACAO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV").Text)
            End If
            If fgSelectSingleNode(objDomNode, "DT_OPER_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_OPER_ATIV").Text)
            End If
            
            .SubItems(COL_OP_MOEDA) = fgSelectSingleNodeText(objDomNode, "CO_MOED_ESTR")
            .SubItems(COL_OP_VALOR_ME) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_MOED_ESTR").Text)
            .SubItems(COL_OP_TAXA) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "PE_TAXA_NEGO").Text)
            .SubItems(COL_OP_CONTRAPARTE) = fgSelectSingleNodeText(objDomNode, "NO_CNPT")
            
            .SubItems(COL_OP_STATUS) = fgSelectSingleNodeText(objDomNode, "DE_SITU_PROC")
            .SubItems(COL_OP_CNPJ_CPF_COMITENTE) = fgSelectSingleNodeText(objDomNode, "NR_CNPJ_CPF")
            
        End With
    Next
    
    Call fgClassificarListview(Me.lstMensagem, lngIndexClassifListOper, True)
    
End Sub

'' Mostra mensagens na tela
Private Sub flCarregarListaMensagem(ByVal plngEmpresa As Long, _
                                    ByVal plngLocalLiquidacao As Long, _
                                    ByVal plngNaturezaMovimento As String, _
                                    ByVal pstrTipoInformacaoMensagem As String, _
                                    Optional ByVal pstrCNPJ As String = "")

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodeCNPJ                          As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

Dim strRetLeitura                           As String
Dim strListaCNPJ                            As String               'lista de todos os cnpjs de todas as mensagens
Dim strListaCNPJMensagem                    As String               'lista de todos os cnpjs de uma mensagem mãe
Dim strCNPJ                                 As String
Dim strCodigoMensagem                       As String
Dim i, j                                    As Long
Dim vntValor                                As Variant
Dim vntValorOriginal                        As Variant
Dim vntDiferenca                            As Variant
Dim vntDiferencaOriginal                    As Variant
Dim objLI                                   As ListItem
Dim objLSI                                  As ListSubItem
Dim blnTemMensagem                          As Boolean
Dim blnDCFinanceiro                         As Boolean              'se o campo deb/cred das operações é em relação ao financeiro
Dim blnValorME                              As Boolean              'se é pra somar o valor em moeda estrangeira
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim blnExisteMensagem                       As Boolean

On Error GoTo ErrorHandler

    strCodigoMensagem = cboCodigoMensagem.Text
    blnValorME = fgIN(strCodigoMensagem, "BMC0101", "BMC0103")

    'Na regra geral do SLCC, operações têm o IN_OPER_DEBT_CRED em relação a títulos, e não ao financeiro:
    '   Ex: Se QT_ATIV_MERC=1000, IN_OPER_DEBT_CRED=1, VA_OPER_ATIV = 60.000 , significa débido de 1000 títulos, com 'crédito' de $60000
    '
    'Mas para as operações abaixo, o campo IN_OPER_DEBT_CRED é em relação à titulos mesmo:
    '
    'enumTipoOperacaoLQS.LiquidacaoEventosJurosBMA
    'enumTipoOperacaoLQS.EventosJurosCETIP

    blnDCFinanceiro = False
    'If fgIN(strCodigoMensagem, "LDL0009R2", "LDL0005R2") Then
    '    blnDCFinanceiro = True
    'End If
        
    flCNPJ_Reset
    lngQtdeCNPJSemMensagem = 0
    lngQtdeRepeticoesMensagem = 0
    lngQtdeMensagensMae = 0

    lstMensagem.ListItems.Clear
    lstMensagem.Sorted = False
    xmlOperacoes.loadXML ""
    
    '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
        
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    If PerfilAcesso = BackOffice Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        If plngLocalLiquidacao = enumLocalLiquidacao.BMC And strCodigoMensagem = "LDL0001" Then
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackoffice)
        End If
    ElseIf PerfilAcesso = AdmArea Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        If pstrTipoInformacaoMensagem = "P" Then
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackofficePrevia)
        Else
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackoffice)
            Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.DiscordanciaBackoffice)
        End If
    End If
    
    If PerfilAcesso = AdmGeral Then
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_StatusContador", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_StatusContador", "Item1", "", "Repeat_Filtros")
        Call fgAppendAttribute(xmlDomFiltros, "Item1", "Status", enumStatusMensagem.AConciliar)
        Call fgAppendAttribute(xmlDomFiltros, "Item1", "StatusMae", "")
        Call fgAppendAttribute(xmlDomFiltros, "Item1", "NU_SEQU_CNTR_REPE", " = 1")
        Call fgAppendNode(xmlDomFiltros, "Grupo_StatusContador", "Item2", "", "Repeat_Filtros")
        Call fgAppendAttribute(xmlDomFiltros, "Item2", "Status", "")
        Call fgAppendAttribute(xmlDomFiltros, "Item2", "StatusMae", enumStatusMensagem.AConciliar)
        Call fgAppendAttribute(xmlDomFiltros, "Item2", "NU_SEQU_CNTR_REPE", " >= 2")

        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_SegregaBackOffice", "Segrega", "False", "Repeat_Filtros")
    Else
        If fgIN(strCodigoMensagem, "BMC0101", "BMC0103") Then
            'Ver somente as repetiçoes da mensagem
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SequenciaControleRepeticao", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_SequenciaControleRepeticao", "Igual", 1)
        Else
            'Ver somente as repetiçoes da mensagem
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SequenciaControleRepeticao", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_SequenciaControleRepeticao", "MaiorOuIgual", 2)
        End If
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", plngLocalLiquidacao)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Mensagem", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Mensagem", "CodMensagem", strCodigoMensagem)
    
    If strCodigoMensagem = "LDL0001" Then
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoInformacao", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoInformacao", "TipoInf", pstrTipoInformacaoMensagem)
    
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DebitoCredito", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_DebitoCredito", "DebCred", plngNaturezaMovimento)
    End If

    'somente do dia
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    If Not fgDesenv Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux) + 1, "YYYYMMDD") & "000000"))
    Else
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
    End If
   '>>> --------------------------------------------------------------------------------------------------
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_OrderBy", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_OrderBy", "Campo", "A.NU_CTRL_IF")
    Call fgAppendNode(xmlDomFiltros, "Grupo_OrderBy", "Campo", "A.DH_REGT_MESG_SPB")
    Call fgAppendNode(xmlDomFiltros, "Grupo_OrderBy", "Campo", "A.NU_SEQU_CNTR_REPE")

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strRetLeitura = objMensagem.ObterDetalheMensagemCamara(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    blnExisteMensagem = False
    
    If strRetLeitura <> vbNullString Then
        blnExisteMensagem = True
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaMensagem")
        End If
        
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            Set objLI = lstMensagem.ListItems.Add()
            With objLI
                strListaCNPJMensagem = ""
                If PerfilAcesso = AdmGeral And fgSelectSingleNode(objDomNode, "NU_SEQU_CNTR_REPE").Text = "1" Then
                    'se for mensagem mae (adm geral), todos os cnpjs da mensagem
                    strListaCNPJMensagem = ";"
                    For Each objDomNodeCNPJ In xmlDomLeitura.selectNodes("//CO_CNPJ_VEIC_LEGA[../NU_SEQU_CNTR_REPE!='1' and  ../NU_CTRL_IF='" & fgSelectSingleNode(objDomNode, "NU_CTRL_IF").Text & "' and ../DH_REGT_MESG_SPB='" & fgSelectSingleNode(objDomNode, "DH_REGT_MESG_SPB").Text & "']")
                        'adiciona os cnpjs dos itens de repeticao
                        If InStr(1, strListaCNPJMensagem, ";" & objDomNodeCNPJ.Text & ";", vbTextCompare) = 0 Then
                            
                            'somente uma vez cada cnpj
                            strListaCNPJMensagem = strListaCNPJMensagem & objDomNodeCNPJ.Text & ";"
                        End If
                        
                    Next
                    'Retira ";" do comeco e do fim da string
                    strListaCNPJMensagem = Mid(strListaCNPJMensagem, 2, Len(strListaCNPJMensagem) - 2)
                    lngQtdeMensagensMae = lngQtdeMensagensMae + 1
                Else
                    'se for item de repeticao, somente operacoes do proprio cnpj
                    strListaCNPJMensagem = fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text
                    lngQtdeRepeticoesMensagem = lngQtdeRepeticoesMensagem + 1
                End If
                
                If blnValorME Then
                    vntValor = Val(Replace(fgSelectSingleNode(objDomNode, "VA_MOED_ESTR").Text, ",", "."))
                Else
                    vntValor = Val(Replace(fgSelectSingleNode(objDomNode, "VA_FINC").Text, ",", "."))
                End If
                If plngNaturezaMovimento = enumTipoDebitoCredito.Debito Then
                    vntValor = vntValor * -1
                End If
                                
                .Key = "k" & "M" & "|" & _
                          fgSelectSingleNode(objDomNode, "NU_CTRL_IF").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "DH_REGT_MESG_SPB").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "NU_SEQU_CNTR_REPE").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "DH_ULTI_ATLZ").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text & "|" & _
                          vntValor & "|" & _
                          fgSelectSingleNode(objDomNode, "CO_ULTI_SITU_PROC").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "CO_MESG_SPB").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "TP_INFO_LDL").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "CAMPO_IN_OPER_DEBT_CRED").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "CO_LOCA_LIQU").Text & "|" & _
                          strListaCNPJMensagem
                                
                .Tag = fgSelectSingleNode(objDomNode, "DH_ULTI_ATLZ").Text
                
                If Val(fgSelectSingleNode(objDomNode, "NU_SEQU_CNTR_REPE").Text) <> 1 Then
                    .SubItems(COL_MSG_AREA) = fgSelectSingleNode(objDomNode, "DE_BKOF").Text
                End If
                
                .SubItems(COL_MSG_CNPJ) = fgFormataCnpj(fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text)
                
                .SubItems(COL_MSG_VEICULO_LEGAL) = IIf(PerfilAcesso = AdmGeral And fgSelectSingleNode(objDomNode, "NU_SEQU_CNTR_REPE").Text > 1, _
                                                "    ", _
                                                "") & _
                                                fgSelectSingleNode(objDomNode, "NO_VEIC_LEGA").Text
                                                
                .SubItems(COL_MSG_VALOR_CAMARA) = fgVlrXml_To_Interface(vntValor)
                .SubItems(COL_MSG_VALOR_CAMARA_2) = fgVlr_To_Xml(vntValor)
                
                .SubItems(COL_MSG_VALOR_OPERACOES) = ""
                .SubItems(COL_MSG_VALOR_DIFERENCA) = ""
                .SubItems(COL_MSG_VALOR_DIFERENCA_ORIG) = ""
                
                .SubItems(COL_MSG_MOEDA) = fgSelectSingleNode(objDomNode, "CO_MOED_ESTR").Text
                
                .SubItems(COL_MSG_STATUS) = fgSelectSingleNode(objDomNode, "DE_SITU_PROC").Text
                
                .SubItems(COL_MSG_HORA_MENSAGEM) = fgDtHrXML_To_Interface(fgSelectSingleNode(objDomNode, "DH_REGT_MESG_SPB").Text)
                
                strListaCNPJ = strListaCNPJ & _
                                IIf(strListaCNPJ = "", "", "|") & _
                                fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text
                                
                If Val(fgSelectSingleNode(objDomNode, "NU_SEQU_CNTR_REPE").Text) = 1 And Not blnValorME Then
                    'Mensagem mãe em negrito
                    For Each objLSI In .ListSubItems
                        objLSI.Bold = True
                    Next
                End If
                '"Corzinha" no status
                If (PerfilAcesso = enumPerfilAcesso.AdmArea _
                    And fgIN(Val(fgSelectSingleNode(objDomNode, "CO_ULTI_SITU_PROC").Text), enumStatusMensagem.ConcordanciaBackoffice, enumStatusMensagem.ConcordanciaBackofficePrevia) _
                    ) _
                Or (PerfilAcesso = enumPerfilAcesso.AdmGeral _
                    And fgIN(Val(fgSelectSingleNode(objDomNode, "CO_ULTI_SITU_PROC").Text), enumStatusMensagem.ConcordanciaAdmArea) _
                    ) Then
                    
                    .ListSubItems(COL_MSG_STATUS).ForeColor = vbBlue
                End If
            End With
            flCNPJ_Add fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text, , vntValor
        Next
    End If
    
    'Busca os valores de operacao
    strRetLeitura = flObterDetalheOperacoes(plngEmpresa, _
                                            plngLocalLiquidacao, _
                                            vbNullString, _
                                            pstrTipoInformacaoMensagem, _
                                            strCodigoMensagem, _
                                            blnExisteMensagem)
    xmlOperacoes.loadXML strRetLeitura
    
    If strRetLeitura <> vbNullString Then
    
        'Colocar também os CNPJ´s das operações no array
        For Each objDomNode In xmlOperacoes.selectNodes("//Grupo_DetalheOperacao")
            If flCNPJ_Index(fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text) = -1 Then
                flCNPJ_Add fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text
                flCNPJ_SetValue fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text, ARR_VALOR_OP, flValorOperacoes(fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text, blnDCFinanceiro, , blnValorME)
            End If
            If flCNPJ_GetValue(fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text, ARR_VALOR_OP) & vbNullString = "" Then
                flCNPJ_SetValue fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text, ARR_VALOR_OP, flValorOperacoes(fgSelectSingleNode(objDomNode, "CO_CNPJ_VEIC_LEGA").Text, blnDCFinanceiro, , blnValorME)
            End If
        Next
    
        'Coloca os CNPJs que estao faltando na lista
        For i = LBound(arrCNPJ, 2) To UBound(arrCNPJ, 2)
            blnTemMensagem = True
            Set objLI = flListItemCNPJ(arrCNPJ(ARR_CNPJ, i) & "", (False Or blnValorME))
            
            'Se mensagem de pagamento, só coloca CNPJs com saldo devedor
            'Se mensagem de recebimento, só CNPJs com saldo credor
            If ((plngNaturezaMovimento = enumTipoDebitoCredito.Debito) And flCNPJ_GetValue(arrCNPJ(ARR_CNPJ, i), ARR_VALOR_OP) < 0) _
            Or ((plngNaturezaMovimento = enumTipoDebitoCredito.Credito) And flCNPJ_GetValue(arrCNPJ(ARR_CNPJ, i), ARR_VALOR_OP) > 0) Then
                
                If objLI Is Nothing Then
                    'Adiciona um novo item
                    lngQtdeCNPJSemMensagem = lngQtdeCNPJSemMensagem + 1
                    blnTemMensagem = False
                    Set objLI = lstMensagem.ListItems.Add()
                    
                    objLI.Key = "k" & "O" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              arrCNPJ(ARR_CNPJ, i) & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              "" & "|" & _
                              arrCNPJ(ARR_CNPJ, i)
                              
                    'Cria de antemão todos os subitems
                    For j = 1 To COL_MSG_HORA_MENSAGEM
                        objLI.SubItems(j) = " "
                    Next
                              
                    objLI.SubItems(COL_MSG_CNPJ) = fgFormataCnpj(arrCNPJ(ARR_CNPJ, i))
                    objLI.SubItems(COL_MSG_AREA) = fgSelectSingleNode(xmlOperacoes, "//DE_BKOF[../CO_CNPJ_VEIC_LEGA='" & arrCNPJ(ARR_CNPJ, i) & "']").Text
                    objLI.SubItems(COL_MSG_VEICULO_LEGAL) = fgSelectSingleNode(xmlOperacoes, "//NO_VEIC_LEGA[../CO_CNPJ_VEIC_LEGA='" & arrCNPJ(ARR_CNPJ, i) & "']").Text
                
                End If
            End If
            
    '        If blnTemMensagem Then
    '            vntDiferenca = Val(Replace((Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_VALOR)), ",", ".")) - vntValor
    '            objLI.SubItems(COL_MSG_VALOR_DIFERENCA) = fgVlrXml_To_Interface(vntDiferenca)
    '            If vntDiferenca <> 0 Then
    '                objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).ForeColor = vbRed
    '            Else
    '                If PerfilAcesso = BackOffice Then
    '                    'seleciona automaticamente para concordancia
    '                    objLI.Checked = True
    '                End If
    '            End If
    '        End If
        Next
    
        'Joga os valores do operação em cada CNPJ
        For Each objLI In lstMensagem.ListItems
            strCNPJ = Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_LISTA_CNPJ)
            
                    vntValor = flValorOperacoes(strCNPJ, blnDCFinanceiro, , blnValorME)
            vntValorOriginal = flValorOperacoesOriginal(strCNPJ, blnDCFinanceiro)
            
            objLI.SubItems(COL_MSG_VALOR_OPERACOES) = fgVlrXml_To_Interface(vntValor)
            objLI.SubItems(COL_MSG_VALOR_OPERACOES_2) = fgVlr_To_Xml(vntValor)
            
            'Calcula a diferenca
            If Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_TIPO_LINHA) = "M" Then
                
                        vntDiferenca = fgVlrXml_To_Decimal(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_VALOR)) - vntValor
                vntDiferencaOriginal = fgVlrXml_To_Decimal(Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_VALOR)) - vntValorOriginal
                
                objLI.SubItems(COL_MSG_VALOR_DIFERENCA) = fgVlrXml_To_Interface(vntDiferenca)
                objLI.SubItems(COL_MSG_VALOR_DIFERENCA_ORIG) = vntDiferencaOriginal
                
                If vntDiferenca <> 0 Then
                    objLI.ListSubItems(COL_MSG_VALOR_DIFERENCA).ForeColor = vbRed
                End If
                
                If PerfilAcesso = BackOffice And Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_TIPO_LINHA) = "M" Then
                    If vntDiferenca = 0 Then
                        'seleciona automaticamente para concordancia
                        objLI.Checked = True
                    End If
                End If
            End If
        Next
        
        'deixa operações com outra cor
        For Each objLI In lstMensagem.ListItems
            If Split(Mid(objLI.Key, 2), "|")(MSG_IDX_POS_TIPO_LINHA) = "O" Then
                For Each objLSI In objLI.ListSubItems
                    objLSI.ForeColor = RGB(0, 0, 128)
                Next
            End If
        Next
    End If

    Call fgClassificarListview(Me.lstMensagem, lngIndexClassifListMesg, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Set xmlDomLeitura = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Sub

'' Existe operações na tela
Private Sub flMostraOperacoes()

Dim lngEmpresa                              As Long
Dim lngLocalLiquidacao                      As Long
Dim strNaturezaMovimento                    As String
Dim strTipoInformacaoMensagem               As String
Dim strCNPJ                                 As String
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    If Not fgIN(PerfilAcesso, BackOffice, AdmArea) Then
        Exit Sub
    End If
    
    Call fgCursor(True)
    
    If cboEmpresa.ListIndex <> -1 And cboLocalLiquidacao.ListIndex <> -1 And cboCodigoMensagem.ListIndex <> -1 Then
        lngEmpresa = fgObterCodigoCombo(Me.cboEmpresa)
        lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)
        strNaturezaMovimento = IIf(optNaturezaMovimento(0).value, enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Debito)
        strTipoInformacaoMensagem = IIf(optTipoMsg(0).value, "D", "P")
        strCNPJ = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_CNPJ)
        strCodigoMensagem = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(MSG_IDX_POS_CO_MESG_SPB)
        
        Call flCarregarListaOperacao(xmlOperacoes, strCNPJ)
                                     
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - flMontaTela", Me.Caption

End Sub

'' Dispara as rotinas de Liquidação, para a geração de mensagens de concordancia e
'' pagamento
'' LDL0003 e LDL0004
Private Function flLiquidar(ByVal plngPerfilAcesso As enumPerfilAcesso, _
                            ByVal pintAcao As enumAcaoConciliacao, _
                            pXMLConciliacao As String) As String

#If EnableSoap = 1 Then
    Dim objConciliacao      As MSSOAPLib30.SoapClient30
#Else
    Dim objConciliacao      As A8MIU.clsOperacaoMensagem
#End If

Dim strRetorno              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objConciliacao = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
    strRetorno = objConciliacao.LiquidarMultilateralFinanceira(plngPerfilAcesso, _
                                                               pintAcao, _
                                                               pXMLConciliacao, _
                                                               vntCodErro, _
                                                               vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objConciliacao = Nothing
    
    'xmlRetornoErro.loadXML strRetorno

    flLiquidar = strRetorno
    
Exit Function
ErrorHandler:

    Set objConciliacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoFinanceiraMultilateral - flLiberar", Me.Caption

End Function
'' Obtém o valor original de uma operação, antes de qualquer ajuste que tenha
'' porventura sido efetuado.
Private Function flValorOperacoesOriginal(ByVal pstrCNPJs As String, Optional ByVal pblnDCFinanceiro As Boolean = False)
    
    flValorOperacoesOriginal = flValorOperacoes(pstrCNPJs, pblnDCFinanceiro, True)

End Function
'' Obtém o valor original de uma operação em moeda estrangeira
Private Function flValorOperacoesMoedaEstrangeira(ByVal pstrCNPJs As String, Optional ByVal pblnDCFinanceiro As Boolean = False)
    
    flValorOperacoesMoedaEstrangeira = flValorOperacoes(pstrCNPJs, pblnDCFinanceiro, False, True)

End Function

'' Mostra o valor atual da operação, ajustado ou não
Private Function flValorOperacoes(ByVal pstrCNPJs As String, _
                         Optional ByVal pblnDCFinanceiro As Boolean = False, _
                         Optional ByVal pblnValorOriginal = False, _
                         Optional ByVal pblnValorMoedaEstrangeira = False)
    
'Soma valores das operacoes destes CNPJs ("soma dos creditos" - "soma dos débitos")
'Pega os valores que estão no xml de operações, e soma com a função 'sum' do XPath

Dim strExpression                   As String
Dim arr                             As Variant
Dim i                               As Long
Dim vntValor                        As Variant
Dim intDebito                       As Integer
Dim intCredito                      As Integer
Dim strCampo                        As String
    
    vntValor = 0
    arr = Split(pstrCNPJs, ";")
    
    intDebito = enumTipoDebitoCredito.Debito
    intCredito = enumTipoDebitoCredito.Credito
    
    'If pblnDCFinanceiro Then
    '    'Débito/Credito diz respeito ao financeiro, entao trata campo DC como financeiro, e nao como titulos
    '    intDebito = enumTipoDebitoCredito.Credito
    '    intCredito = enumTipoDebitoCredito.Debito
    'End If

    If pblnValorOriginal Then
        strCampo = "VA_OPER_ATIV_ORIG_VLRXML"
    ElseIf pblnValorMoedaEstrangeira Then
        strCampo = "VA_MOED_ESTR_VLRXML"
    Else
        strCampo = "VA_OPER_ATIV_VLRXML"
    End If
     
    For i = LBound(arr) To UBound(arr)
        strExpression = "sum(//" & strCampo & "[../CAMPO_IN_OPER_DEBT_CRED='" & intDebito & "' " & _
                                  " and ../CO_CNPJ_VEIC_LEGA='" & arr(i) & "' ]) - " & _
                        "sum(//" & strCampo & "[../CAMPO_IN_OPER_DEBT_CRED='" & intCredito & "' " & _
                                  " and ../CO_CNPJ_VEIC_LEGA='" & arr(i) & "' ]) "
        
        vntValor = vntValor + Val(fgFuncaoXPath(xmlOperacoes, strExpression))
    Next
    
    If pblnValorMoedaEstrangeira Then
        vntValor = vntValor * -1
    End If
    
    flValorOperacoes = vntValor

End Function

Private Sub flMostrarResultado(ByVal pstrResultadoLiberacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " processados "
        .Resultado = pstrResultadoLiberacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'' Ajuda na montagem do XML de conciliacao
Private Function flMontarXMLConciliacao_InsereMensagemEOperacoes(pxmlDom As MSXML2.DOMDocument40, _
                                                                 pobjLI As ListItem, _
                                                  Optional ByVal pblnInsereOperacoes As Boolean = True) As String

'///////////////////////////////////////////////////////////////////////
'// Insere Mensagem e Operações correspondentes no XML de conciliacao
'///////////////////////////////////////////////////////////////////////
Dim strErro                                 As String
Dim lngOperacoes                            As Long
Dim objNode                                 As MSXML2.IXMLDOMNode

Dim lngLocalLiquidacao                      As Long
Dim lngStatusOperacao                       As Long
Dim strCodigoMensagem                       As String
Dim dblValorDiferenca                       As Double
    
    lngLocalLiquidacao = Val(Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_CO_LOCA_LIQU))
    strCodigoMensagem = Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_CO_MESG_SPB)
    
    Call fgAppendNode(pxmlDom, "Repeat_Conciliacao", "Grupo_Mensagem", "")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "NU_CTRL_IF", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_NU_CTRL_IF), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "DH_REGT_MESG_SPB", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_DH_REGT_MESG_SPB), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "NU_SEQU_CNTR_REPE", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "DH_ULTI_ATLZ", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_DH_ULTI_ATLZ), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "CO_ULTI_SITU_PROC", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "CO_MESG_SPB", _
                            strCodigoMensagem, "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "TP_INFO_LDL", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_TP_INFO_LDL), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "IN_OPER_DEBT_CRED", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_IN_OPER_DEBT_CRED), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "CO_LOCA_LIQU", _
                            lngLocalLiquidacao, "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "ValorCamara", _
                            pobjLI.SubItems(COL_MSG_VALOR_CAMARA_2), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "ValorOperacoes", _
                            pobjLI.SubItems(COL_MSG_VALOR_OPERACOES_2), "Repeat_Conciliacao")
    
    Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "CNPJ", _
                            Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_CNPJ), "Repeat_Conciliacao")
                            
    If pobjLI.Text = STR_HORARIO_EXCEDIDO Then
        Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "IgnoraGradeHorario", _
                            "QualquerCoisaAqui", "Repeat_Conciliacao")
    End If
    
    'Pega Operacoes desta mensagem
    If pblnInsereOperacoes Then
        
        Call fgAppendNode(pxmlDom, "Grupo_Mensagem", "Repeat_Operacoes", "", "Repeat_Conciliacao")
        lngOperacoes = 0
        dblValorDiferenca = 0
        
        For Each objNode In xmlOperacoes.selectNodes("//Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_CNPJ_VEIC_LEGA='" & Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_CNPJ) & "']")
            
            If lngLocalLiquidacao = enumLocalLiquidacao.BMC And (strCodigoMensagem = "LDL0001" Or strCodigoMensagem = "BMC0101") Then
                
                lngStatusOperacao = Val(objNode.selectSingleNode("CO_ULTI_SITU_PROC").Text)
                
                'Correção RATS 750 - Quando status da operação estava A Conciliar e Perfil AdmArea
                '                    (Regularização de Contingência),
                '                    não adicionava as operações, não alterando o status final.
                If lngStatusOperacao = enumStatusOperacao.AConciliar Or _
                   lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice Then
                        
                    Call fgAppendXML(pxmlDom, "Repeat_Operacoes", objNode.xml, "Repeat_Conciliacao/Grupo_Mensagem[position()=last()]/Repeat_Operacoes")
                    lngOperacoes = lngOperacoes + 1
                
                Else
                
                    If strCodigoMensagem = "LDL0001" Then
                        dblValorDiferenca = dblValorDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(objNode.selectSingleNode("VA_OPER_ATIV").Text))
                    ElseIf strCodigoMensagem = "BMC0101" Then
                        dblValorDiferenca = dblValorDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(objNode.selectSingleNode("VA_MOED_ESTR").Text))
                    End If
                    
                End If
                
            Else
                
                Call fgAppendXML(pxmlDom, "Repeat_Operacoes", objNode.xml, "Repeat_Conciliacao/Grupo_Mensagem[position()=last()]/Repeat_Operacoes")
                lngOperacoes = lngOperacoes + 1
            
            End If
        
        Next
    
        If dblValorDiferenca <> 0 Then
            pxmlDom.selectSingleNode("//ValorCamara").Text = fgVlrXml_To_Decimal(fgVlr_To_Xml(pxmlDom.selectSingleNode("//ValorCamara").Text)) + dblValorDiferenca
            pxmlDom.selectSingleNode("//ValorOperacoes").Text = fgVlrXml_To_Decimal(fgVlr_To_Xml(pxmlDom.selectSingleNode("//ValorOperacoes").Text)) + dblValorDiferenca
        End If
        
    End If
    
    flMontarXMLConciliacao_InsereMensagemEOperacoes = strErro

End Function

'' Habilita/Desabilita botões de acordo com os itens selecionados
Private Function flHabilitaBotoesPorSelecao()
   
    'Habilita os botões de acordo com a seleção do usuário

End Function

Private Function flSelecaoMensagem(Optional ByVal MensagemMae As Boolean = True, _
                                   Optional ByVal MensagemFilha As Boolean = True, _
                                   Optional ByVal Status = Null) As Long
                                   
'Verifica quantas mensagens estão selecionadas, de acordo com os parametros passados

Dim lngQtde                 As Long
Dim objLI                   As ListItem

Dim lngStatus               As Long
Dim lngSeq                  As Long

'    For Each objLI In lstMensagem.ListItems
'        If objLI.Checked Then
'            lngSeq = Val(Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE))
'            lngStatus = Val(Split(Mid(pobjLI.Key, 2), "|")(MSG_IDX_POS_CO_ULTI_SITU_PROC))
'
'            If (MensagemMae And lngSeq = 1) Then
'
'            lngQtde = lngQtde + (IIf(MensagemMae And lngSeq = 1, 1, 0) _
'                              + IIf(MensagemFilha And lngSeq > 1, 1, 0) _
'                              + IIf(MensagemMae And lngSeq = 1, 1, 0))
'
'        End If
'    Next

End Function
'' Configura controles de acordo com o perfil do usuário
Private Function flConfiguraControles()

Dim strMsg                As String
Dim strTipoMensagem       As String
Dim lngNatMov             As enumNaturezaMovimento
Dim lngLocal              As Long

On Error GoTo ErrorHandler
    
    lngLocal = fgDECODE(fgObterCodigoCombo(Me.cboLocalLiquidacao), vbNullString, 0, fgObterCodigoCombo(Me.cboLocalLiquidacao))

    'blnConsultaAtivada = False
    
    strMsg = cboCodigoMensagem.Text
    'tipomsg  0 = definitiva
    'tipomsg  1 = previa
    'natureza 0 = pagamento
    'natureza 1 = recebimento
    strTipoMensagem = IIf(optTipoMsg(0).value = True, "D", "P")
    lngNatMov = IIf(optNaturezaMovimento(0).value = True, enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
    
    Select Case strMsg
        Case "LDL0001"
            If strTipoMensagem = "D" Then
                optTipoMsg(0).Enabled = True
                If lngLocal = enumLocalLiquidacao.BMC Then
                    optTipoMsg(1).Enabled = False
                Else
                    optTipoMsg(1).Enabled = True
                End If
                optNaturezaMovimento(0).Enabled = True
                optNaturezaMovimento(1).Enabled = False
            Else
                optTipoMsg(0).Enabled = True
                optTipoMsg(1).Enabled = True
                optNaturezaMovimento(0).Enabled = True
                optNaturezaMovimento(1).Enabled = True
            End If
            If PerfilAcesso <> AdmGeral Then
                fraTipoMensagem.Visible = True
            End If
        Case "LDL0005R2"
            optTipoMsg(0).Enabled = True
            optTipoMsg(1).Enabled = False
            optNaturezaMovimento(0).Enabled = False
            optNaturezaMovimento(1).Enabled = True
            fraTipoMensagem.Visible = True
        Case "LDL0009R2"
            optTipoMsg(0).Enabled = True
            optTipoMsg(1).Enabled = False
            optNaturezaMovimento(0).Enabled = False
            optNaturezaMovimento(1).Enabled = True
            fraTipoMensagem.Visible = True
        Case "LDL0026R1"
            optTipoMsg(0).Enabled = False
            optTipoMsg(1).Enabled = True
            optNaturezaMovimento(0).Enabled = False
            optNaturezaMovimento(1).Enabled = True
            fraTipoMensagem.Visible = True
        Case "BMC0101"
            optTipoMsg(0).Enabled = True
            optTipoMsg(1).Enabled = False
            optNaturezaMovimento(0).Enabled = True
            optNaturezaMovimento(1).Enabled = False
            fraTipoMensagem.Visible = False
        Case "BMC0103"
            optTipoMsg(0).Enabled = False
            optTipoMsg(1).Enabled = True
            optNaturezaMovimento(0).Enabled = False
            optNaturezaMovimento(1).Enabled = True
            fraTipoMensagem.Visible = False
    
    End Select
    
    If optTipoMsg(0).Enabled = False Then
        optTipoMsg(1).value = True
    End If
    If optTipoMsg(1).Enabled = False Then
        optTipoMsg(0).value = True
    End If
    If optNaturezaMovimento(0).Enabled = False Then
        optNaturezaMovimento(1).value = True
    End If
    If optNaturezaMovimento(1).Enabled = False Then
        optNaturezaMovimento(0).value = True
    End If
        
    If PerfilAcesso = BackOffice Then
        If strMsg = "LDL0001" And strTipoMensagem = "D" Then
            tlbFiltro.Buttons("Discordar").Enabled = True
        Else
            tlbFiltro.Buttons("Discordar").Enabled = False
        End If
    Else
        tlbFiltro.Buttons("Discordar").Enabled = True
    End If
        
    'blnConsultaAtivada = True

Exit Function
ErrorHandler:

    'blnConsultaAtivada = True

End Function
Private Function flCNPJ_Add(ByVal pstrCNPJ, _
                            Optional pvntValorOp As Variant, _
                            Optional pvntValorMsg As Variant) As Long
'Adiciona um CNPJ no array
Dim ub As Long
    
    flCNPJ_Add = -1
    
    If flCNPJ_Index(pstrCNPJ) = -1 Then
        ReDim Preserve arrCNPJ(ARR_MAX_COLS, IIf(blnCNPJVazio, 0, UBound(arrCNPJ, 2) + 1))
        ub = UBound(arrCNPJ, 2)
        arrCNPJ(ARR_CNPJ, ub) = pstrCNPJ
        If Not IsMissing(pvntValorMsg) Then
            arrCNPJ(ARR_VALOR_MSG, ub) = pvntValorMsg
        End If
        If Not IsMissing(pvntValorOp) Then
            arrCNPJ(ARR_VALOR_OP, ub) = pvntValorOp
        End If
        blnCNPJVazio = False
        flCNPJ_Add = ub
    End If

End Function
Private Function flCNPJ_Index(ByVal pstrCNPJ) As Long
    
'Procura este CNPJ no array, e retorna o índice. Se não achar, retorna -1
Dim i
    
    flCNPJ_Index = -1
    For i = LBound(arrCNPJ, 2) To UBound(arrCNPJ, 2)
        If arrCNPJ(ARR_CNPJ, i) = pstrCNPJ Then
            flCNPJ_Index = i
            Exit Function
        End If
    Next

End Function
Private Function flCNPJ_GetValue(ByVal pstrCNPJ, ByVal column)
    
Dim i
    
    i = flCNPJ_Index(pstrCNPJ)
    If i >= 0 Then
        flCNPJ_GetValue = arrCNPJ(column, i)
    End If

End Function
Private Function flCNPJ_SetValue(ByVal pstrCNPJ, ByVal column, ByVal value)
    
Dim i
    
    i = flCNPJ_Index(pstrCNPJ)
    If i >= 0 Then
        arrCNPJ(column, i) = value
        blnCNPJVazio = False
    End If

End Function
Private Function flCNPJ_Reset()
    
    ReDim arrCNPJ(ARR_MAX_COLS, 0)
    blnCNPJVazio = True

End Function

Private Function flListItemCNPJ(pstrCNPJ As String, Optional pblnPodeMae As Boolean = False) As ListItem
    'Retorna o ListItem de um determinado CNPJ
    'Se não achar, retorna Nothing

Dim li As ListItem

    Set flListItemCNPJ = Nothing

    For Each li In lstMensagem.ListItems
        If Split(Mid(li.Key, 2), "|")(MSG_IDX_POS_CNPJ) = pstrCNPJ Then
            If Val(Split(Mid(li.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) > 1 _
            Or (Val(Split(Mid(li.Key, 2), "|")(MSG_IDX_POS_NU_SEQU_CNTR_REPE)) = 1 And pblnPodeMae) Then
            
                Set flListItemCNPJ = li
                Exit Function
            End If
        End If
    Next

End Function

'' Adicionas itens a mais no filtro de operações CETIP
Function flAdicionaFiltroOperacoesCETIP(xmlDomFiltros As MSXML2.DOMDocument40)
    
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiro, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.ResgateFundoInvestimento, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.OperacaoDefinitivaCETIP, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.OperacaoRetencaoIR, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.OperacaoCessaoContratoDerivativo, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.OperAnuenciaCessaoContratoDerivativo, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.OperIntermediacaoContratoDerivativo, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.EventosJurosSWAP, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.EventosJurosTERMO, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.EventosCETIP, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao, "Repeat_Filtros")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "Tipo", enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiroCTP4001, "Repeat_Filtros")

End Function
'' Exibe com destaque itens que tenha sido rejeitados por motivo de grade de
'' horário
Private Sub flMarcarRejeitadosPorGradeHorario(ByVal strRetorno As String)

Dim objDom                                  As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    Set objDom = CreateObject("MSXML2.DOMDocument.4.0")
    objDom.loadXML strRetorno
    
    If objDom.selectNodes("//Grupo_ControleErro").length = 1 Then
        'Se só deu um erro...
        If fgSelectSingleNode(objDom, "//CodigoErro").Text = COD_ERRO_NEGOCIO_GRADE Then
            '... e o erro foi de grade de horário.
            
            With lstMensagem.ListItems(1)
                For lngCont = 1 To .ListSubItems.Count
                        .ListSubItems(lngCont).ForeColor = vbRed
                        
                Next
                .Text = STR_HORARIO_EXCEDIDO
                .ToolTipText = "Horário limite excedido. Comande a operação novamente para ignorar grade."
                .ForeColor = vbRed

            End With
        End If
    End If
    
Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0
    Exit Sub
    Resume

End Sub

Private Function flCarregaMensagensPorCamara()

Dim lngLocal                        As Long

    cboCodigoMensagem.Clear
    
    lngLocal = Val(fgObterCodigoCombo(Me.cboLocalLiquidacao))
    
    If lngLocal <> 0 Then
        
        cboCodigoMensagem.AddItem "LDL0001"
        If PerfilAcesso <> AdmGeral Then
            
            If (lngLocal <> enumLocalLiquidacao.BMC) _
                Or (lngLocal = enumLocalLiquidacao.BMC And PerfilAcesso = BackOffice) Then
                
                cboCodigoMensagem.AddItem "LDL0005R2"
            End If
            
            If lngLocal = enumLocalLiquidacao.BMC And PerfilAcesso = BackOffice Then
                cboCodigoMensagem.AddItem "BMC0101"
                cboCodigoMensagem.AddItem "BMC0103"
            End If
            
            If lngLocal = enumLocalLiquidacao.BMA Or _
                lngLocal = enumLocalLiquidacao.CETIP Then
                cboCodigoMensagem.AddItem "LDL0009R2"
            End If
            
            If lngLocal = enumLocalLiquidacao.CETIP Then
                cboCodigoMensagem.AddItem "LDL0026R1"
            End If
        End If
        'cboCodigoMensagem.ListIndex = 0
    End If

End Function
'Define a largura da coluna no grid, de acordo com a camara selecionada (exibe/esconde)
Function flLarguraColunaCamara(ByVal lngLargura, strCamaras, lngCamara)
    
    'Passar camaras permitidas separados por ';'
    
    If lngCamara = 0 Or InStr(";" & strCamaras & ";", ";" & lngCamara & ";") > 0 Then
        flLarguraColunaCamara = lngLargura
    Else
        flLarguraColunaCamara = 0
    End If
        
End Function
