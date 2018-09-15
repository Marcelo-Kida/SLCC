VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConciliacaoRegistroOperacao 
   Caption         =   "Ferramentas - Conciliação (Registro e Operação)"
   ClientHeight    =   8640
   ClientLeft      =   1575
   ClientTop       =   690
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTimer 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4230
      TabIndex        =   8
      Text            =   "10"
      Top             =   7860
      Width           =   420
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   60000
      Left            =   5040
      Top             =   7800
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   7380
      Top             =   180
      _extentx        =   3307
      _extenty        =   873
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   9480
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
            Picture         =   "frmConciliacaoRegistroOperacao.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoRegistroOperacao.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboLocalLiquidacao 
      Height          =   315
      Left            =   4620
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2490
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4350
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   8310
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   582
      ButtonWidth     =   2884
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela "
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Concordar        "
            Key             =   "Concordar"
            Object.ToolTipText     =   "Concordar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Liberar              "
            Key             =   "Liberar"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   4080
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   7197
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
      NumItems        =   11
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
         Text            =   "CNPJ/CPF Comitente"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ListView lstMensagem 
      Height          =   1605
      Left            =   120
      TabIndex        =   5
      Top             =   4980
      Width           =   10005
      _ExtentX        =   17674
      _ExtentY        =   2831
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
      NumItems        =   0
   End
   Begin MSComCtl2.UpDown udTimer 
      Height          =   315
      Left            =   4680
      TabIndex        =   9
      Top             =   7860
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtTimer"
      BuddyDispid     =   196609
      OrigLeft        =   4860
      OrigTop         =   4470
      OrigRight       =   5100
      OrigBottom      =   4815
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "Intervalo para Refresh automático da tela (em minutos) :"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   7920
      Width           =   3945
   End
   Begin VB.Label lblMensagem 
      Caption         =   "lblMensagem"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   6660
      Width           =   9975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Local de Liquidação"
      Height          =   195
      Left            =   4620
      TabIndex        =   2
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
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
      Top             =   4860
      Width           =   9900
   End
End
Attribute VB_Name = "frmConciliacaoRegistroOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Efetua o registro de operações

Option Explicit

Private lngItemCheckedMensagem              As Long
Private lngItemCheckedOperacao              As Long
Private lngPerfil                           As Long

Private lngListItemLocalLiquidacao          As Long         'ultima selecao do combo de local liquidacao
Private intControleMenuPopUp                As enumTipoConfirmacao

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlDomOperacao                      As MSXML2.DOMDocument40     '<- Utilizado para conciliação de DESPESAS
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlDominioTpNegcBMA                 As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmConciliacaoRegistroOperacao"
Private fblnDummyH                          As Boolean

Private Const STR_CAMPOS_OPCIONAIS_DIVERGENTES = "Campos opcionais estão divergentes"

Private Const POS_MSG_NUMERO_CONTROLE_IF                    As Integer = 0
Private Const POS_MSG_DATA_REGISTRO_MESG_SPB                As Integer = 1
Private Const POS_MSG_NUMERO_SEQUENCIA_CONTADOR_REPETICAO   As Integer = 2
Private Const POS_MSG_NUMERO_SEQUENCIA_OPERACAO             As Integer = 3

Private Const POS_OP_NUMERO_SEQUENCIA_OPERACAO              As Integer = 0
Private Const POS_OP_TIPO_OPERACAO                          As Integer = 1

Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3023

Private Const POS_VALOR_MESG                As Integer = 1
Private Const POS_NUMERO_COMANDO            As Integer = 2
Private Const POS_TEXTO_XML                 As Integer = 3

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_CONCILIAR              As Integer = 0
Private Const COL_OP_TIPO_OPERACAO          As Integer = 1
Private Const COL_OP_NUMERO_OPERACAO        As Integer = 2
Private Const COL_OP_CO_CNTR_SISB           As Integer = 3
Private Const COL_OP_CD_ASSO_CAMB           As Integer = 4
Private Const COL_OP_CO_PRAC                As Integer = 5
Private Const COL_OP_CO_MOED_ESTR           As Integer = 6
Private Const COL_OP_DATA_OPERACAO          As Integer = 7
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 8
Private Const COL_OP_DATA_LIQUIDACAO_ME     As Integer = 9
Private Const COL_OP_DC                     As Integer = 10
Private Const COL_OP_ID_TITULO              As Integer = 11
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 12
Private Const COL_OP_QUANTIDADE             As Integer = 13
Private Const COL_OP_PU                     As Integer = 14
Private Const COL_OP_VALOR                  As Integer = 15
Private Const COL_OP_VALOR_ME               As Integer = 16  'Moeda estrangeira
Private Const COL_OP_VEICULO_LEGAL          As Integer = 17
Private Const COL_OP_CNPJ_CONTRAPARTE       As Integer = 18
Private Const COL_OP_NOME_CONTRAPARTE       As Integer = 19
Private Const COL_OP_CONTA_CUSTODIA         As Integer = 20
Private Const COL_OP_TAXA                   As Integer = 21
Private Const COL_OP_COD_TITULAR_CUTD       As Integer = 22
Private Const COL_OP_CONTRAPARTE_CAMARA     As Integer = 23
Private Const COL_OP_CODIGO_OPERACAO_CETIP  As Integer = 24
Private Const COL_OP_DESCRICAO_ATIVO        As Integer = 25
Private Const COL_OP_MODALIDADE_LIQUIDACAO  As Integer = 26
Private Const COL_OP_VALOR_RETORNO          As Integer = 27
Private Const COL_OP_DATA_RETORNO           As Integer = 28
Private Const COL_OP_PRAZO_DIAS_RETORNO     As Integer = 29
Private Const COL_OP_ISPB_IF_CNPT           As Integer = 30
Private Const COL_OP_CO_SISB_COTR           As Integer = 31
Private Const COL_OP_CNPJ_CPF_COMITENTE     As Integer = 32

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_CONCILIAR             As Integer = 0
Private Const COL_MSG_TIPO_NEGOCIACAO       As Integer = 1
Private Const COL_MSG_NUMERO_OPERACAO       As Integer = 2
Private Const COL_MSG_CO_CNTR_SISB          As Integer = 3
Private Const COL_MSG_CD_ASSO_CAMB          As Integer = 4
Private Const COL_MSG_CO_PRAC               As Integer = 5
Private Const COL_MSG_CO_MOED_ESTR          As Integer = 6
Private Const COL_MSG_DATA_OPERACAO         As Integer = 7
Private Const COL_MSG_DATA_LIQUIDACAO       As Integer = 8
Private Const COL_MSG_DATA_LIQUIDACAO_ME    As Integer = 9      'Moeda Estrangeira
Private Const COL_MSG_DC                    As Integer = 10
Private Const COL_MSG_ID_ATIVO              As Integer = 11
Private Const COL_MSG_DATA_VENCIMENTO       As Integer = 12
Private Const COL_MSG_QUANTIDADE            As Integer = 13
Private Const COL_MSG_PU                    As Integer = 14
Private Const COL_MSG_VALOR                 As Integer = 15
Private Const COL_MSG_VALOR_ME              As Integer = 16
Private Const COL_MSG_VEICULO_LEGAL         As Integer = 17
Private Const COL_MSG_CNPJ_CONTRAPARTE      As Integer = 18
Private Const COL_MSG_NOME_CONTRAPARTE      As Integer = 19
Private Const COL_MSG_CONTA_CUSTODIA        As Integer = 20
Private Const COL_MSG_TAXA                  As Integer = 21
Private Const COL_MSG_COD_TITULAR_CUTD      As Integer = 22
Private Const COL_MSG_CONTRAPARTE_CAMARA    As Integer = 23
Private Const COL_MSG_CODIGO_OPERACAO_CETIP As Integer = 24
Private Const COL_MSG_DESCRICAO_ATIVO       As Integer = 25
Private Const COL_MSG_MODALIDADE_LIQUIDACAO As Integer = 26
Private Const COL_MSG_VALOR_RETORNO         As Integer = 27
Private Const COL_MSG_DATA_RETORNO          As Integer = 28
Private Const COL_MSG_PRAZO_DIAS_RETORNO    As Integer = 29
Private Const COL_MSG_ISPB_IF_CNPT          As Integer = 30
Private Const COL_MSG_DATA_MENSAGEM         As Integer = 31
Private Const COL_MSG_CO_SISB_COTR          As Integer = 32
Private Const COL_MSG_CO_MESG_SPB           As Integer = 33

'Indexes dos Radio Buttons para natureza de movimento
Private Const OPT_NATUREZA_ENTREGA          As Integer = 0
Private Const OPT_NATUREZA_RECEBIMENTO      As Integer = 1

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Enum enumNaturezaMovimento
    Entrega = 1
    Recebimento = 2
End Enum

'Perfis de acesso
Private blnPerfilBO       As Boolean
Private blnPerfilAdmArea  As Boolean

'Controla o timer de refresh da tela
Private intContMinutos                      As Integer

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

Private Sub cboEmpresa_Click()

    Call flMontaTela

End Sub

'' Exibe na tela as operações e mensagens passíveis de conciliação
Private Sub flMontaTela()

Dim lngEmpresa                              As Long
Dim lngLocalLiquidacao                      As Long

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    lblMensagem = ""
    If cboEmpresa.ListIndex <> -1 And cboLocalLiquidacao.ListIndex <> -1 Then
        lngEmpresa = fgObterCodigoCombo(Me.cboEmpresa)
        lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)
    
        Call flFormatarListas
        
        DoEvents
        
        Call flCarregarListaOperacao(lngEmpresa, lngLocalLiquidacao)
        If PerfilAcesso = BackOffice Then
            Call flCarregarListaMensagem(lngEmpresa, lngLocalLiquidacao)
        End If
        
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - flMontaTela", Me.Caption

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
            lngListItemLocalLiquidacao = cboLocalLiquidacao.ListIndex
            Call flMontaTela
        End If
    End If

End Sub

Private Sub Form_Click()

    flConfigurarBotoesPorPerfil ""

End Sub

Private Sub Form_Load()

#If EnableSoap = 1 Then
    Dim objMsg              As MSSOAPLib30.SoapClient30
#Else
    Dim objMsg              As A8MIU.clsMensagem
#End If
    
Dim strRetorno              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    lblMensagem = ""
    
    Call fgCenterMe(Me)
    Call fgCursor(True)
    Set Me.Icon = mdiLQS.Icon
    Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDominioTpNegcBMA = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMsg = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strRetorno = objMsg.ObterDominioSPB("TpNegcBMA", vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    xmlDominioTpNegcBMA.loadXML strRetorno
    
    Call flInicializar
    
    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    Call fgCarregarCombos(Me.cboLocalLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "SG_LOCA_LIQU")
    
    Call flFormatarListas
    lngListItemLocalLiquidacao = -1

    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - Form_Load", Me.Caption

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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - lstMensagem_ColumnClick", Me.Caption

End Sub

Private Sub lstMensagem_DblClick()

    If Not lstMensagem.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .NumeroControleIF = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_MSG_NUMERO_CONTROLE_IF)
            .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_MSG_DATA_REGISTRO_MESG_SPB))
            .NumeroSequenciaRepeticao = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_MSG_NUMERO_SEQUENCIA_CONTADOR_REPETICAO)
            .SequenciaOperacao = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_MSG_NUMERO_SEQUENCIA_OPERACAO)
            .Show vbModal
        End With
    End If
    
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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - lstOperacao_ColumnClick", Me.Caption

End Sub

Private Sub lstOperacao_DblClick()
    
    If Not lstOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Split(Mid(lstOperacao.SelectedItem.Key, 2), "|")(POS_OP_NUMERO_SEQUENCIA_OPERACAO)
            .Show vbModal
        End With
    End If
    
End Sub

Private Sub lstOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Item.Selected = True

Exit Sub
ErrorHandler:

    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ItemCheck"

End Sub

Private Sub lstOperacao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.operacao
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - lstOperacao_MouseDown", Me.Caption

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strRetorno                              As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    
    fgCursor True
    Dim intAcao As enumAcaoConciliacao
    
    lblMensagem = ""
    Select Case Button.Key
        Case "Refresh"
            Call flMontaTela                        '<-- Recarrega as Listas (Operação e Mensagem)
            
        Case "Concordar" 'BACKOFFICE
            strRetorno = flValidarCampos()
            
            If strRetorno <> "" Then
                frmMural.Caption = Me.Caption
                frmMural.Display = strRetorno
                frmMural.Show vbModal
                GoTo ExitSub
            End If
            
            intAcao = BOConcordar
                        
            strRetorno = flConciliar(intAcao)
            
            If strRetorno = vbNullString Then                   '<-- Conciliação OK
                MsgBox "Mensagem conciliada com sucesso!", vbInformation, Me.Caption
                Call flMontaTela                            '<-- Recarrega as Listas (Operação e Mensagem)
            Else
                Call flApresentarErrosNegocio(strRetorno)   '<-- Erros de Negócio encontrados (Warnings!)
                Call flMarcaCampoOpcional(strRetorno)
            End If
            
        Case "Liberar" 'ADMINISTRADOR DE AREA
            strRetorno = flValidarCampos()
            
            If strRetorno <> "" Then
                frmMural.Caption = Me.Caption
                frmMural.Display = strRetorno
                frmMural.Show vbModal
                GoTo ExitSub
            End If
            
            intAcao = AdmAreaLiberar
                        
            xmlRetornoErro.loadXML ""
            
            strRetorno = flLiberar(intAcao)
            
            If strRetorno <> vbNullString Then
                Call flMostrarResultado(strRetorno)
            End If
            Call flMontaTela
            Call flMarcarRejeitadosPorGradeHorario
            
        Case gstrSair
            Unload Me
            
    End Select
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call flMontaTela
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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    flPosicionaControles
    Exit Sub

End Sub

'' Posiciona controles de acordo com o redimensionamento da tela
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
        
        lstMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        lstMensagem.Height = .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 600
        lstMensagem.Width = .ScaleWidth - (lstMensagem.Left * 2)
        
        'Configuração por perfil de usuário
        If PerfilAcesso = BackOffice Then
            lstMensagem.Visible = True
            lstOperacao.Height = .imgDummyH.Top - .imgDummyH.Height - 800
        ElseIf PerfilAcesso = AdmArea Then
            lstMensagem.Visible = False
            lstOperacao.Height = (lstMensagem.Height + lstMensagem.Top) - lstOperacao.Top
        End If
        
        lstOperacao.Width = .ScaleWidth - (lstOperacao.Left * 2)
        lblTimer.Top = lstMensagem.Top + lstMensagem.Height + 200
        lblTimer.Left = lstMensagem.Left
        txtTimer.Top = lblTimer.Top - 50
        udTimer.Top = lblTimer.Top - 50
    
    End With
    
    lblMensagem.Left = lstMensagem.Left
    lblMensagem.Width = lstMensagem.Width
    lblMensagem.Top = lstMensagem.Top + lstMensagem.Height + 30
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConciliacaoRegistroOperacao = Nothing
End Sub

'' Inicializa os controles da tela
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
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConciliacaoRegistroOperacao", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
    Exit Function

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Function

'' Chama as rotinas de conciliação de operação e mensagem
Private Function flConciliar(ByVal pintAcao As enumAcaoConciliacao) As String

#If EnableSoap = 1 Then
    Dim objConciliacao      As MSSOAPLib30.SoapClient30
#Else
    Dim objConciliacao      As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLOperacao          As String
Dim strXMLMensagem          As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
Dim strRetorno              As String

On Error GoTo ErrorHandler

    strXMLOperacao = flMontarXMLOperacao()
    strXMLMensagem = flMontarXMLMensagem()
    
    Set objConciliacao = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
    flConciliar = objConciliacao.ConciliarCamara(enumTipoConciliacao.Registro, _
                                                 pintAcao, _
                                                 strXMLOperacao, _
                                                 strXMLMensagem, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
                                                 
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objConciliacao = Nothing
    
    Exit Function
    
ErrorHandler:

    Set objConciliacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Function

'' Chama as rotinas de liberação, de operação e mensagem
Private Function flLiberar(ByVal pintAcao As enumAcaoConciliacao) As String

#If EnableSoap = 1 Then
    Dim objConciliacao      As MSSOAPLib30.SoapClient30
#Else
    Dim objConciliacao      As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLOperacao          As String
Dim strRetorno              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    strXMLOperacao = flMontarXMLOperacao()
    
    Set objConciliacao = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
    strRetorno = objConciliacao.LiberarRegistro(strXMLOperacao, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objConciliacao = Nothing
    
    xmlRetornoErro.loadXML strRetorno

    flLiberar = strRetorno
    
Exit Function
ErrorHandler:

    Set objConciliacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoRegistroOperacao - flLiberar", Me.Caption

End Function

'' Mostra operações passíveis de conciliação
Private Sub flCarregarListaOperacao(ByVal plngEmpresa As Long, _
                                    ByVal plngLocalLiquidacao As Long)

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
#End If

Dim objDomNode              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strRetLeitura           As String
Dim xmlDomLeitura           As MSXML2.DOMDocument40
Dim datDmais1               As Date

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    lstOperacao.ListItems.Clear
    lstOperacao.Sorted = False
    
    '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    If PerfilAcesso = BackOffice Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliarRegistro)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliarAceite)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ManualEmSer)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
    ElseIf PerfilAcesso = AdmArea Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.Concordancia)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAutomatica)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAceite)
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAceiteAutomatica)
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", plngLocalLiquidacao)
    
    Select Case plngLocalLiquidacao
        Case enumLocalLiquidacao.BMA
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_LayoutEntrada", "Layout", enumTipoMensagemLQS.RegistroOperacaoBMA)
        Case enumLocalLiquidacao.CETIP
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_OperacaoCetipConciliacao", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_OperacaoCetipConciliacao", "SoConciliacaoCetip", "sim")
        Case enumLocalLiquidacao.BMC
            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
            Call fgAppendNode(xmlDomFiltros, "Grupo_LayoutEntrada", "Layout", enumTipoMensagemLQS.RegistroOperacoesBMC)
            Call fgAppendNode(xmlDomFiltros, "Grupo_LayoutEntrada", "Layout", enumTipoMensagemLQS.RegistroOperacoesRodaDolar)
            Call fgAppendNode(xmlDomFiltros, "Grupo_LayoutEntrada", "Layout", enumTipoMensagemLQS.RegistroOperacaoInterbancaria)
    End Select

    'somente do dia
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    
    datDmais1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, proximo)
    
    If Not fgDesenv Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(Format(datDmais1, "YYYYMMDD") & "000000"))
    Else
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
    End If
    
    '>>> --------------------------------------------------------------------------------------------------
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
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
            If fgSelectSingleNode(objDomNode, "TP_MESG_RECB_INTE").Text <> enumTipoMensagemBUS.RegistroOperacaoInterbancaria _
            Or (fgSelectSingleNode(objDomNode, "CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ManualEmSer _
            And fgSelectSingleNode(objDomNode, "CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Concordancia) Then
                With lstOperacao.ListItems.Add(, _
                        "k" & fgSelectSingleNode(objDomNode, "NU_SEQU_OPER_ATIV").Text & "|" & _
                              fgSelectSingleNode(objDomNode, "TP_OPER").Text)
                        
                    .Tag = fgSelectSingleNode(objDomNode, "DH_ULTI_ATLZ").Text
                    
                    .SubItems(COL_OP_NUMERO_OPERACAO) = fgSelectSingleNode(objDomNode, "NU_COMD_OPER").Text
                    
                    If fgSelectSingleNode(objDomNode, "DT_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_OPER_ATIV").Text)
                    End If
                    If fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_LIQUIDACAO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV").Text)
                    End If
                    
                    .SubItems(COL_OP_DC) = fgSelectSingleNode(objDomNode, "IN_OPER_DEBT_CRED").Text
                    .SubItems(COL_OP_ID_TITULO) = fgSelectSingleNode(objDomNode, "NU_ATIV_MERC").Text
                    '.SubItems(COL_OP_DESCRICAO_TITULO) = fgSelectSingleNode(objDomNode, "DE_ATIV_MERC").Text
                    
                    If fgSelectSingleNode(objDomNode, "DT_VENC_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_VENC_ATIV").Text)
                    End If
                    
                    .SubItems(COL_OP_QUANTIDADE) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "QT_ATIV_MERC").Text)
                    .SubItems(COL_OP_PU) = fgVlrXml_To_InterfaceDecimais(fgSelectSingleNode(objDomNode, "PU_ATIV_MERC").Text, 8)
                    .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_OPER_ATIV").Text)
                    
                    .SubItems(COL_OP_VEICULO_LEGAL) = fgSelectSingleNode(objDomNode, "NO_VEIC_LEGA").Text
                    .SubItems(COL_OP_CNPJ_CONTRAPARTE) = fgFormataCnpj(fgSelectSingleNode(objDomNode, "CO_CNPJ_CNPT").Text)
                    .SubItems(COL_OP_NOME_CONTRAPARTE) = fgSelectSingleNode(objDomNode, "NO_CNPT").Text
                    .SubItems(COL_OP_TIPO_OPERACAO) = fgSelectSingleNode(objDomNode, "NO_TIPO_OPER").Text
                    
                    .SubItems(COL_OP_CONTA_CUSTODIA) = fgSelectSingleNode(objDomNode, "CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                    .SubItems(COL_OP_TAXA) = fgVlrXml_To_InterfaceDecimais(fgSelectSingleNode(objDomNode, "PE_TAXA_NEGO").Text, 6)
                    .SubItems(COL_OP_COD_TITULAR_CUTD) = fgSelectSingleNode(objDomNode, "CO_TITL_CUTD").Text
    
                    .SubItems(COL_OP_CONTRAPARTE_CAMARA) = fgSelectSingleNode(objDomNode, "CO_CNTA_CUTD_SELIC_CNPT").Text
                    .SubItems(COL_OP_CODIGO_OPERACAO_CETIP) = fgSelectSingleNode(objDomNode, "CO_OPER_CETIP").Text
                    .SubItems(COL_OP_DESCRICAO_ATIVO) = fgSelectSingleNode(objDomNode, "DE_ATIV_MERC").Text
                    .SubItems(COL_OP_MODALIDADE_LIQUIDACAO) = fgSelectSingleNode(objDomNode, "NO_TIPO_LIQU_OPER_ATIV").Text
                    
                    If fgSelectSingleNode(objDomNode, "DT_OPER_ATIV_RETN").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_RETORNO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_OPER_ATIV_RETN").Text)
                    End If
                    .SubItems(COL_OP_PRAZO_DIAS_RETORNO) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "PZ_DIAS_RETN_OPER_ATIV").Text, False)
                    .SubItems(COL_OP_VALOR_RETORNO) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_OPER_ATIV_RETN").Text)
                    
                    'BMC
                    .SubItems(COL_OP_CD_ASSO_CAMB) = fgSelectSingleNode(objDomNode, "CD_ASSO_CAMB").Text
                    .SubItems(COL_OP_CO_CNTR_SISB) = fgSelectSingleNode(objDomNode, "CO_CNTR_SISB").Text
                    .SubItems(COL_OP_VALOR_ME) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_MOED_ESTR").Text)
                    .SubItems(COL_OP_CO_MOED_ESTR) = fgSelectSingleNode(objDomNode, "CO_MOED_ESTR").Text
                    .SubItems(COL_OP_CO_PRAC) = fgSelectSingleNode(objDomNode, "CO_PRAC").Text
                    If fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV_MOED_ESTR").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_LIQUIDACAO_ME) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV_MOED_ESTR").Text)
                    End If
                    .SubItems(COL_OP_ISPB_IF_CNPT) = fgSelectSingleNode(objDomNode, "CO_ISPB_IF_CNPT").Text
                    .SubItems(COL_OP_CO_SISB_COTR) = fgSelectSingleNode(objDomNode, "CO_SISB_COTR").Text
                    
                    .SubItems(COL_OP_CNPJ_CPF_COMITENTE) = fgSelectSingleNode(objDomNode, "NR_CNPJ_CPF").Text
    
                End With
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lstOperacao, lngIndexClassifListOper, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Set objOperacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Sub

'' Mostra mensagens passíveis de conciliação
Private Sub flCarregarListaMensagem(ByVal plngEmpresa As Long, _
                                    ByVal plngLocalLiquidacao As Long)

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim objDomNode              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strRetLeitura           As String
Dim xmlDomLeitura           As MSXML2.DOMDocument40
Dim datDmais1               As Date
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    lstMensagem.ListItems.Clear
    lstMensagem.Sorted = False
    
    '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
        
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    If PerfilAcesso = BackOffice Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    ElseIf PerfilAcesso = AdmArea Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.Conciliada)
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SequenciaControleRepeticao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_SequenciaControleRepeticao", "Igual", 1)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Mensagem", "")
    If plngLocalLiquidacao = enumLocalLiquidacao.BMA Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Mensagem", "CodMensagem", "BMA0002")
    ElseIf plngLocalLiquidacao = enumLocalLiquidacao.CETIP Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Mensagem", "CodMensagem", "CTP1002")
    ElseIf plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Mensagem", "CodMensagem", "BMC0005")
        Call fgAppendNode(xmlDomFiltros, "Grupo_Mensagem", "CodMensagem", "BMC0011")
    End If

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", plngLocalLiquidacao)
    
    datDmais1 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 1, proximo)
    
    'somente do dia
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    If Not fgDesenv Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(Format(datDmais1, "YYYYMMDD") & "000000"))
    Else
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
    End If
   '>>> --------------------------------------------------------------------------------------------------
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strRetLeitura = objMensagem.ObterDetalheMensagemCamara(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
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
                    "k" & fgSelectSingleNode(objDomNode, "NU_CTRL_IF").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "DH_REGT_MESG_SPB").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "NU_SEQU_CNTR_REPE").Text & "|" & _
                          fgSelectSingleNode(objDomNode, "NU_SEQU_OPER_ATIV").Text)
                    
                .Tag = fgSelectSingleNode(objDomNode, "DH_ULTI_ATLZ").Text
                
                If plngLocalLiquidacao = enumLocalLiquidacao.BMA Then
                    .SubItems(COL_MSG_TIPO_NEGOCIACAO) = fgSelectSingleNode(objDomNode, "TP_NEGO_BMA").Text & " - " & flDescricaoTpNegcBMA(fgSelectSingleNode(objDomNode, "TP_NEGO_BMA").Text)  '<<----<<
                End If
                
                .SubItems(COL_MSG_NUMERO_OPERACAO) = fgSelectSingleNode(objDomNode, "NU_COMD_OPER").Text
                If fgSelectSingleNode(objDomNode, "DT_OPER").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_OPERACAO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_OPER").Text)
                End If
                If fgSelectSingleNode(objDomNode, "DT_LIQU").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_LIQUIDACAO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_LIQU").Text)
                End If
                .SubItems(COL_MSG_DC) = fgSelectSingleNode(objDomNode, "IN_OPER_DEBT_CRED").Text
                .SubItems(COL_MSG_ID_ATIVO) = fgSelectSingleNode(objDomNode, "NU_ATIV_MERC").Text
                If fgSelectSingleNode(objDomNode, "DT_VENC").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_VENCIMENTO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_VENC").Text)
                End If
                .SubItems(COL_MSG_QUANTIDADE) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "QT_ATIV_MERC").Text)
                .SubItems(COL_MSG_PU) = fgVlrXml_To_InterfaceDecimais(fgSelectSingleNode(objDomNode, "PU_ATIV_MERC").Text, 8)
                .SubItems(COL_MSG_VALOR) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_FINC").Text)
                .SubItems(COL_MSG_VEICULO_LEGAL) = fgSelectSingleNode(objDomNode, "NO_VEIC_LEGA").Text
                .SubItems(COL_MSG_CNPJ_CONTRAPARTE) = fgFormataCnpj(fgSelectSingleNode(objDomNode, "CO_CNPJ_CNPT").Text)
                .SubItems(COL_MSG_NOME_CONTRAPARTE) = fgFormataCnpj(fgSelectSingleNode(objDomNode, "NO_CNPT").Text)

                .SubItems(COL_MSG_CONTA_CUSTODIA) = fgSelectSingleNode(objDomNode, "CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                .SubItems(COL_MSG_TAXA) = fgVlrXml_To_InterfaceDecimais(fgSelectSingleNode(objDomNode, "PE_TAXA_NEGO").Text, 6)
                .SubItems(COL_MSG_COD_TITULAR_CUTD) = fgSelectSingleNode(objDomNode, "CO_TITL_CUTD").Text

                .SubItems(COL_MSG_CONTRAPARTE_CAMARA) = fgSelectSingleNode(objDomNode, "CO_CNTA_CUTD_CNPT").Text
                .SubItems(COL_MSG_CODIGO_OPERACAO_CETIP) = fgSelectSingleNode(objDomNode, "CO_OPER_CETIP").Text
                .SubItems(COL_MSG_DESCRICAO_ATIVO) = fgSelectSingleNode(objDomNode, "DE_ATIV_MERC").Text
                .SubItems(COL_MSG_MODALIDADE_LIQUIDACAO) = fgSelectSingleNode(objDomNode, "NO_TIPO_LIQU_OPER_ATIV").Text
                
                .SubItems(COL_MSG_DATA_MENSAGEM) = fgDtHrXML_To_Interface(fgSelectSingleNode(objDomNode, "DH_REGT_MESG_SPB").Text)
                
                If fgSelectSingleNode(objDomNode, "DT_OPER_ATIV_RETN").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_RETORNO) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_OPER_ATIV_RETN").Text)
                End If
                .SubItems(COL_MSG_PRAZO_DIAS_RETORNO) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "PZ_DIAS_RETN_OPER_ATIV").Text, False)
                .SubItems(COL_MSG_VALOR_RETORNO) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_OPER_ATIV_RETN").Text)

                'BMC
                .SubItems(COL_MSG_CO_CNTR_SISB) = fgSelectSingleNode(objDomNode, "CO_CNTR_SISB").Text
                .SubItems(COL_MSG_CD_ASSO_CAMB) = fgSelectSingleNode(objDomNode, "CD_ASSO_CAMB").Text
                .SubItems(COL_MSG_VALOR_ME) = fgVlrXml_To_Interface(fgSelectSingleNode(objDomNode, "VA_MOED_ESTR").Text)
                .SubItems(COL_MSG_CO_MOED_ESTR) = fgSelectSingleNode(objDomNode, "CO_MOED_ESTR").Text
                .SubItems(COL_MSG_CO_PRAC) = fgSelectSingleNode(objDomNode, "CO_PRAC").Text
                If fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV_MOED_ESTR").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_LIQUIDACAO_ME) = fgDtXML_To_Date(fgSelectSingleNode(objDomNode, "DT_LIQU_OPER_ATIV_MOED_ESTR").Text)
                End If
                .SubItems(COL_MSG_ISPB_IF_CNPT) = fgSelectSingleNode(objDomNode, "CO_ISPB_IF_CNPT").Text
                .SubItems(COL_MSG_CO_SISB_COTR) = fgSelectSingleNode(objDomNode, "CO_SISB_COTR").Text
                .SubItems(COL_MSG_CO_MESG_SPB) = fgSelectSingleNode(objDomNode, "CO_MESG_SPB").Text

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
    
End Sub

'' Valida os campos selecionados e preenchidos para verificar se permite o
'' processo de registro
Private Function flValidarCampos() As String

Dim strRetorno                              As String

    If PerfilAcesso = BackOffice Then
        'Verifica Qtde de items selecionados
        If fgItemsCheckedListView(lstMensagem) = 0 Then
            strRetorno = "Selecione um item de mensagem para a conciliação."
            lstMensagem.SetFocus
        ElseIf fgItemsCheckedListView(lstMensagem) > 1 Then
            strRetorno = "Selecione apenas um item de mensagem para a conciliação."
            lstMensagem.SetFocus
        ElseIf fgItemsCheckedListView(lstOperacao) = 0 Then
            strRetorno = "Selecione um item de operação para a conciliação."
            lstOperacao.SetFocus
        ElseIf fgItemsCheckedListView(lstOperacao) > 1 Then
            strRetorno = "Selecione apenas um item de operação para a conciliação."
            lstOperacao.SetFocus
        End If
    ElseIf PerfilAcesso = AdmArea Then
        If fgItemsCheckedListView(lstOperacao) = 0 Then
            strRetorno = "Selecione ao menos um item de operação para a liberação."
            lstOperacao.SetFocus
        End If
    Else
        strRetorno = "Perfil de acesso do usuário incompatível. Deve ser 'Backoffice' ou 'Administrador de Área'."
    End If
    
    flValidarCampos = strRetorno
    
End Function

'' Apresenta quaisquer erros de negócio que tenham ocorrido.
Private Sub flApresentarErrosNegocio(ByVal pstrRetorno As String)

Dim xmlErroNegocio                          As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRetorno                              As String

    Set xmlErroNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlErroNegocio.loadXML(pstrRetorno)
    
    'apresenta os erros
    For Each objDomNode In xmlErroNegocio.selectNodes("//Grupo_ErrorInfo/Description")
        strRetorno = strRetorno & objDomNode.Text & vbNewLine & vbNewLine
    Next
    
    Set xmlErroNegocio = Nothing
    
    If strRetorno <> vbNullString Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
    End If

End Sub

'' Formatas os cabeçalhos das colunas dos grids
Private Sub flFormatarListas()

Dim lngLocal
    
    lstOperacao.CheckBoxes = True
    lstMensagem.CheckBoxes = True
    
    If cboLocalLiquidacao.ListIndex <> -1 Then
        lngLocal = fgObterCodigoCombo(Me.cboLocalLiquidacao)
    End If
    
    With lstOperacao.ColumnHeaders
        .Clear
        .Add , , "Operação", 1100
        .Add , , "Tipo Operação", 2550
        .Add , , "Número Comando", flLarguraColunaCamara(1600, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Cód. Contr. SISBACEN", flLarguraColunaCamara(2000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Cód. Assoc. Câmbio", flLarguraColunaCamara(2000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Praça", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Moeda", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Data Operação", 1275
        .Add , , "Data Liquidação", 1275
        .Add , , "Data Liquidação Moeda Estr", flLarguraColunaCamara(3000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "D/C", 850
        .Add , , "ID Ativo", flLarguraColunaCamara(1440, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Data Vencimento", 1275
        .Add , , "Quantidade", 1244, lvwColumnRight
        .Add , , "PU", 1344, lvwColumnRight
        .Add , , "Valor", 1695, lvwColumnRight
        .Add , , "Valor Moeda Estrangeira", flLarguraColunaCamara(2500, enumLocalLiquidacao.BMC, lngLocal), lvwColumnRight
        .Add , , "Veículo Legal", 1440
        .Add , , "CNPJ Contraparte", 1695
        .Add , , "Contraparte", 2900
        .Add , , "Conta Custódia", 1440
        .Add , , "Taxa", 1100, lvwColumnRight
        .Add , , "Cód. Titular Custodiante", 1440
        .Add , , "Contraparte Câmara", 1440
        .Add , , "Operação CETIP", flLarguraColunaCamara(1440, enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Descrição Ativo", 1440
        .Add , , "Modalidade Liquidação", 1440
        .Add , , "Valor Retorno", flLarguraColunaCamara(1695, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal), lvwColumnRight
        .Add , , "Data Retorno", flLarguraColunaCamara(1275, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Prazo Dias Retorno", flLarguraColunaCamara(1275, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal), lvwColumnRight
        .Add , , "ISPB IF Contraparte", 1000
        .Add , , "Canal SISBACEN Corretora", 3000
        .Add , , "CNPJ/CPF Comitente", 4000

    End With


    With lstMensagem.ColumnHeaders
        .Clear
        .Add , , "Mensagem", 1100
        .Add , , "Tipo Negociação", 2550
        .Add , , "Número Comando", flLarguraColunaCamara(1600, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Cód. Contr. SISBACEN", flLarguraColunaCamara(2000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Cód. Assoc. Câmbio", flLarguraColunaCamara(2000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Praça", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Moeda", flLarguraColunaCamara(1000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Data Operação", 1275
        .Add , , "Data Liquidação", 1275
        .Add , , "Data Liquidação Moeda Estr", flLarguraColunaCamara(3000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "D/C", 850
        .Add , , "ID Ativo", flLarguraColunaCamara(1440, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Data Vencimento", 1275
        .Add , , "Quantidade", 1244, lvwColumnRight
        .Add , , "PU", 1344, lvwColumnRight
        .Add , , "Valor", 1695, lvwColumnRight
        .Add , , "Valor Moeda Estrangeira", flLarguraColunaCamara(2500, enumLocalLiquidacao.BMC, lngLocal), lvwColumnRight
        .Add , , "Veículo Legal", 1440
        .Add , , "CNPJ Contraparte", 1695
        .Add , , "Contraparte", 2900
        .Add , , "Conta Custódia", 1440
        .Add , , "Taxa", 1100, lvwColumnRight
        .Add , , "Cód. Titular Custodiante", 1440
        .Add , , "Contraparte Câmara", 1440
        .Add , , "Operação CETIP", flLarguraColunaCamara(1440, enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Descrição Ativo", 1440
        .Add , , "Modalidade Liquidação", 1440
        .Add , , "Valor Retorno", flLarguraColunaCamara(1695, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal), lvwColumnRight
        .Add , , "Data Retorno", flLarguraColunaCamara(1275, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal)
        .Add , , "Prazo Dias Retorno", flLarguraColunaCamara(1275, enumLocalLiquidacao.BMA & ";" & enumLocalLiquidacao.CETIP, lngLocal), lvwColumnRight
        .Add , , "ISPB IF Contraparte", 1000
        .Add , , "Data Mensagem", 2200
        .Add , , "Canal SISBACEN Corretora", flLarguraColunaCamara(3000, enumLocalLiquidacao.BMC, lngLocal)
        .Add , , "Código Mensagem", flLarguraColunaCamara(2000, enumLocalLiquidacao.BMC, lngLocal)
    
    End With

End Sub

'Define a largura da coluna no grid, de acordo com a camara selecionada (exibe/esconde)
Function flLarguraColunaCamara(ByVal lngLargura, strCamaras, lngCamara)
    
    'Passar camaras permitidas separados por ';'
    
    If lngCamara = 0 Or InStr(";" & strCamaras & ";", ";" & lngCamara & ";") > 0 Then
        flLarguraColunaCamara = lngLargura
    Else
        flLarguraColunaCamara = 0
    End If
        
End Function

'' Monta XML de conciliação com os dados da mensagem
Private Function flMontarXMLMensagem() As String

Dim xmlDomMensagem                          As MSXML2.DOMDocument40
Dim lngCont                                 As Long

    'Monta XML para a mensagem
    Set xmlDomMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomMensagem, "", "Repeat_Conciliacao", "")
    Call fgAppendNode(xmlDomMensagem, "Repeat_Conciliacao", "Grupo_Mensagem", "")
    
    With lstMensagem.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                Call fgAppendNode(xmlDomMensagem, "Grupo_Mensagem", "NU_CTRL_IF", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_MSG_NUMERO_CONTROLE_IF))
                Call fgAppendNode(xmlDomMensagem, "Grupo_Mensagem", "DH_REGT_MESG_SPB", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_MSG_DATA_REGISTRO_MESG_SPB))
                Call fgAppendNode(xmlDomMensagem, "Grupo_Mensagem", "NU_SEQU_CNTR_REPE", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_MSG_NUMERO_SEQUENCIA_CONTADOR_REPETICAO))
                Call fgAppendNode(xmlDomMensagem, "Grupo_Mensagem", "DH_ULTI_ATLZ", _
                                        .Item(lngCont).Tag)
                Exit For
            End If
        Next
    End With
    
    flMontarXMLMensagem = xmlDomMensagem.xml
    
    Set xmlDomMensagem = Nothing

End Function

'' Monta XML de conciliação com os dados da operacao
Private Function flMontarXMLOperacao() As String

Dim xmlDomOperacao                          As MSXML2.DOMDocument40
Dim lngCont                                 As Long
Dim lngLocalLiquidacao                      As Long
Dim intIgnoraGradeHorario                   As Integer

    lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)

    'Monta XML para a operacao
    Set xmlDomOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlDomOperacao.loadXML ""
    Call fgAppendNode(xmlDomOperacao, "", "Repeat_Conciliacao", "")
    
    With lstOperacao.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                Call fgAppendNode(xmlDomOperacao, "Repeat_Conciliacao", "Grupo_Operacao", "")

                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "CO_EMPR", fgObterCodigoCombo(cboEmpresa.Text), "Repeat_Conciliacao")
                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "NU_SEQU_OPER_ATIV", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_OP_NUMERO_SEQUENCIA_OPERACAO), "Repeat_Conciliacao")
                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "TP_OPER", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_OP_TIPO_OPERACAO), "Repeat_Conciliacao")
                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "DH_ULTI_ATLZ", _
                                        .Item(lngCont).Tag, "Repeat_Conciliacao")
                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "LocalLiquidacao", _
                                        lngLocalLiquidacao, "Repeat_Conciliacao")
                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "TipoConfirmacao", enumTipoConfirmacao.operacao, "Repeat_Conciliacao")
                
                intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed, 1, 0)
                Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "IgnoraGradeHorario", _
                                        intIgnoraGradeHorario, "Repeat_Conciliacao")
            
                If .Item(lngCont).Text = STR_CAMPOS_OPCIONAIS_DIVERGENTES Then
                    Call fgAppendNode(xmlDomOperacao, "Grupo_Operacao", "NaoVerificaCampoOpcionais", "1", "Repeat_Conciliacao")
                End If
            End If
        Next
    End With
    
    flMontarXMLOperacao = xmlDomOperacao.xml
    
    Set xmlDomOperacao = Nothing

End Function

'' Habilita/Desabilita os botões de acordo com o perfil de acesso do usuário
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)
'Mostra somente os botões permitidos ao usuário de acordo com o seu perfil de acesso
    
    With tlbFiltro
    
        .Buttons("Concordar").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice)
        .Buttons("Liberar").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)

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

'' Configura o perfil de acesso do usuário
Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
'Controla o perfil de acesso do usuário

    lngPerfil = pPerfil
    
    If pPerfil = BackOffice Then
        Me.Caption = "Conciliação (Registro e Operação)"
    Else
        Me.Caption = "Liberação (Registro e Operação)"
    End If
    
    flConfigurarBotoesPorPerfil PerfilAcesso
    flPosicionaControles
    flMontaTela
    
End Property

Property Get PerfilAcesso() As enumPerfilAcesso
'Retorna o perfil de acesso do usuário
        
    PerfilAcesso = lngPerfil
    
End Property

'' Mostra o resultado do último processo de registro efetuado
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

'' Exibe de forma diferenciada quaisquer itens que tenham sido rejeitados por
'' grade de horário.
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            With lstOperacao.ListItems
                For lngCont = 1 To .Count
                    If UCase(Split(Mid(.Item(lngCont).Key, 2), "|")(0)) = UCase(fgSelectSingleNode(objDomNode, "Operacao").Text) Then
                        For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                            .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                        Next
                        
                        .Item(lngCont).Text = "Horário Excedido"
                        .Item(lngCont).ToolTipText = "Horário limite p/envio da mensagem excedido"
                        .Item(lngCont).ForeColor = vbRed
                        
                        .Item(lngCont).Checked = False
                        
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

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, _
             enumTipoSelecao.DesmarcarTodas
            
            Call fgMarcarDesmarcarTodas(IIf(intControleMenuPopUp = enumTipoConfirmacao.operacao, _
                                                lstOperacao, lstMensagem), _
                                        Retorno)
    
    End Select
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

'' Retorna a descrição de um Tipo de Negociação BMA
Private Function flDescricaoTpNegcBMA(ByVal pstrCodigo As String) As String
    
'Obtem a descrição do domínio TpNegcBMA
Dim objNode                     As MSXML2.IXMLDOMNode

    Set objNode = xmlDominioTpNegcBMA.selectSingleNode("//DE_DOMI[../CO_DOMI=" & pstrCodigo & "]")

    If objNode Is Nothing Then
        flDescricaoTpNegcBMA = "Tipo Inesperado"
    Else
        flDescricaoTpNegcBMA = objNode.Text
    End If

End Function

Private Sub flMarcaCampoOpcional(ByVal pstrRetorno As String)

Dim xmlErroNegocio                          As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRetorno                              As String
Dim blnTudoOpcional                         As Boolean
Dim objLSI                                  As ListSubItem

    Set xmlErroNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlErroNegocio.loadXML(pstrRetorno)
    
    'Verifica se todos os erros que ocorreram são em campos opcionais
    
    blnTudoOpcional = True
    'apresenta os erros
    For Each objDomNode In xmlErroNegocio.selectNodes("//Grupo_ErrorInfo/Number")
        If Not fgIN(Val(objDomNode.Text), 3078) Then
            blnTudoOpcional = False
            Exit For
        End If
    Next
    
    If blnTudoOpcional Then
        lblMensagem = "ATENÇÃO: Todos os campos divergentes da última tentativa são opcionais. Clique em 'Concordar' novamente para ignorar a validação."
        lstOperacao.SelectedItem.Text = STR_CAMPOS_OPCIONAIS_DIVERGENTES
        lstOperacao.SelectedItem.ForeColor = vbRed
        For Each objLSI In lstOperacao.SelectedItem.ListSubItems
            objLSI.ForeColor = vbRed
        Next
    End If
    
End Sub

Private Sub tmrRefresh_Timer()

On Error GoTo ErrorHandler

    If Not IsNumeric(txtTimer.Text) Then Exit Sub
    
    If CLng(txtTimer.Text) = 0 Then Exit Sub
    
    If fgVerificaJanelaVerificacao() Then Exit Sub
    
    fgCursor True

    intContMinutos = intContMinutos + 1
    
    If intContMinutos >= txtTimer.Text Then

        Call flMontaTela

        intContMinutos = 0
    End If

    fgCursor False

Exit Sub
ErrorHandler:
    
    fgCursor False
    
    fgRaiseError App.EXEName, TypeName(Me), "tmrRefresh_Timer", 0

End Sub


