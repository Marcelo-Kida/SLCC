VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConciliacaoOperacao 
   Caption         =   "Conciliação de Operação e Mensagem SPB"
   ClientHeight    =   9900
   ClientLeft      =   480
   ClientTop       =   1080
   ClientWidth     =   14295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   14295
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   12240
      Top             =   240
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
            Picture         =   "frmConciliacaoOperacao.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoOperacao.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   9570
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   582
      ButtonWidth     =   2566
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
            Caption         =   "Conciliar          "
            Key             =   "Conciliar"
            Object.ToolTipText     =   "Conciliar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraControles 
      Caption         =   "fraControles"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   7920
      Width           =   14055
      Begin VB.ComboBox cboTipoJustificativa 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   4350
      End
      Begin VB.TextBox txtComentario 
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   8685
      End
      Begin VB.Frame fraResumo 
         Caption         =   "Quantidades A Conciliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   8880
         TabIndex        =   10
         Top             =   0
         Width           =   5175
         Begin VB.TextBox txtDiferenca 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0"
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkConciliacaoFinalizada 
            Caption         =   "Conciliação Finalizada"
            Enabled         =   0   'False
            Height          =   252
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1875
         End
         Begin VB.TextBox txtQtdVlrOperacao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtQtdVlrMensagem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblDiferenca 
            AutoSize        =   -1  'True
            Caption         =   "Diferença"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   3480
            TabIndex        =   17
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Operação"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem"
            Height          =   195
            Left            =   1800
            TabIndex        =   15
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Label lblTipoJustificativa 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Justificativa"
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblComentario 
         AutoSize        =   -1  'True
         Caption         =   "Comentário"
         Enabled         =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   720
         Width           =   795
      End
   End
   Begin MSComctlLib.ListView lstMensagem 
      Height          =   3285
      Left            =   105
      TabIndex        =   7
      Top             =   4560
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   5794
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
         Text            =   "Conciliar Mensagem"
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
   End
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   3525
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   6218
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
   End
   Begin VB.ComboBox cboLocalLiquidacao 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2790
   End
   Begin VB.ComboBox cboTipoOperacao 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   4350
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   4350
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   120
      MousePointer    =   7  'Size N S
      Top             =   4440
      Width           =   14040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Operação"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Local Liquidação"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmConciliacaoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:39
'-------------------------------------------------
'' Objeto responsável pela conciliação de mensagens e operações através de
'' interação com a camada de controle de caso de uso MIU.
''
'' Classes especificamente consideradas de destino:
''   A8MIU.clsOperacao
''   A8MIU.clsOperacaoMensagem
''   A8MIU.clsMensagem
''
Option Explicit

Private lngItemCheckedOperacao              As Long
Private lngItemCheckedMensagem              As Long
Private intTipoOperacao                     As enumTipoOperacaoLQS
Private lngLocalLiquidacao                  As Long

Private dblSomaValorOperacoes               As Double
Private dblSomaValorMensagens               As Double

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlDomTipoOperacao                  As MSXML2.DOMDocument40
Private xmlDomOperacao                      As MSXML2.DOMDocument40     '<- Utilizado para conciliação de DESPESAS
Private Const strFuncionalidade             As String = "frmConciliacaoOperacao"
Private fblnDummyH                          As Boolean

Private Const POS_NUMERO_CONTROLE_IF        As Integer = 0
Private Const POS_DATA_REGISTRO_MESG_SPB    As Integer = 1
Private Const POS_NUMERO_SEQUENCIA_OPERACAO As Integer = 2
Private Const POS_CO_VEIC_LEGA_MSG          As Integer = 3

Private Const POS_NU_SEQU_OPER_ATIV         As Integer = 0
Private Const POS_CO_VEIC_LEGA_OPER         As Integer = 1
Private Const POS_TP_OPER                   As Integer = 2

Private Const POS_DATA_ULTIMA_ATLZ_MESG     As Integer = 0
Private Const POS_VALOR_MESG                As Integer = 1
Private Const POS_NUMERO_COMANDO            As Integer = 2
Private Const POS_TEXTO_XML                 As Integer = 3

'Constantes de Configuração de Colunas de Operação (Por Quantidade)
Private Const COL_OP_VEICULO_LEGAL          As Integer = 1
Private Const COL_OP_CONTA_CUSTODIA         As Integer = 2
Private Const COL_OP_NUMERO_COMANDO         As Integer = 3
Private Const COL_OP_NUMERO_COMANDO_ORIG    As Integer = 4
Private Const COL_OP_TIPO_OPERACAO          As Integer = 5
Private Const COL_OP_ID_ATIVO               As Integer = 6
Private Const COL_OP_DT_VENC_ATIV           As Integer = 7
Private Const COL_OP_QTD_TOTAL              As Integer = 8
Private Const COL_OP_QTD_A_CONCILIAR        As Integer = 9
Private Const COL_OP_QTD_CONCILIADA         As Integer = 10
Private Const COL_OP_PU                     As Integer = 11
Private Const COL_OP_VALOR                  As Integer = 12
Private Const COL_OP_CODIGO                 As Integer = 13
Private Const COL_OP_DATA_OPERACAO          As Integer = 14

'Constantes de Configuração de Colunas de Operação  (Por Valor)
Private Const COL_OP2_GRUPO_VEICULO         As Integer = 1
Private Const COL_OP2_VALOR                 As Integer = 2

'Constantes de Configuração de Colunas de Mensagem  (Por Quantidade)
Private Const COL_MSG_VEICULO_LEGAL         As Integer = 1
Private Const COL_MSG_CONTA_CUSTODIA        As Integer = 2
Private Const COL_MSG_NUMERO_COMANDO        As Integer = 3
Private Const COL_MSG_NUMERO_COMANDO_ORIG   As Integer = 4
Private Const COL_MSG_TIPO_ROTINA_ABER      As Integer = 5
Private Const COL_MSG_ID_ATIVO              As Integer = 6
Private Const COL_MSG_DT_VENC_ATIV          As Integer = 7
Private Const COL_MSG_QTD_TOTAL             As Integer = 8
Private Const COL_MSG_QTD_A_CONCILIAR       As Integer = 9
Private Const COL_MSG_QTD_CONCILIADA        As Integer = 10
Private Const COL_MSG_PU                    As Integer = 11
Private Const COL_MSG_VALOR                 As Integer = 12
Private Const COL_MSG_DATA_MENSAGEM         As Integer = 13

'Constantes de Configuração de Colunas de Operação  (Por Valor)
Private Const COL_MSG2_CODIGO_MENSAGEM      As Integer = 1
Private Const COL_MSG2_VALOR                As Integer = 2

'Constantes Tipo Operação SELIC Eventos Todos
Private Const INT_CODIGO_EVENTOS_TODOS      As Integer = 9999
Private Const STR_CODIGO_EVENTOS_TODOS      As String = "Eventos - Todos"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

Private Sub cboEmpresa_Click()
    
    flMontaTela
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoOperacao - cboEmpresa_Click", Me.Caption
    
End Sub

Private Sub cboLocalLiquidacao_Click()

    If cboLocalLiquidacao.ListIndex <> -1 Then
        lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)
    
        'Habilita combo de justificativa por local de liquidacao
        Select Case lngLocalLiquidacao
            Case enumLocalLiquidacao.BMA
                cboTipoJustificativa.ListIndex = -1
                cboTipoJustificativa.Enabled = False
                cboTipoJustificativa.BackColor = txtDiferenca.BackColor
                txtComentario.BackColor = txtDiferenca.BackColor
            Case Else
                cboTipoJustificativa.Enabled = True
                cboTipoJustificativa.BackColor = vbWhite
                txtComentario.BackColor = vbWhite
        End Select
    
        intTipoOperacao = 0
        Call flCarregarTipoOperacaoConciliacao(lngLocalLiquidacao)
        
        If lngLocalLiquidacao = enumLocalLiquidacao.SELIC And cboTipoOperacao.ListCount > 0 Then
            cboTipoOperacao.AddItem INT_CODIGO_EVENTOS_TODOS & " - " & STR_CODIGO_EVENTOS_TODOS
        End If
            
        flSugestaoCombo cboTipoOperacao
    
    End If
    
'    flMontaTela
    flLimparLista Me.lstOperacao
    flLimparLista Me.lstMensagem
    flLimparControleJustificativa

End Sub

Private Sub cboTipoJustificativa_Click()
    lblComentario.Enabled = (cboTipoJustificativa.ListIndex <> -1)
    txtComentario.Enabled = lblComentario.Enabled
End Sub

Private Sub cboTipoOperacao_Click()
        
    intTipoOperacao = fgObterCodigoCombo(Me.cboTipoOperacao)
'    chkDispRepasseFinc.Enabled = (intTipoOperacao = enumTipoOperacaoLQS.EventosResgate Or _
'                                  intTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
'                                  intTipoOperacao = enumTipoOperacaoLQS.EventosAmortização)

    If cboEmpresa.ListIndex = -1 Then
        flSugestaoCombo cboEmpresa
    Else
        flMontaTela
    End If
    
    Exit Sub
    
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoOperacao - cboTipoOperacao_Click", Me.Caption

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("Refresh"))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoOperacao - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .tlbFiltro.Left = 0
        .tlbFiltro.Top = .ScaleHeight - .tlbFiltro.Height
        
        .imgDummyH.Left = 0
        .imgDummyH.Width = .ScaleWidth
        
        .lstOperacao.Height = .imgDummyH.Top - .imgDummyH.Height - 800
        .lstOperacao.Width = .Width - 350
        
        .lstMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        .lstMensagem.Height = .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 1800
        .lstMensagem.Width = .Width - 350
        
        .fraControles.Left = lstMensagem.Left
        .fraControles.Width = lstMensagem.Width
        .fraControles.Top = .lstMensagem.Top + .lstMensagem.Height + 120
        '.lblTipoJustificativa.Top = .lstMensagem.Top + .lstMensagem.Height + 120
        '.cboTipoJustificativa.Top = .lblTipoJustificativa.Top + 240
        .cboTipoJustificativa.Width = .txtComentario.Width - (.txtComentario.Width \ 2)
        '.lblComentario.Top = .cboTipoJustificativa.Top + 480
        '.txtComentario.Top = .lblComentario.Top + 240
        .txtComentario.Width = fraControles.Width - .fraResumo.Width - 480
        .fraResumo.Left = .txtComentario.Width + 300
        '.fraResumo.Top = .lblTipoJustificativa.Top
    
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConciliacaoOperacao = Nothing
End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fblnDummyH = True
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not fblnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyH.Top = Y + imgDummyH.Top

    On Error Resume Next
    
    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        If .imgDummyH.Top > (.Height - 3500) And (.Height - 3500) > 0 Then
            .imgDummyH.Top = .Height - 3500
        End If
        
        .lstOperacao.Height = .imgDummyH.Top - .imgDummyH.Height - 800
        .lstMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        .lstMensagem.Height = .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 1800
        
        '.lblTipoJustificativa.Top = .lstMensagem.Top + .lstMensagem.Height + 120
        '.cboTipoJustificativa.Top = .lblTipoJustificativa.Top + 240
        '.lblComentario.Top = .cboTipoJustificativa.Top + 480
        '.txtComentario.Top = .lblComentario.Top + 240
        
        .fraResumo.Top = .lblTipoJustificativa.Top
    End With
    
    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fblnDummyH = False
End Sub

Private Sub lstMensagem_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstMensagem, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstMensagem_ColumnClick"
End Sub

Private Sub lstMensagem_DblClick()

    If Not lstMensagem.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_NUMERO_SEQUENCIA_OPERACAO)
            .NumeroControleIF = Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_NUMERO_CONTROLE_IF)
            .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(Mid(lstMensagem.SelectedItem.Key, 2), "|")(POS_DATA_REGISTRO_MESG_SPB))
            .Show vbModal
        End With
    End If
    
End Sub

Private Sub lstMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
Dim intTipoOperacao                         As Integer
Dim objListItem                             As ListItem
Dim strMensagem                             As String
Dim strTipoRotinaAberSelec                  As String
Dim strTipoRotinaAbertura                   As String

On Error GoTo ErrorHandler

    Item.Selected = True
    fgCursor True
    Call flMostrarDiferencaConciliacao(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)
    fgCursor

    If Item.Checked Then
        strMensagem = vbNullString
        intTipoOperacao = fgObterCodigoCombo(cboTipoOperacao.Text)
        
        If intTipoOperacao = INT_CODIGO_EVENTOS_TODOS Then
            strTipoRotinaAberSelec = Left$(Item.SubItems(COL_MSG_TIPO_ROTINA_ABER), 4)
            
            For Each objListItem In lstMensagem.ListItems
                If objListItem.Checked Then
                    strTipoRotinaAbertura = Left$(objListItem.SubItems(COL_MSG_TIPO_ROTINA_ABER), 4)
                    
                    If strTipoRotinaAbertura <> strTipoRotinaAberSelec Then
                        strMensagem = "Somente itens de um mesmo tipo de operação SELIC podem ser selecionados."
                        GoTo ExibirMensagemErro
                    End If
                End If
            Next
            
            Select Case strTipoRotinaAberSelec
                Case enumCodigoOperacaoSelic.Resgate '"1012"
                    intTipoOperacao = enumTipoOperacaoLQS.EventosResgate
                Case enumCodigoOperacaoSelic.Amortizacao ' "1010"
                    intTipoOperacao = enumTipoOperacaoLQS.EventosAmortização
                Case enumCodigoOperacaoSelic.Juros '"1060"
                    intTipoOperacao = enumTipoOperacaoLQS.EventosJuros
            End Select
                    
            For Each objListItem In lstOperacao.ListItems
                If objListItem.Checked Then
                    If Split(Mid$(objListItem.Key, 2), "|")(POS_TP_OPER) <> intTipoOperacao Then
                        strMensagem = "Somente itens com um código de operação SELIC correspondentes ao selecionado na lista de operações, podem ser selecionados na lista de mensagens."
                        GoTo ExibirMensagemErro
                    End If
                End If
            Next
            
        End If
    End If
                
ExibirMensagemErro:
    If strMensagem <> vbNullString Then
        frmMural.Display = strMensagem
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        
        Item.Checked = False
    End If

Exit Sub
ErrorHandler:
   
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstMensagem_ItemCheck"

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
    
    If intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic Then
        If Not lstOperacao.SelectedItem Is Nothing Then
            With frmDetalheOperacao
                .SequenciaOperacao = Mid(lstOperacao.SelectedItem.Key, 2, InStr(2, lstOperacao.SelectedItem.Key, "|") - 2)
                .Show vbModal
            End With
        End If
    End If
    
End Sub

Private Sub lstOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim intTipoOperacao                         As Integer
Dim objListItem                             As ListItem
Dim intTipoOperSelecionado                  As Integer
Dim strMensagem                             As String
Dim strTipoRotinaAbertura                   As String

On Error GoTo ErrorHandler

    Item.Selected = True
    fgCursor True
    Call flMostrarDiferencaConciliacao(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)
    fgCursor False

    If Item.Checked Then
        strMensagem = vbNullString
        intTipoOperacao = fgObterCodigoCombo(cboTipoOperacao.Text)
        
        If intTipoOperacao = INT_CODIGO_EVENTOS_TODOS Then
            intTipoOperSelecionado = Split(Mid$(Item.Key, 2), "|")(POS_TP_OPER)
            
            For Each objListItem In lstOperacao.ListItems
                If objListItem.Checked Then
                    intTipoOperacao = Split(Mid$(objListItem.Key, 2), "|")(POS_TP_OPER)
                    
                    If intTipoOperacao <> intTipoOperSelecionado Then
                        strMensagem = "Somente itens de um mesmo tipo de operação podem ser selecionados."
                        GoTo ExibirMensagemErro
                    End If
                End If
            Next
            
            Select Case intTipoOperSelecionado
                Case enumTipoOperacaoLQS.EventosResgate
                    strTipoRotinaAbertura = enumCodigoOperacaoSelic.Resgate '"1012"
                Case enumTipoOperacaoLQS.EventosAmortização
                    strTipoRotinaAbertura = enumCodigoOperacaoSelic.Amortizacao ' "1010"
                Case enumTipoOperacaoLQS.EventosJuros
                    strTipoRotinaAbertura = enumCodigoOperacaoSelic.Juros ' "1060"
            End Select
                    
            For Each objListItem In lstMensagem.ListItems
                If objListItem.Checked Then
                    If Left$(objListItem.SubItems(COL_MSG_TIPO_ROTINA_ABER), 4) <> strTipoRotinaAbertura Then
                        strMensagem = "Somente itens com um tipo de operação correspondente ao selecionado na lista de mensagens, podem ser selecionados na lista de operações."
                        GoTo ExibirMensagemErro
                    End If
                End If
            Next
            
        End If
    End If
                
ExibirMensagemErro:
    If strMensagem <> vbNullString Then
        frmMural.Display = strMensagem
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        
        Item.Checked = False
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ItemCheck"

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strRetorno                              As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    fgCursor True
    
    Select Case Button.Key
        Case "Refresh"
            Call cboEmpresa_Click                       '<-- Recarrega as Listas (Operação e Mensagem)
            
        Case "Conciliar"
           'Verifica inconsistências da INTERFACE
            strRetorno = flValidarCampos(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)
            If strRetorno <> "" Then
                frmMural.Caption = Me.Caption
                frmMural.Display = strRetorno
                frmMural.Show vbModal
                GoTo ExitSub
            End If
            
            strRetorno = flConciliar(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)
           
            If strRetorno = vbNullString Then                   '<-- Conciliação OK
                MsgBox "Conciliação efetuada com sucesso.", vbInformation, Me.Caption
                Call cboEmpresa_Click                           '<-- Recarrega as Listas (Operação e Mensagem)
            Else
                If InStr(strRetorno, "Repeat_Info") > 0 Then
                    Call flApresentaResultDespesa(strRetorno)   '<-- Resultado de despesas geradas
                    Call cboEmpresa_Click                       '<-- Recarrega as Listas (Operação e Mensagem)
                Else
                    Call flApresentarErrosNegocio(strRetorno)   '<-- Erros de Negócio encontrados (Warnings!)
                End If
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
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

'' Obtém , através da camada de controle de caso de uso, as propriedades
'' utilizadas pelo objeto, através da camada de controle de caso de uso, método
'' A8MIU.clsMIU.ObterMapaNavegacao
Public Function flInicializar() As Boolean

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
    Dim objTipoOperacao     As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
    Dim objTipoOperacao     As A8MIU.clsTipoOperacao
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
    
On Error GoTo ErrorHandler
    
    Set xmlDomTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set objTipoOperacao = fgCriarObjetoMIU("A8MIU.clsTipoOperacao")

   
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConciliacaoOperacao", "flInicializar")
    End If
    
    If Not xmlDomTipoOperacao.loadXML(objTipoOperacao.ObterTiposOperacaoConciliacao(0, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlDomTipoOperacao, App.EXEName, "frmConciliacaoOperacao", "flInicializar")
    End If
   
    fraControles.BorderStyle = 0
    
    Set objMIU = Nothing
    
Exit Function
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Function

'' Concilia mensagens e operações, através  de interação com a camada de controle
'' de caso de uso MIU, método A8MIU.clsMensagem.Conciliar. Retorna uma String
'' contendo os erros que ocorreram em uma operação ou vbNullString caso nenhum
'' erro tenha ocorrido
Private Function flConciliar(Optional ByVal pblnPorQuantidade As Boolean = True) As String

#If EnableSoap = 1 Then
    Dim objConciliacao      As MSSOAPLib30.SoapClient30
#Else
    Dim objConciliacao      As A8MIU.clsOperacaoMensagem
#End If

Dim strRetorno              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    strRetorno = flMontarXMLConciliacao(pblnPorQuantidade)
    
    Set objConciliacao = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
    flConciliar = objConciliacao.Conciliar(strRetorno, vntCodErro, vntMensagemErro)
    
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

'' Obtem as operações passíveis de conciliação e preenche o listview de operações
'' com as mesmas, através de interação com a camada de controle de caso de uso MIU,
'' método A8MIU.clsOperacao.ObterDetalheOperacao
Private Sub flCarregarListaOperacao(ByVal pintTipoOperacao As Integer, _
                                    ByVal plngEmpresa As Long, _
                                    ByVal plngLocalLiquidacao As Long, _
                           Optional ByVal pblnVisaoPorQuantidade As Boolean = True)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim strRetLeituraQtdConciliada              As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomLeituraQtdConciliada              As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacaoRotinaAbertura", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacaoRotinaAbertura", "TipoOperacaoRotinaAbertura", "S")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
    
    If pintTipoOperacao <> INT_CODIGO_EVENTOS_TODOS Then
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", pintTipoOperacao)
    Else
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.EventosAmortização)
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.EventosJuros)
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.EventosResgate)
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LL", plngLocalLiquidacao)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    
    'Filtra as operações A CONCILIAR pela Data Atual em diante
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
    '>>> --------------------------------------------------------------------------------------------------
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    
    If pblnVisaoPorQuantidade Then
        strRetLeituraQtdConciliada = objOperacao.ObterQtdConciliadaOperacao(vbNullString, vntCodErro, vntMensagemErro)
        strRetLeitura = objOperacao.ObterDetalheOperacao(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    Else
        strRetLeitura = objOperacao.ObterValoresPorGrupoVeiculoLegal(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    End If
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flLimparLista(Me.lstOperacao)
    
    Set objOperacao = Nothing
    
    If pblnVisaoPorQuantidade Then
        If strRetLeituraQtdConciliada <> vbNullString Then
            Set xmlDomLeituraQtdConciliada = CreateObject("MSXML2.DOMDocument.4.0")
            
            If Not xmlDomLeituraQtdConciliada.loadXML(strRetLeituraQtdConciliada) Then
                Call fgErroLoadXML(xmlDomLeituraQtdConciliada, App.EXEName, TypeName(Me), "flCarregarListaOperacao")
            End If
        End If
    
        If pintTipoOperacao <> INT_CODIGO_EVENTOS_TODOS Then
            lstOperacao.ColumnHeaders(COL_OP_TIPO_OPERACAO + 1).Width = 0
        Else
            lstOperacao.ColumnHeaders(COL_OP_TIPO_OPERACAO + 1).Width = 1200
        End If
        
        If pintTipoOperacao = enumTipoOperacaoLQS.EventosAmortização Or _
           pintTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
           pintTipoOperacao = enumTipoOperacaoLQS.EventosResgate Or _
           pintTipoOperacao = INT_CODIGO_EVENTOS_TODOS Then
            lstOperacao.ColumnHeaders(COL_OP_NUMERO_COMANDO_ORIG + 1).Width = 0
        Else
            lstOperacao.ColumnHeaders(COL_OP_NUMERO_COMANDO_ORIG + 1).Width = 1340
        End If
    End If
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaOperacao")
        End If

        If pblnVisaoPorQuantidade Then
            For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                With lstOperacao.ListItems.Add(, _
                        "k" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & "|" & _
                              objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & "|" & _
                              objDomNode.selectSingleNode("TP_OPER").Text)

                    .Tag = objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text

                    .SubItems(COL_OP_VEICULO_LEGAL) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_OP_CONTA_CUSTODIA) = objDomNode.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                    .SubItems(COL_OP_NUMERO_COMANDO_ORIG) = objDomNode.selectSingleNode("NU_COMD_OPER_RETN").Text
                    .SubItems(COL_OP_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text

                    If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                    End If

                    .SubItems(COL_OP_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                    .SubItems(COL_OP_ID_ATIVO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
                    .SubItems(COL_OP_DT_VENC_ATIV) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)

                    If objDomNode.selectSingleNode("QT_ATIV_MERC").Text <> vbNullString Then
                        .SubItems(COL_OP_QTD_TOTAL) = objDomNode.selectSingleNode("QT_ATIV_MERC").Text
                    Else
                        .SubItems(COL_OP_QTD_TOTAL) = 0
                    End If

                    'Verifica se o XML de Quantidade Conciliada foi carregado...
                    If Not xmlDomLeituraQtdConciliada Is Nothing Then
                        '...se sim, verifica se existe quantidade para a operação corrente...
                        If Not xmlDomLeituraQtdConciliada.selectSingleNode("Repeat_QtdConciliadaOperacao/Grupo_QtdConciliadaOperacao[NU_SEQU_OPER_ATIV='" & _
                            objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & "']/QT_ATIV_MERC_CNCL") Is Nothing Then
                            
                            .SubItems(COL_OP_QTD_CONCILIADA) = Val(xmlDomLeituraQtdConciliada.selectSingleNode("Repeat_QtdConciliadaOperacao/Grupo_QtdConciliadaOperacao[NU_SEQU_OPER_ATIV='" & _
                                        objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & "']/QT_ATIV_MERC_CNCL").Text)
                        
                        '...se não, carrega ZERO
                        Else
                            .SubItems(COL_OP_QTD_CONCILIADA) = 0
                        End If
                    
                    '...se não, carrega ZERO
                    Else
                        .SubItems(COL_OP_QTD_CONCILIADA) = 0
                    End If
                    
                    'Quantidade Total - Quantidade Conciliada
                    .SubItems(COL_OP_QTD_A_CONCILIAR) = (Val(.SubItems(COL_OP_QTD_TOTAL)) - Val(.SubItems(COL_OP_QTD_CONCILIADA)))
                    
                    .SubItems(COL_OP_PU) = fgVlrXml_To_InterfaceDecimais(objDomNode.selectSingleNode("PU_ATIV_MERC").Text, 8)
                    .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                    .SubItems(COL_OP_CODIGO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
                            
                End With
            Next
        Else
            Set xmlDomOperacao = Nothing    '<-- Destrói o XML anterior, se existir
            Set xmlDomOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            If xmlDomLeitura.documentElement.selectNodes("Repeat_CodigoOperacao").length > 0 Then
                'Armazena as operações que compõe os totais
                Call xmlDomOperacao.loadXML(xmlDomLeitura.xml)
                
                For Each objDomNode In xmlDomLeitura.documentElement.selectSingleNode("Repeat_ValoresPorGrupoVeiculoLegal").childNodes
                    With lstOperacao.ListItems.Add(, _
                            "k" & objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text)
                        
                        .Tag = objDomNode.selectSingleNode("VA_OPER_ATIV").Text
                        
                        .SubItems(COL_OP2_GRUPO_VEICULO) = objDomNode.selectSingleNode("NO_GRUP_VEIC_LEGA").Text
                        .SubItems(COL_OP2_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                    End With
                Next
            End If
        End If
    End If
    
    Call fgClassificarListview(Me.lstOperacao, lngIndexClassifListOper, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Set objOperacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Sub

'' Obtem as mensagens passíveis de conciliação através de interação com a camada
'' de controle de caso de uso MIU, método A8Miu.clsMensagem.ObterDetalheMensagem,
'' e preenche o listview de mensagens com as mesmas.
Private Sub flCarregarListaMensagem(ByVal pintTipoOperacao As Integer, _
                                    ByVal pstrTipoOperacaoRotinaAbertura As String, _
                                    ByVal plngEmpresa As Long, _
                                    ByVal plngLocalLiquidacao As Long, _
                           Optional ByVal pblnVisaoPorQuantidade As Boolean = True)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim strRetLeituraQtdConciliada              As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomLeituraQtdConciliada              As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacaoRotinaAbertura", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacaoRotinaAbertura", _
                                     "TipoOperacaoRotinaAbertura", pstrTipoOperacaoRotinaAbertura)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    
    'Filtra as mensagens A CONCILIAR pela Data Atual em diante
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", _
                            fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
    '>>> --------------------------------------------------------------------------------------------------
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    If pblnVisaoPorQuantidade Then
        strRetLeituraQtdConciliada = objMensagem.ObterQtdConciliadaMensagem(vbNullString, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    End If
    
    strRetLeitura = objMensagem.ObterDetalheMensagem(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flLimparLista(Me.lstMensagem)
    
    Set objMensagem = Nothing
    
    If pblnVisaoPorQuantidade Then
        If strRetLeituraQtdConciliada <> vbNullString Then
            Set xmlDomLeituraQtdConciliada = CreateObject("MSXML2.DOMDocument.4.0")
            
            If Not xmlDomLeituraQtdConciliada.loadXML(strRetLeituraQtdConciliada) Then
                Call fgErroLoadXML(xmlDomLeituraQtdConciliada, App.EXEName, TypeName(Me), "flCarregarListaMensagem")
            End If
        End If
    
        If pintTipoOperacao <> INT_CODIGO_EVENTOS_TODOS Then
            lstMensagem.ColumnHeaders(COL_MSG_TIPO_ROTINA_ABER + 1).Width = 0
        Else
            lstMensagem.ColumnHeaders(COL_MSG_TIPO_ROTINA_ABER + 1).Width = 1200
        End If
        
        If pintTipoOperacao = enumTipoOperacaoLQS.EventosAmortização Or _
           pintTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
           pintTipoOperacao = enumTipoOperacaoLQS.EventosResgate Or _
           pintTipoOperacao = INT_CODIGO_EVENTOS_TODOS Then
            lstMensagem.ColumnHeaders(COL_MSG_NUMERO_COMANDO_ORIG + 1).Width = 0
        Else
            lstMensagem.ColumnHeaders(COL_MSG_NUMERO_COMANDO_ORIG + 1).Width = 1340
        End If
    End If

    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaMensagem")
        End If

        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            With lstMensagem.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                          objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & "|" & _
                          objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & "|" & _
                          objDomNode.selectSingleNode("CO_VEIC_LEGA").Text)

                If pblnVisaoPorQuantidade Then
                    .Tag = objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text

                    .SubItems(COL_MSG_VEICULO_LEGAL) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_MSG_CONTA_CUSTODIA) = objDomNode.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                    .SubItems(COL_MSG_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    .SubItems(COL_MSG_NUMERO_COMANDO_ORIG) = objDomNode.selectSingleNode("NU_COMD_OPER_ORIG").Text
                    .SubItems(COL_MSG_DATA_MENSAGEM) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                    .SubItems(COL_MSG_ID_ATIVO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text

                    .SubItems(COL_MSG_DT_VENC_ATIV) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_VENC").Text)

                    Select Case objDomNode.selectSingleNode("CO_OPER_SELIC").Text
                        Case enumCodigoOperacaoSelic.Resgate '"1012"
                            .SubItems(COL_MSG_TIPO_ROTINA_ABER) = "1012 - Resgate"
                        Case enumCodigoOperacaoSelic.Amortizacao '"1010"
                            .SubItems(COL_MSG_TIPO_ROTINA_ABER) = "1010 - Amortização"
                        Case enumCodigoOperacaoSelic.Juros '"1060"
                            .SubItems(COL_MSG_TIPO_ROTINA_ABER) = "1060 - Juros"
                    End Select
            
                    If objDomNode.selectSingleNode("QT_ATIV_MERC").Text <> vbNullString Then
                        .SubItems(COL_MSG_QTD_TOTAL) = objDomNode.selectSingleNode("QT_ATIV_MERC").Text
                    Else
                        .SubItems(COL_MSG_QTD_TOTAL) = 0
                    End If

                    'Verifica se o XML de Quantidade Conciliada foi carregado...
                    If Not xmlDomLeituraQtdConciliada Is Nothing Then
                        '...se sim, verifica se existe quantidade para a mensagem corrente...
                        If Not xmlDomLeituraQtdConciliada.selectSingleNode("Repeat_QtdConciliadaMensagem/Grupo_QtdConciliadaMensagem[NU_CTRL_IF='" & _
                            objDomNode.selectSingleNode("NU_CTRL_IF").Text & "' and DH_REGT_MESG_SPB='" & _
                            objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & "']/QT_ATIV_MERC_CNCL") Is Nothing Then

                            .SubItems(COL_MSG_QTD_CONCILIADA) = xmlDomLeituraQtdConciliada.selectSingleNode("Repeat_QtdConciliadaMensagem/Grupo_QtdConciliadaMensagem[NU_CTRL_IF='" & _
                                objDomNode.selectSingleNode("NU_CTRL_IF").Text & "' and DH_REGT_MESG_SPB='" & _
                                objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & "']/QT_ATIV_MERC_CNCL").Text

                        '...se não, carrega ZERO
                        Else
                            .SubItems(COL_MSG_QTD_CONCILIADA) = 0
                        End If

                    '...se não, carrega ZERO
                    Else
                        .SubItems(COL_MSG_QTD_CONCILIADA) = 0
                    End If

                    .SubItems(COL_MSG_QTD_A_CONCILIAR) = (.SubItems(COL_MSG_QTD_TOTAL) - .SubItems(COL_MSG_QTD_CONCILIADA))

                    .SubItems(COL_MSG_PU) = fgVlrXml_To_InterfaceDecimais(objDomNode.selectSingleNode("PU_ATIV_MERC").Text, 8)
                    .SubItems(COL_MSG_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                            
                Else
                    .Tag = objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                           objDomNode.selectSingleNode("VA_FINC").Text & "|" & _
                           objDomNode.selectSingleNode("NU_COMD_OPER").Text & "|" & _
                           objDomNode.selectSingleNode("CO_TEXT_XML").Text

                    .SubItems(COL_MSG2_CODIGO_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                    .SubItems(COL_MSG2_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                End If
            End With
        Next
    End If

    Call fgClassificarListview(Me.lstMensagem, lngIndexClassifListMesg, True)
    
    Set xmlDomLeitura = Nothing

ErrorHandler:

    Set objMensagem = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    Call fgCenterMe(Me)
    Call fgCursor(True)
    Set Me.Icon = mdiLQS.Icon
    Call flInicializar
    
    Call fgCarregarCombos(Me.cboLocalLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "SG_LOCA_LIQU")
    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    Call fgCarregarCombos(Me.cboTipoJustificativa, xmlMapaNavegacao, "TipoJustificativa", "TP_JUST_CNCL", "NO_TIPO_JUST_CNCL")
   
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoOperacao - Form_Load", Me.Caption

End Sub

'' Limpa o conteúdo dos listviews
Private Sub flLimparLista(ByVal lstListView As ListView)
    
    lstListView.ListItems.Clear

End Sub

'' Obtém os tipos de operação através de interação com a camada de controle de
'' caso de uso MIU.
Private Sub flCarregarTipoOperacaoConciliacao(plngLocalLiquidacao As Long)

    Call fgCarregarCombos(Me.cboTipoOperacao, xmlDomTipoOperacao, "TipoOperacaoConciliacao", "TP_OPER", "NO_TIPO_OPER", , , , "[CO_LOCA_LIQU='" & plngLocalLiquidacao & "']")
    
End Sub

'' Valida os campos a serem preenchidos pelo usuário no objeto. Retorna
'' vbNullString ou o error de consitência, se houver.
Private Function flValidarCampos(Optional ByVal pblnPorQuantidade As Boolean = True) As String

Dim strRetorno                              As String

    If lngLocalLiquidacao = enumLocalLiquidacao.BMA Then
        If fgItemsCheckedListView(lstOperacao) = 0 Then
            strRetorno = "Selecione ao menos um item de Operação para a conciliação."
        ElseIf fgItemsCheckedListView(lstMensagem) = 0 Then
            strRetorno = "Selecione ao menos um item de Mensagem para a conciliação."
        ElseIf fgItemsCheckedListView(lstMensagem) > 1 Then
            strRetorno = "Selecione somente um item de Mensagem para a conciliação."
        End If
    Else
        If pblnPorQuantidade Then
            'Operação NÃO selecionada e Mensagem NÃO selecionada
            If lngItemCheckedOperacao = 0 And lngItemCheckedMensagem = 0 Then
                strRetorno = "Selecione um item de operação e/ou mensagem para a conciliação."
            End If
        Else
            'Operação NÃO selecionada e Mensagem NÃO selecionada
            If lngItemCheckedOperacao = 0 And lngItemCheckedMensagem = 0 Then
                strRetorno = "Não existem itens de operação e/ou mensagem para a conciliação."
            End If
        End If
    End If
    
    flValidarCampos = strRetorno
    
End Function

'' Calcula e exibe a diferençaa entre os valores na conciliação.
Private Sub txtDiferenca_Change()

    chkConciliacaoFinalizada.Enabled = _
                    (Val(txtDiferenca.Text) <> 0 And _
                     lngItemCheckedOperacao > 0 And _
                     lngItemCheckedMensagem > 0 And _
                     (intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic Or _
                      intTipoOperacao <> enumTipoOperacaoLQS.EventosAmortização Or _
                      intTipoOperacao <> enumTipoOperacaoLQS.EventosJuros Or _
                      intTipoOperacao <> enumTipoOperacaoLQS.EventosResgate))

    chkConciliacaoFinalizada.value = IIf(chkConciliacaoFinalizada.Enabled, vbUnchecked, vbChecked)

End Sub

Private Sub flMostrarDiferencaConciliacao(Optional ByVal pblnVisaoPorQuantidade As Boolean = True)

Dim lngTotalOperacao                        As Long
Dim lngTotalMensagem                        As Long
Dim dblTotalOperacao                        As Double
Dim dblTotalMensagem                        As Double
Dim lngCont                                 As Long

On Error GoTo ErrorHandler

    lngItemCheckedOperacao = 0
    dblSomaValorOperacoes = 0
    With lstOperacao.ListItems
        If .Count <> 0 Then
            For lngCont = 1 To .Count
                If pblnVisaoPorQuantidade Then
                    If .Item(lngCont).Checked Then
                        lngTotalOperacao = lngTotalOperacao + _
                                           Val(.Item(lngCont).SubItems(COL_OP_QTD_A_CONCILIAR))
                        
                        lngItemCheckedOperacao = lngItemCheckedOperacao + 1
                        
                        dblSomaValorOperacoes = dblSomaValorOperacoes + _
                                           fgVlrXml_To_Decimal(.Item(lngCont).SubItems(COL_OP_VALOR))
                    End If
                Else
                    dblTotalOperacao = dblTotalOperacao + _
                                       fgVlrXml_To_Decimal(.Item(lngCont).Tag)
                    
                    lngItemCheckedOperacao = lngItemCheckedOperacao + 1
                End If
            Next
        End If
    End With
    
    lngItemCheckedMensagem = 0
    dblSomaValorMensagens = 0
    With lstMensagem.ListItems
        If .Count <> 0 Then
            For lngCont = 1 To .Count
                If pblnVisaoPorQuantidade Then
                    If .Item(lngCont).Checked Then
                        lngTotalMensagem = lngTotalMensagem + _
                                           Val(.Item(lngCont).SubItems(COL_MSG_QTD_A_CONCILIAR))
                        
                        lngItemCheckedMensagem = lngItemCheckedMensagem + 1
                    
                        dblSomaValorMensagens = dblSomaValorMensagens + _
                                           fgVlrXml_To_Decimal(.Item(lngCont).SubItems(COL_MSG_VALOR))
                    End If
                Else
                    dblTotalMensagem = dblTotalMensagem + _
                                       fgVlrXml_To_Decimal(Split(.Item(lngCont).Tag, "|")(POS_VALOR_MESG))
                    
                    lngItemCheckedMensagem = lngItemCheckedMensagem + 1
                End If
            Next
        End If
    End With
    
    If pblnVisaoPorQuantidade Then
        txtQtdVlrOperacao.Text = lngTotalOperacao
        txtQtdVlrMensagem.Text = lngTotalMensagem
        '---------------------------------------------------------------------------------
        'Força a chamada do método CHANGE para o controle << txtDiferenca >>,
        'mesmo se não ocorrer a mudança de conteúdo
        '---------------------------------------------------------------------------------
        txtDiferenca.Text = Abs(lngTotalOperacao - lngTotalMensagem)
        Call txtDiferenca_Change
        '---------------------------------------------------------------------------------
    Else
        txtQtdVlrOperacao.Text = fgVlrXml_To_Interface(dblTotalOperacao)
        txtQtdVlrMensagem.Text = fgVlrXml_To_Interface(dblTotalMensagem)
        txtDiferenca.Text = fgVlrXml_To_Interface(Abs(dblTotalOperacao - dblTotalMensagem))
    End If
    
Exit Sub
ErrorHandler:

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "flMostrarDiferencaConciliacao", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

'' Limpa o controle de justificativa
Private Sub flLimparControleJustificativa(Optional ByVal pblnVisaoPorQuantidade As Boolean = True)
    
    lngItemCheckedOperacao = 0
    lngItemCheckedMensagem = 0
    
    cboTipoJustificativa.ListIndex = -1
    txtComentario.Text = vbNullString
    
    If pblnVisaoPorQuantidade Then
        txtQtdVlrOperacao.Text = 0
        txtQtdVlrMensagem.Text = 0
        txtDiferenca.Text = 0
    Else
        txtQtdVlrOperacao.Text = fgVlrXml_To_Interface(0)
        txtQtdVlrMensagem.Text = txtQtdVlrOperacao.Text
        txtDiferenca.Text = txtQtdVlrOperacao.Text
    End If
End Sub

'' Exibe os erros de negócio que ocorreram na operação.
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

'' Configura os listviews para serem utilizados.
Private Sub flFormatarListas(Optional ByVal pblnPorQuantidade As Boolean = True)

    If pblnPorQuantidade Then
        lstOperacao.CheckBoxes = True
        lstMensagem.CheckBoxes = True
        
        With lstOperacao.ColumnHeaders
            .Clear
            
            .Add , , "Conciliar Operação", 1300
            .Add , , "Veículo Legal", 2000
            .Add , , "Conta Custódia SELIC", 2000
            .Add , , "Nro Comando", 1340
            .Add , , "Comando Original", 1340
            .Add , , "Tipo Operação", 1200
            .Add , , "ID Ativo", 1200
            .Add , , "Vencimento", 1100
            .Add , , "Qtde Total", 1340, lvwColumnRight
            .Add , , "Qtde a Conciliar", 1340, lvwColumnRight
            .Add , , "Qtde Conciliada", 1340, lvwColumnRight
            .Add , , "PU", 1300, lvwColumnRight
            .Add , , "Valor", 1440, lvwColumnRight
            .Add , , "Código Operação", 2000
            .Add , , "Data", 1100
    
        End With
        
        With lstMensagem.ColumnHeaders
            .Clear
            
            .Add , , "Conciliar Mensagem", 1300
            .Add , , "Veículo Legal", 2000
            .Add , , "Conta Custódia SELIC", 2000
            .Add , , "Nro Comando", 1340
            .Add , , "Comando Original", 1340
            .Add , , "Código Operação SELIC", 1200
            .Add , , "ID Ativo", 1200
            .Add , , "Vencimento", 1100
            .Add , , "Qtde Total", 1340, lvwColumnRight
            .Add , , "Qtde a Conciliar", 1340, lvwColumnRight
            .Add , , "Qtde Conciliada", 1340, lvwColumnRight
            .Add , , "PU", 1300, lvwColumnRight
            .Add , , "Valor", 1440, lvwColumnRight
            .Add , , "Data", 1100
    
        End With
        
        fraResumo.Caption = "Quantidades A Conciliar"
    
    '...se não, Por Valor
    Else
        lstOperacao.CheckBoxes = False
        lstMensagem.CheckBoxes = False
        
        With lstOperacao.ColumnHeaders
            .Clear
            
            .Add , , "Conciliar Operação", 1600
            .Add , , "Grupo Veículo Legal", 3000
            .Add , , "Valor", 1440
            .Item(COL_OP2_VALOR + 1).Alignment = lvwColumnRight
        End With

        With lstMensagem.ColumnHeaders
            .Clear
            
            .Add , , "Conciliar Mensagem", 1600
            .Add , , "Mensagem", 1440
            .Add , , "Valor", 1440
            .Item(COL_MSG2_VALOR + 1).Alignment = lvwColumnRight
        End With
        
        fraResumo.Caption = "Valores A Conciliar"
    
    End If
    
End Sub

'' Gera o XML que será enviado para a camada MIU
Private Function flMontarXMLConciliacao(Optional ByVal pblnPorQuantidade As Boolean = True) As String

Dim xmlDomConciliacao                       As MSXML2.DOMDocument40
Dim lngCont                                 As Long
Dim intTipoOperacao                         As Integer

    'Monta XML para a conciliação
    Set xmlDomConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomConciliacao, "", "Repeat_Conciliacao", "")
    Call fgAppendNode(xmlDomConciliacao, "Repeat_Conciliacao", "Grupo_Conciliacao", "")
    
    intTipoOperacao = fgObterCodigoCombo(Me.cboTipoOperacao)
    Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "TipoOperacao", intTipoOperacao)
    Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Empresa", fgObterCodigoCombo(Me.cboEmpresa))
    Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeItemSelOperacao", lngItemCheckedOperacao)
    Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeItemSelMensagem", lngItemCheckedMensagem)
    
    If lngItemCheckedMensagem > 1 Then
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ConciliacaoFinalizada", chkConciliacaoFinalizada.value)
    End If
    
    If pblnPorQuantidade Then
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeTotalOperacao", Val(txtQtdVlrOperacao.Text))
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeTotalMensagem", Val(txtQtdVlrMensagem.Text))
        
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ValorTotalOperacao", fgVlr_To_Xml(dblSomaValorOperacoes))
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ValorTotalMensagem", fgVlr_To_Xml(dblSomaValorMensagens))
        
        'Verifica se é uma conciliação 1 MSG -> N OP...                 (Sem divergência de quantidades)
        If lngItemCheckedOperacao > 1 Then
            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PontaN", "O")
            
            With lstMensagem.ListItems
                For lngCont = 1 To .Count
                    If .Item(lngCont).Checked Then
                        
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "NumeroControleIF", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataRegistroMensagemSPB", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MESG_SPB))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataUltimaAtualizacaoMensagem", _
                                                .Item(lngCont).Tag)
                        
                        If lngLocalLiquidacao = enumLocalLiquidacao.BMA Then
                            'Só altera status da mensagem se zerar a quantidade. Senao deixa 'a conciliar' para continuar na tela
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", IIf(Val(txtQtdVlrOperacao.Text) = Val(txtQtdVlrMensagem.Text), "1", "0"))
                        Else
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", 1)
                        End If
                        
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "VeiculoLegal", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_MSG))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Titulo", _
                                                .Item(lngCont).SubItems(COL_MSG_ID_ATIVO))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Vencimento", _
                                                fgDate_To_DtXML(.Item(lngCont).SubItems(COL_MSG_DT_VENC_ATIV)))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comando", _
                                                IIf(Trim(.Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO)) = vbNullString, "0", .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO)))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ComandoOriginal", _
                                                IIf(Trim(.Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG)) = vbNullString, "0", .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG)))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PU", _
                                                .Item(lngCont).SubItems(COL_MSG_PU))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ContaCustodia", _
                                                Mid(fgCompletaString(.Item(lngCont).SubItems(COL_MSG_CONTA_CUSTODIA), "0", 9, True), 1, 4))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "TipoRotinaAbertura", _
                                                Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4))

                        Select Case Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4)
                            Case enumCodigoOperacaoSelic.Resgate '"1012"
                                xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosResgate
                            Case enumCodigoOperacaoSelic.Amortizacao '"1010"
                                xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosAmortização
                            Case enumCodigoOperacaoSelic.Juros ' "1060"
                                xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosJuros
                        End Select
                        
                        Exit For
                    
                    End If
                Next
            End With
            
            With lstOperacao.ListItems
                For lngCont = 1 To .Count
                    If .Item(lngCont).Checked Then
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Repeat_Operacao", "", _
                                                "Repeat_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "Operacao", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NU_SEQU_OPER_ATIV), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "AlteraStatusOperacao", 1, _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "DataUltimaAtualizacaoOperacao", _
                                                .Item(lngCont).Tag, "Grupo_Conciliacao")
                        
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "VeiculoLegal", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_OPER), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "QuantidadeConciliada", _
                                                .Item(lngCont).SubItems(COL_OP_QTD_A_CONCILIAR), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "Titulo", _
                                                .Item(lngCont).SubItems(COL_OP_ID_ATIVO), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "Vencimento", _
                                                fgDate_To_DtXML(.Item(lngCont).SubItems(COL_OP_DT_VENC_ATIV)), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "Comando", _
                                                .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "ComandoOriginal", _
                                                .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO_ORIG), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "PU", _
                                                .Item(lngCont).SubItems(COL_OP_PU), "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Operacao", "ContaCustodia", _
                                                Mid(fgCompletaString(.Item(lngCont).SubItems(COL_OP_CONTA_CUSTODIA), "0", 9, True), 1, 4), "Grupo_Conciliacao")
                    
                        xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = Split(Mid$(.Item(lngCont).Key, 2), "|")(POS_TP_OPER)
                    End If
                Next
            End With
        
        '...se não, verifica se é uma conciliação 1 OP -> N MSG...  (Sem divergência de quantidades)
        ElseIf lngItemCheckedMensagem > 1 Then
            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PontaN", "M")
            
            With lstOperacao.ListItems
                For lngCont = 1 To .Count
                    If .Item(lngCont).Checked Then
                        
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Operacao", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NU_SEQU_OPER_ATIV))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusOperacao", 1)
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataUltimaAtualizacaoOperacao", _
                                                .Item(lngCont).Tag)
                        
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "VeiculoLegal", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_OPER))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeConciliada", _
                                                .Item(lngCont).SubItems(COL_OP_QTD_A_CONCILIAR))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Titulo", _
                                                .Item(lngCont).SubItems(COL_OP_ID_ATIVO))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Vencimento", _
                                                fgDate_To_DtXML(.Item(lngCont).SubItems(COL_OP_DT_VENC_ATIV)))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comando", _
                                                .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ComandoOriginal", _
                                                .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO_ORIG))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PU", _
                                                .Item(lngCont).SubItems(COL_OP_PU))
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ContaCustodia", _
                                                Mid(fgCompletaString(.Item(lngCont).SubItems(COL_OP_CONTA_CUSTODIA), "0", 9, True), 1, 4))
                    
                        xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = Split(Mid$(.Item(lngCont).Key, 2), "|")(POS_TP_OPER)
                            
                        Exit For
                    
                    End If
                Next
            End With
            
            With lstMensagem.ListItems
                For lngCont = 1 To .Count
                    If .Item(lngCont).Checked Then
                        
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Repeat_Mensagem", "", _
                                                "Repeat_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "NumeroControleIF", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "DataRegistroMensagemSPB", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MESG_SPB), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "AlteraStatusMensagem", 1, _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "DataUltimaAtualizacaoMensagem", _
                                                .Item(lngCont).Tag, _
                                                "Grupo_Conciliacao")
                        
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "VeiculoLegal", _
                                                Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_MSG), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "QuantidadeConciliada", _
                                                .Item(lngCont).SubItems(COL_MSG_QTD_A_CONCILIAR), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "Titulo", _
                                                .Item(lngCont).SubItems(COL_MSG_ID_ATIVO), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "Vencimento", _
                                                fgDate_To_DtXML(.Item(lngCont).SubItems(COL_MSG_DT_VENC_ATIV)), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "Comando", _
                                                .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "ComandoOriginal", _
                                                .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "PU", _
                                                .Item(lngCont).SubItems(COL_MSG_PU), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "ContaCustodia", _
                                                Mid(fgCompletaString(.Item(lngCont).SubItems(COL_MSG_CONTA_CUSTODIA), "0", 9, True), 1, 4), _
                                                "Grupo_Conciliacao")
                        Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "TipoRotinaAbertura", _
                                                Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4), _
                                                "Grupo_Conciliacao")
                        
                        Select Case Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4)
                            Case enumCodigoOperacaoSelic.Resgate '"1012"
                                xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosResgate
                            Case enumCodigoOperacaoSelic.Amortizacao ' "1010"
                                xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosAmortização
                            Case enumCodigoOperacaoSelic.Juros '"1060"
                                xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosJuros
                        End Select
                        
                    End If
                Next
            End With
            
        Else
            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PontaN", vbNullString)
            
            'Verifica se é uma conciliação 1 OP -> 1 MSG...
            If lngItemCheckedOperacao = 1 And lngItemCheckedMensagem = 1 Then
                Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeConciliada", _
                                                        fgMenorValor(Val(txtQtdVlrOperacao.Text), Val(txtQtdVlrMensagem.Text)))
                
                'Verifica se as quantidades estão batidas...
                If Val(txtDiferenca.Text) = 0 Then
                    Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusOperacao", 1)
                    Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", 1)
                
                '...se não, verifica se foi forçado o encerramento da conciliação...
                Else
                    If chkConciliacaoFinalizada.value = vbChecked Then
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusOperacao", 1)
                        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", 1)
                    
                    '...se não, verifica qual ponta possui o menor valor para o controle da mudança de status
                    Else
                        If fgMenorValor(Val(txtQtdVlrOperacao.Text), Val(txtQtdVlrMensagem.Text)) = Val(txtQtdVlrOperacao.Text) Then
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusOperacao", 1)
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", 0)
                        Else
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusOperacao", 0)
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", 1)
                        End If
                    End If
                End If
                
                With lstOperacao.ListItems
                    For lngCont = 1 To .Count
                        If .Item(lngCont).Checked Then
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Operacao", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NU_SEQU_OPER_ATIV))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataUltimaAtualizacaoOperacao", _
                                                    .Item(lngCont).Tag)
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "VeiculoLegal", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_OPER))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Titulo", _
                                                    .Item(lngCont).SubItems(COL_OP_ID_ATIVO))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Vencimento", _
                                                    fgDate_To_DtXML(.Item(lngCont).SubItems(COL_OP_DT_VENC_ATIV)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comando", _
                                                    .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ComandoOriginal", _
                                                    .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO_ORIG))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PU", _
                                                    .Item(lngCont).SubItems(COL_OP_PU))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ContaCustodia", _
                                                    Mid(fgCompletaString(.Item(lngCont).SubItems(COL_OP_CONTA_CUSTODIA), "0", 9, True), 1, 4))
                        
                            xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = Split(Mid$(.Item(lngCont).Key, 2), "|")(POS_TP_OPER)
                                
                            Exit For
                        
                        End If
                    Next
                End With
                
                With lstMensagem.ListItems
                    For lngCont = 1 To .Count
                        If .Item(lngCont).Checked Then
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "NumeroControleIF", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataRegistroMensagemSPB", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MESG_SPB))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataUltimaAtualizacaoMensagem", _
                                                    .Item(lngCont).Tag)
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "VeiculoLegal", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_MSG))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Titulo", _
                                                    .Item(lngCont).SubItems(COL_MSG_ID_ATIVO))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Vencimento", _
                                                    fgDate_To_DtXML(.Item(lngCont).SubItems(COL_MSG_DT_VENC_ATIV)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comando", _
                                                    IIf(Trim(.Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO)) = vbNullString, "0", .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ComandoOriginal", _
                                                    IIf(Trim(.Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG)) = vbNullString, "0", .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PU", _
                                                    .Item(lngCont).SubItems(COL_MSG_PU))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ContaCustodia", _
                                                    Mid(fgCompletaString(.Item(lngCont).SubItems(COL_MSG_CONTA_CUSTODIA), "0", 9, True), 1, 4))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "TipoRotinaAbertura", _
                                                    Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4))
                            
                            Select Case Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4)
                                Case enumCodigoOperacaoSelic.Resgate '"1012"
                                    xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosResgate
                                Case enumCodigoOperacaoSelic.Amortizacao '"1010"
                                    xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosAmortização
                                Case enumCodigoOperacaoSelic.Juros '"1060"
                                    xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosJuros
                            End Select
                            
                            Exit For
                        
                        End If
                    Next
                End With
            
            '...se não, verifica se é uma conciliação 1 OP -> 0 MSG...
            ElseIf lngItemCheckedOperacao > 0 And lngItemCheckedMensagem = 0 Then
                With lstOperacao.ListItems
                    For lngCont = 1 To .Count
                        If .Item(lngCont).Checked Then
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Operacao", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NU_SEQU_OPER_ATIV))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusOperacao", 1)
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeConciliada", _
                                                    Val(txtQtdVlrOperacao.Text))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataUltimaAtualizacaoOperacao", _
                                                    .Item(lngCont).Tag)
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "VeiculoLegal", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_OPER))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Titulo", _
                                                    .Item(lngCont).SubItems(COL_OP_ID_ATIVO))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Vencimento", _
                                                    fgDate_To_DtXML(.Item(lngCont).SubItems(COL_OP_DT_VENC_ATIV)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comando", _
                                                    .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ComandoOriginal", _
                                                    .Item(lngCont).SubItems(COL_OP_NUMERO_COMANDO_ORIG))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PU", _
                                                    .Item(lngCont).SubItems(COL_OP_PU))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ContaCustodia", _
                                                    Mid(fgCompletaString(.Item(lngCont).SubItems(COL_OP_CONTA_CUSTODIA), "0", 9, True), 1, 4))
                        
                            xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = Split(Mid$(.Item(lngCont).Key, 2), "|")(POS_TP_OPER)
                                
                            Exit For
                        
                        End If
                    Next
                End With
                
            '...se não, verifica se é uma conciliação 1 MSG -> 0 OP...
            ElseIf lngItemCheckedMensagem > 0 And lngItemCheckedOperacao = 0 Then
                With lstMensagem.ListItems
                    For lngCont = 1 To .Count
                        If .Item(lngCont).Checked Then
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "NumeroControleIF", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataRegistroMensagemSPB", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MESG_SPB))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "AlteraStatusMensagem", 1)
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "QuantidadeConciliada", _
                                                    Val(txtQtdVlrMensagem.Text))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "DataUltimaAtualizacaoMensagem", _
                                                    .Item(lngCont).Tag)
                            
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "VeiculoLegal", _
                                                    Split(Mid(.Item(lngCont).Key, 2), "|")(POS_CO_VEIC_LEGA_MSG))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Titulo", _
                                                    .Item(lngCont).SubItems(COL_MSG_ID_ATIVO))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Vencimento", _
                                                    fgDate_To_DtXML(.Item(lngCont).SubItems(COL_MSG_DT_VENC_ATIV)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comando", _
                                                    IIf(Trim(.Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO)) = vbNullString, "0", .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ComandoOriginal", _
                                                    IIf(Trim(.Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG)) = vbNullString, "0", .Item(lngCont).SubItems(COL_MSG_NUMERO_COMANDO_ORIG)))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PU", _
                                                    .Item(lngCont).SubItems(COL_MSG_PU))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ContaCustodia", _
                                                    Mid(fgCompletaString(.Item(lngCont).SubItems(COL_MSG_CONTA_CUSTODIA), "0", 9, True), 1, 4))
                            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "TipoRotinaAbertura", _
                                                    Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4))
                            
                            Select Case Left$(.Item(lngCont).SubItems(COL_MSG_TIPO_ROTINA_ABER), 4)
                                Case enumCodigoOperacaoSelic.Resgate '"1012"
                                    xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosResgate
                                Case enumCodigoOperacaoSelic.Amortizacao ' "1010"
                                    xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosAmortização
                                Case enumCodigoOperacaoSelic.Juros '"1060"
                                    xmlDomConciliacao.selectSingleNode("//TipoOperacao").Text = enumTipoOperacaoLQS.EventosJuros
                            End Select
                            
                            Exit For
                        
                        End If
                    Next
                End With
                
            End If
        End If
        
    '...se não, monta por VALOR (específico para o caso de DESPESAS)
    Else
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "PontaN", "MO")       'N MSG -> N OP
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ValorTotalOperacao", Val(txtQtdVlrOperacao.Text))
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ValorTotalMensagem", Val(txtQtdVlrMensagem.Text))
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "ValorDiferenca", fgVlrXml_To_Decimal(txtDiferenca.Text))

        With lstMensagem.ListItems
            For lngCont = 1 To .Count
                'Monta a ponta das mensagens de acordo com o LISTVIEW de mensagens
                Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Repeat_Mensagem", "")
                Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "NumeroControleIF", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF), _
                                        "Grupo_Conciliacao")
                Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "DataRegistroMensagemSPB", _
                                        Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MESG_SPB), _
                                        "Grupo_Conciliacao")
                Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "DataUltimaAtualizacaoMensagem", _
                                        Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATLZ_MESG), _
                                        "Grupo_Conciliacao")
                Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "NumeroComando", _
                                        Split(.Item(lngCont).Tag, "|")(POS_NUMERO_COMANDO), _
                                        "Grupo_Conciliacao")
                Call fgAppendNode(xmlDomConciliacao, "Repeat_Mensagem", "TextoXML", _
                                        Split(.Item(lngCont).Tag, "|")(POS_TEXTO_XML), _
                                        "Grupo_Conciliacao")
            Next
        End With
        
        'Adiciona o XML de apoio contendo OPERAÇÕES e TOTAIS        (ponta das operações)
        If xmlDomOperacao.xml = vbNullString Then
            Call fgAppendNode(xmlDomConciliacao, "Repeat_Conciliacao", "ConciliacaoOperacao", "")
        Else
            Call fgAppendXML(xmlDomConciliacao, "Repeat_Conciliacao", xmlDomOperacao.xml)
        End If
    End If
    
    If cboTipoJustificativa.ListIndex <> -1 Then
        Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "TipoJustificativa", _
                            fgObterCodigoCombo(Me.cboTipoJustificativa))
        
        If Trim(txtComentario.Text) <> vbNullString Then
            Call fgAppendNode(xmlDomConciliacao, "Grupo_Conciliacao", "Comentario", _
                                txtComentario.Text)
        End If
    End If
    
    flMontarXMLConciliacao = xmlDomConciliacao.xml
    
    Set xmlDomConciliacao = Nothing

End Function

'' Exibe o Resultado da despesa
Private Sub flApresentaResultDespesa(ByVal pstrRetorno As String)

    With frmResultOperacaoLote
        .ApresentaInfo = True
        .Resultado = pstrRetorno
        .Show vbModal
    End With

End Sub

Private Sub flMontaTela()

Dim lngEmpresa                              As Long
Dim strTipoOperacaoRotinaAbertura           As String

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    If cboTipoOperacao.ListIndex <> -1 _
        And cboEmpresa.ListIndex <> -1 _
        And cboLocalLiquidacao.ListIndex <> -1 Then

        Call flFormatarListas(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)
        Call flLimparControleJustificativa(intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic)

        lngEmpresa = fgObterCodigoCombo(Me.cboEmpresa)
        lngLocalLiquidacao = fgObterCodigoCombo(Me.cboLocalLiquidacao)
        
        If intTipoOperacao <> INT_CODIGO_EVENTOS_TODOS Then
            strTipoOperacaoRotinaAbertura = xmlDomTipoOperacao.selectSingleNode("Repeat_TipoOperacaoConciliacao/Grupo_TipoOperacaoConciliacao[TP_OPER='" & intTipoOperacao & "']/CO_OPER_SELIC").Text
        Else
            'Tipo Operação Rotina Abertura fixo para eventos, de acordo com a SEL1611
            'strTipoOperacaoRotinaAbertura = "1012, 1010, 1060"
            strTipoOperacaoRotinaAbertura = enumCodigoOperacaoSelic.Resgate & "," & _
                                            enumCodigoOperacaoSelic.Amortizacao & "," & _
                                            enumCodigoOperacaoSelic.Juros
        End If

        Call flCarregarListaOperacao(intTipoOperacao, _
                                     lngEmpresa, _
                                     lngLocalLiquidacao, _
                                     (intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic))
                                     
        Call flCarregarListaMensagem(intTipoOperacao, _
                                     strTipoOperacaoRotinaAbertura, _
                                     lngEmpresa, _
                                     lngLocalLiquidacao, _
                                     (intTipoOperacao <> enumTipoOperacaoLQS.DespesaSelic))
                                     
        If intTipoOperacao = enumTipoOperacaoLQS.DespesaSelic Then
            Call flMostrarDiferencaConciliacao(False)
        End If

    End If
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConciliacaoOperacao - flMontaTela", Me.Caption
    
End Sub

Private Sub flSugestaoCombo(cbo As ComboBox)
    
    'Se o combobox só tiver um item, seleciona-o
    'Se tiver mais, coloca o foco nele e já o abre, usando ALT+DOWN

    cbo.SetFocus
    If cbo.ListCount = 1 Then
        cbo.ListIndex = 0
    ElseIf cbo.ListCount > 1 Then
        SendKeys "%{DOWN}"
    End If

End Sub

Private Sub flRecebeRemessa()

Dim r As Object
    
    Set r = CreateObject("A8LQS.clsRemessa")
    
    r.ReceberMensagemMQ "!", "q", "", 1, "1"

End Sub
