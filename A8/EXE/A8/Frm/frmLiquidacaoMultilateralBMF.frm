VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiquidacaoMultilateralBMF 
   Caption         =   "Liquidação Multilateral BMF"
   ClientHeight    =   8625
   ClientLeft      =   975
   ClientTop       =   855
   ClientWidth     =   12975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   12975
   WindowState     =   2  'Maximized
   Begin VB.Frame fraComplementos 
      Height          =   765
      Left            =   60
      TabIndex        =   5
      Top             =   7530
      Width           =   12855
      Begin VB.TextBox txtJustificativa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   6
         Top             =   270
         Width           =   8595
      End
      Begin VB.Label lblJustificativa 
         AutoSize        =   -1  'True
         Caption         =   "Justificativa"
         Height          =   195
         Left            =   510
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.TextBox txtValorMensagem 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6930
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   300
      Width           =   2415
   End
   Begin VB.Frame fraTipoMensagem 
      Caption         =   "Tipo de Mensagem "
      Height          =   570
      Left            =   4500
      TabIndex        =   8
      Top             =   60
      Width           =   2325
      Begin VB.OptionButton optTipoMensagem 
         Caption         =   "Prévia"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTipoMensagem 
         Caption         =   "Definitiva"
         Height          =   195
         Index           =   1
         Left            =   1170
         TabIndex        =   9
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4350
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   10680
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
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralBMF.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwNet 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5741
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
      NumItems        =   0
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   11310
      Top             =   0
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   8295
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   582
      ButtonWidth     =   2566
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela   "
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Concordar      "
            Key             =   "concordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Discordar        "
            Key             =   "discordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Rejeitar          "
            Key             =   "retorno"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Liberar            "
            Key             =   "liberacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pg. Conting.   "
            Key             =   "pagamentocontingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Regularizar     "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetalhe 
      Height          =   2745
      Left            =   60
      TabIndex        =   1
      Top             =   4140
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   4842
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
   Begin VB.Label lblValorMensagem 
      AutoSize        =   -1  'True
      Caption         =   "Valor Mensagem"
      Height          =   195
      Left            =   6960
      TabIndex        =   11
      Top             =   60
      Width           =   1185
   End
   Begin VB.Image imgDummyH 
      Height          =   60
      Left            =   60
      MousePointer    =   7  'Size N S
      Top             =   3990
      Width           =   14040
   End
   Begin VB.Label lblConciliacao 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmLiquidacaoMultilateralBMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Liquidação Multilateral BMF Derivativa

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40
Private xmlArquivoCamara                    As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Net de Operações - Backoffice e Adm.Área
Private Const COL_BOAA_NET_AGEN_COMP        As Integer = 0
Private Const COL_BOAA_NET_VEIC_LEGA        As Integer = 1
Private Const COL_BOAA_NET_VALR_ACON        As Integer = 2
Private Const COL_BOAA_NET_VALR_CONC        As Integer = 3
Private Const COL_BOAA_NET_TOTL_SIST        As Integer = 4
Private Const COL_BOAA_NET_VALR_CAMR        As Integer = 5
Private Const COL_BOAA_NET_DIFE_VALR        As Integer = 6

'Constantes de Configuração de Colunas de Net de Operações - Adm.Geral - Prévia
Private Const COL_AGPV_NET_AGEN_COMP        As Integer = 0
Private Const COL_AGPV_NET_VALR_SIST        As Integer = 1
Private Const COL_AGPV_NET_VALR_PREV        As Integer = 2
Private Const COL_AGPV_NET_DIFE_VALR        As Integer = 3
Private Const COL_AGPV_NET_VLDL_0001        As Integer = 4
Private Const COL_AGPV_NET_VLDL_1001        As Integer = 5

'Constantes de Configuração de Colunas de Net de Operações - Adm.Geral - Definitiva
Private Const COL_AGDF_NET_AGEN_COMP        As Integer = 0
Private Const COL_AGDF_NET_VALR_SIST        As Integer = 1
Private Const COL_AGDF_NET_VALR_DEFI        As Integer = 2
Private Const COL_AGDF_NET_DIFE_VALR        As Integer = 3
Private Const COL_AGDF_NET_VLDL_0004        As Integer = 4
Private Const COL_AGDF_NET_VLDL_PGRC        As Integer = 5

'Constantes de posicionamento de campos na propriedade Key do item de Net de Operações - Backoffice e Adm.Área
Private Const KEY_BOAA_NET_AGEN_COMP        As Integer = 1
Private Const KEY_BOAA_NET_VEIC_LEGA        As Integer = 2
Private Const KEY_BOAA_NET_TIPO_BKOF        As Integer = 3

'Constantes de posicionamento de campos na propriedade Key do item de Net de Operações - Adm.Geral
Private Const KEY_ADMG_NET_AGEN_COMP        As Integer = 1

'Constantes de posicionamento de campos na propriedade Tag do item de Net de Operações
Private Const TAG_MSG_NU_CTRL_IF               As Integer = 1
Private Const TAG_MSG_DH_REGT_MESG_SPB         As Integer = 2
Private Const TAG_MSG_NU_SEQU_CNTR_REPE        As Integer = 3
Private Const TAG_MSG_DH_ULTI_ATLZ             As Integer = 4
Private Const TAG_MSG_CO_ULTI_SITU_PROC        As Integer = 5
Private Const TAG_MSG_TP_ACAO_MESG_SPB_EXEC    As Integer = 6
Private Const TAG_MSG_IN_DEBT_CRED             As Integer = 7
Private Const TAG_MSG_NU_CTRL_CAMR             As Integer = 8
Private Const TAG_MSG_NU_CTRL_IF_LDL1001       As Integer = 9
Private Const TAG_MSG_DH_REGT_MESG_SPB_LDL1001 As Integer = 10

'Constantes de Configuração de Colunas de Detalhes de Operação - Backoffice e Adm.Área
Private Const COL_BOAA_DET_CODI_LANC        As Integer = 0
Private Const COL_BOAA_DET_CODI_OPER        As Integer = 1
Private Const COL_BOAA_DET_DEBT_CRED        As Integer = 2
Private Const COL_BOAA_DET_VALR_SIST        As Integer = 3
Private Const COL_BOAA_DET_STAT_OPER        As Integer = 4

'Constantes de Configuração de Colunas de Detalhes de Operação - Adm.Geral
Private Const COL_ADMG_DET_AREA_RESP        As Integer = 0
Private Const COL_ADMG_DET_VALR_SIST        As Integer = 1
Private Const COL_ADMG_DET_VALR_CAMR        As Integer = 2
Private Const COL_ADMG_DET_DIFE_VALR        As Integer = 3
Private Const COL_ADMG_DET_JUST_DIVG        As Integer = 4

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const KEY_DET_NU_SEQU_OPER_ATIV     As Integer = 1

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações Por Área
Private Const KEY_ADMG_DET_TIPO_BKOF        As Integer = 1

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmLiquidacaoMultilateralBMF"

'Constantes de Strings utilizadas pelo Form
Private Const strChaveTotais                As String = "|Totais"
'------------------------------------------------------------------------------------------
'Fim declaração constantes

Private Enum enumValoresCalculados
    AConcordar = 0
    Concordado = 1
    SistemaOrigem = 2
    Camara = 3
End Enum

Private Enum enumPreviaDefinitiva
    Previa = 0
    Definitiva = 1
End Enum

Private intAcaoProcessamento                As enumAcaoConciliacao

Private lngPerfil                           As Long
Private blnDummyH                           As Boolean

Private lngIndexClassifListNet              As Long
Private lngIndexClassifListDet              As Long

'Calcula diferenças entre colunas do listview
Private Sub flCalcularDiferencasListViewDetalhe()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

On Error GoTo ErrorHandler

    For Each objListItem In lvwDetalhe.ListItems
        With objListItem
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_DET_VALR_SIST)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_DET_VALR_CAMR)))

            .SubItems(COL_ADMG_DET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorOperacao - dblValorMensagem)

            If dblValorOperacao - dblValorMensagem <> 0 Then
                .ListSubItems(COL_ADMG_DET_DIFE_VALR).ForeColor = vbRed
            End If
        End With
    Next

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListViewDetalhe", 0)

End Sub

'Calcula diferenças entre colunas do listview
Private Sub flCalcularDiferencasListViewNet()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorCamara                          As Double

On Error GoTo ErrorHandler

    For Each objListItem In lvwNet.ListItems
        With objListItem
            Select Case PerfilAcesso
                Case enumPerfilAcesso.AdmGeralPrevia
                
                    dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VALR_SIST)))
                    dblValorCamara = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VALR_PREV)))
    
                    .SubItems(COL_AGPV_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorCamara - dblValorOperacao)
    
                    If dblValorCamara - dblValorOperacao <> 0 Then
                        .ListSubItems(COL_AGPV_NET_DIFE_VALR).ForeColor = vbRed
                    End If
            
                Case enumPerfilAcesso.AdmGeral
                
                    dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VALR_SIST)))
                    dblValorCamara = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VALR_DEFI)))
    
                    .SubItems(COL_AGDF_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorCamara - dblValorOperacao)
    
                    If dblValorCamara - dblValorOperacao <> 0 Then
                        .ListSubItems(COL_AGDF_NET_DIFE_VALR).ForeColor = vbRed
                    End If
            
                Case Else
                
                    dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_TOTL_SIST)))
                    dblValorCamara = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_CAMR)))
    
                    .SubItems(COL_BOAA_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorCamara - dblValorOperacao)
    
                    If dblValorCamara - dblValorOperacao <> 0 Then
                        .ListSubItems(COL_BOAA_NET_DIFE_VALR).ForeColor = vbRed
                    End If
            
            End Select

        End With
    Next

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListViewNet", 0)

End Sub

'Calcula Totais de Valores e os exibe na última linha do Listview
Private Sub flCalcularTotais()

Dim objListItem                             As ListItem

Dim dblAConcordar                           As Double
Dim dblConcordado                           As Double
Dim dblArquivo                              As Double
Dim dblSistemaOrigem                        As Double
Dim dblDiferenca                            As Double
Dim dblLDL0001                              As Double
Dim dblLDL1001                              As Double
Dim dblLDL0004                              As Double
Dim dblPagarReceber                         As Double

On Error GoTo ErrorHandler

    dblAConcordar = 0
    dblConcordado = 0
    dblArquivo = 0
    dblSistemaOrigem = 0
    dblDiferenca = 0
    dblLDL0001 = 0
    dblLDL1001 = 0
    dblLDL0004 = 0
    dblPagarReceber = 0
    
    For Each objListItem In lvwNet.ListItems
        With objListItem
            If .Key <> strChaveTotais Then
                Select Case PerfilAcesso
                    Case enumPerfilAcesso.AdmGeralPrevia
                        
                        dblSistemaOrigem = dblSistemaOrigem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VALR_SIST)))
                        dblArquivo = dblArquivo + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VALR_PREV)))
                        dblDiferenca = dblDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_DIFE_VALR)))
                        dblLDL0001 = dblLDL0001 + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))
                        dblLDL1001 = dblLDL1001 + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))
    
                    Case enumPerfilAcesso.AdmGeral
                        
                        dblSistemaOrigem = dblSistemaOrigem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VALR_SIST)))
                        dblArquivo = dblArquivo + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VALR_DEFI)))
                        dblDiferenca = dblDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_DIFE_VALR)))
                        dblLDL0004 = dblLDL0004 + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VLDL_0004)))
                        dblPagarReceber = dblPagarReceber + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VLDL_PGRC)))
    
                    Case Else
                        
                        dblAConcordar = dblAConcordar + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_ACON)))
                        dblConcordado = dblConcordado + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_CONC)))
                        dblSistemaOrigem = dblSistemaOrigem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_TOTL_SIST)))
                        dblArquivo = dblArquivo + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_CAMR)))
                        dblDiferenca = dblDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_DIFE_VALR)))
                
                End Select
            End If
        End With
    Next

    If Not fgExisteItemLvw(Me.lvwNet, strChaveTotais) Then
        Set objListItem = lvwNet.ListItems.Add(, strChaveTotais)
    Else
        Set objListItem = lvwNet.ListItems(strChaveTotais)
    End If
    
    With objListItem
        .Text = "Totais"
        .Bold = True
        .Tag = strChaveTotais
        
        Select Case PerfilAcesso
            Case enumPerfilAcesso.AdmGeralPrevia
                
                .SubItems(COL_AGPV_NET_VALR_SIST) = fgVlrXml_To_Interface(dblSistemaOrigem)
                .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface(dblArquivo)
                .SubItems(COL_AGPV_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblDiferenca)
                .SubItems(COL_AGPV_NET_VLDL_0001) = fgVlrXml_To_Interface(dblLDL0001)
                .SubItems(COL_AGPV_NET_VLDL_1001) = fgVlrXml_To_Interface(dblLDL1001)
                
                .ListSubItems(COL_AGPV_NET_VALR_SIST).Bold = True
                .ListSubItems(COL_AGPV_NET_VALR_PREV).Bold = True
                .ListSubItems(COL_AGPV_NET_DIFE_VALR).Bold = True
                .ListSubItems(COL_AGPV_NET_VLDL_0001).Bold = True
                .ListSubItems(COL_AGPV_NET_VLDL_1001).Bold = True
        
            Case enumPerfilAcesso.AdmGeral
                
                .SubItems(COL_AGDF_NET_VALR_SIST) = fgVlrXml_To_Interface(dblSistemaOrigem)
                .SubItems(COL_AGDF_NET_VALR_DEFI) = fgVlrXml_To_Interface(dblArquivo)
                .SubItems(COL_AGDF_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblDiferenca)
                .SubItems(COL_AGDF_NET_VLDL_0004) = fgVlrXml_To_Interface(dblLDL0004)
                .SubItems(COL_AGDF_NET_VLDL_PGRC) = fgVlrXml_To_Interface(dblPagarReceber)
                
                .ListSubItems(COL_AGPV_NET_VALR_SIST).Bold = True
                .ListSubItems(COL_AGPV_NET_VALR_PREV).Bold = True
                .ListSubItems(COL_AGPV_NET_DIFE_VALR).Bold = True
                .ListSubItems(COL_AGDF_NET_VLDL_0004).Bold = True
                .ListSubItems(COL_AGDF_NET_VLDL_PGRC).Bold = True
        
            Case Else
                
                .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(dblAConcordar)
                .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(dblConcordado)
                .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(dblSistemaOrigem)
                .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblArquivo)
                .SubItems(COL_BOAA_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblDiferenca)
                
                .ListSubItems(COL_BOAA_NET_VALR_ACON).Bold = True
                .ListSubItems(COL_BOAA_NET_VALR_CONC).Bold = True
                .ListSubItems(COL_BOAA_NET_TOTL_SIST).Bold = True
                .ListSubItems(COL_BOAA_NET_VALR_CAMR).Bold = True
                .ListSubItems(COL_BOAA_NET_DIFE_VALR).Bold = True
        
        End Select
    End With
    
Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularTotais", 0)

End Sub

'Mostra os campos de detalhes das operações
Private Sub flCarregarListaDetalheOperacoes()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strItemKey                              As String
Dim dblValorCamara                          As Double
Dim dblValorSistema                         As Double

On Error GoTo ErrorHandler

    If lvwNet.SelectedItem Is Nothing Then Exit Sub

    strItemKey = lvwNet.SelectedItem.Key
    lvwDetalhe.ListItems.Clear
    
    If strItemKey = strChaveTotais Then Exit Sub

    For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))

        With lvwDetalhe.ListItems.Add(, "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            .Text = objDomNode.selectSingleNode("DE_GRUP_LANC_FINC").Text
            .SubItems(COL_BOAA_DET_CODI_OPER) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            .SubItems(COL_BOAA_DET_VALR_SIST) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
            .SubItems(COL_BOAA_DET_STAT_OPER) = objDomNode.selectSingleNode("DE_SITU_PROC").Text

            If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                .SubItems(COL_BOAA_DET_DEBT_CRED) = "Débito"
            Else
                .SubItems(COL_BOAA_DET_DEBT_CRED) = "Crédito"
            End If
            
        End With

    Next

    Call fgClassificarListview(Me.lvwDetalhe, lngIndexClassifListDet, True)
    
Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaDetalheOperacoes", 0)

End Sub

'Carregar dados com NET de operações
Private Sub flCarregarListaNetArquivoCamara(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objRemessa                         As MSSOAPLib30.SoapClient30
#Else
    Dim objRemessa                          As A8MIU.clsRemessaFinanceiraCBLC
#End If

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String

Dim dblValorCamara                          As Double
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objRemessa = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraCBLC")
    strRetLeitura = objRemessa.ObterDetalheRemessaCBLC(pstrFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objRemessa = Nothing

    Call xmlArquivoCamara.loadXML(strRetLeitura)

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetLeitura.loadXML(strRetLeitura)

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheRemessa/*")

            strListItemKey = flMontarChaveItemListview(objDomNode)
            dblValorCamara = flValorOperacoes(strListItemKey, enumValoresCalculados.Camara)

            If Not fgExisteItemLvw(Me.lvwNet, strListItemKey) Then
            
                With lvwNet.ListItems.Add(, strListItemKey)
                    
                    .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                    If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                        .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                    End If
                    
                    Select Case PerfilAcesso
                        Case enumPerfilAcesso.AdmGeralPrevia
                        
                            .SubItems(COL_AGPV_NET_VALR_SIST) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface(dblValorCamara)
                            .SubItems(COL_AGPV_NET_VLDL_0001) = " "
                            .SubItems(COL_AGPV_NET_VLDL_1001) = " "
                        
                        Case enumPerfilAcesso.AdmGeral
                        
                            .SubItems(COL_AGDF_NET_VALR_SIST) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_AGDF_NET_VALR_DEFI) = fgVlrXml_To_Interface(dblValorCamara)
                            .SubItems(COL_AGDF_NET_VLDL_0004) = " "
                            .SubItems(COL_AGDF_NET_VLDL_PGRC) = " "
                        
                        Case Else
                            
                            .SubItems(COL_BOAA_NET_VEIC_LEGA) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                            .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                    
                    End Select
                End With
            Else

                With lvwNet.ListItems(strListItemKey)
                
                    Select Case PerfilAcesso
                        Case enumPerfilAcesso.AdmGeralPrevia
                            .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface(dblValorCamara)
                        Case enumPerfilAcesso.AdmGeral
                            .SubItems(COL_AGDF_NET_VALR_DEFI) = fgVlrXml_To_Interface(dblValorCamara)
                        Case Else
                            .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                    End Select
                End With
            End If
        Next
    End If

    Call fgClassificarListview(Me.lvwNet, lngIndexClassifListNet, True)
    
    Set xmlRetLeitura = Nothing

Exit Sub
ErrorHandler:
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetArquivoCamara", 0)

End Sub

'Carregar dados com NET de operações
Private Sub flCarregarListaNetArquivoCamaraPorArea(ByVal pstrFiltro As String)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
Dim dblValorCamara                          As Double

On Error GoTo ErrorHandler

    For Each objDomNode In xmlArquivoCamara.selectNodes("Repeat_DetalheRemessa/*")

        strListItemKey = flMontarChaveItemListviewPorArea(objDomNode)
        dblValorCamara = flValorOperacoesPorArea(strListItemKey, enumValoresCalculados.Camara)

        If Not fgExisteItemLvw(Me.lvwDetalhe, strListItemKey) Then
        
            With lvwDetalhe.ListItems.Add(, strListItemKey)
                
                .Text = objDomNode.selectSingleNode("DE_BKOF").Text
                .SubItems(COL_ADMG_DET_VALR_SIST) = fgVlrXml_To_Interface(0)
                .SubItems(COL_ADMG_DET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                
                If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Then
'                    .SubItems(COL_ADMG_DET_JUST_DIVG) = objDomNode.selectSingleNode("DE_JUST").Text
                End If
            
            End With
            
        Else

            With lvwDetalhe.ListItems(strListItemKey)
                .SubItems(COL_ADMG_DET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
            End With
        End If
    Next

    Call fgClassificarListview(Me.lvwDetalhe, lngIndexClassifListDet, True)
    
Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetArquivoCamaraPorArea", 0)

End Sub

'Carregar dados com valores de Mensagens
Private Sub flCarregarListaNetMensagens(ByVal pstrFiltro As String)

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

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim NU_CTRL_IF_LDL1001                      As String
Dim DH_REGT_MESG_SPB_LDL1001                As String
Dim xmlLDL0004Previa                        As String
Dim xmlDocLDL0004Previa                     As MSXML2.DOMDocument40
Dim xmlNodeLDL0004P                         As MSXML2.IXMLDOMNode
Dim strXPath                                As String

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(pstrFiltro, vntCodErro, vntMensagemErro)
    Set objMensagem = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetLeitura.loadXML(strRetLeitura)
        '-------------------------------------------------------
        Set objDomNode = xmlRetLeitura.selectSingleNode("Repeat_DetalheMensagem/*[./CO_MESG_SPB='LDL0004  ']")
        
        If Not objDomNode Is Nothing Then
          If Not objDomNode.selectSingleNode("CO_TEXT_XML") Is Nothing Then

              Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
              xmlLDL0004Previa = objMensagem.ObterXMLMensagem(objDomNode.selectSingleNode("CO_TEXT_XML").Text, vntCodErro, vntMensagemErro)
              Set objMensagem = Nothing
              
              If xmlLDL0004Previa <> "" Then
                  Set xmlDocLDL0004Previa = CreateObject("MSXML2.DOMDocument.4.0")
                  xmlDocLDL0004Previa.loadXML xmlLDL0004Previa
              End If
              
              If vntCodErro <> 0 Then
                  GoTo ErrorHandler
              End If
          End If
        End If
        '-------------------------------------------------------


        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*[./CO_MESG_SPB='LDL0001  ' or ./CO_MESG_SPB='LDL1001  ']")


            If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
              If objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.MensagemLiquidada Then
                  GoTo proximo
              End If
            End If

            strListItemKey = flMontarChaveItemListview(objDomNode)
              
            strXPath = "Repeat_DetalheMensagem/*[./CO_MESG_SPB='LDL1001  ' and CO_PARP_CAMR ='" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text & "']"
            
            If xmlRetLeitura.selectNodes(strXPath).length <> 0 Then
               NU_CTRL_IF_LDL1001 = xmlRetLeitura.selectNodes(strXPath).Item(0).selectSingleNode("NU_CTRL_IF").Text
               DH_REGT_MESG_SPB_LDL1001 = xmlRetLeitura.selectNodes(strXPath).Item(0).selectSingleNode("DH_REGT_MESG_SPB").Text
            Else
               NU_CTRL_IF_LDL1001 = ""
               DH_REGT_MESG_SPB_LDL1001 = ""
            End If
              
            strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                             "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & _
                             "|" & objDomNode.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text & _
                             "|" & objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text & _
                             "|" & objDomNode.selectSingleNode("NU_CTRL_CAMR").Text & _
                             "|" & NU_CTRL_IF_LDL1001 & _
                             "|" & DH_REGT_MESG_SPB_LDL1001



            If Not fgExisteItemLvw(Me.lvwNet, strListItemKey) Then
                
                With lvwNet.ListItems.Add(, strListItemKey)
                    
                    .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                    If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                        .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                    End If

                    If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito And _
                       objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                        
                        objDomNode.selectSingleNode("VA_FINC").Text = "-" & objDomNode.selectSingleNode("VA_FINC").Text
                    
                    End If
                    
                    Select Case PerfilAcesso
                        Case enumPerfilAcesso.AdmGeralPrevia
                        
                            .SubItems(COL_AGPV_NET_VALR_SIST) = fgVlrXml_To_Interface(0)
                        
                            If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                                .SubItems(COL_AGPV_NET_VLDL_0001) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                            Else
                                .SubItems(COL_AGPV_NET_VLDL_1001) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                            End If
                        
                            If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                              If IsNumeric(.SubItems(COL_AGPV_NET_VLDL_0001)) And IsNumeric(.SubItems(COL_AGPV_NET_VLDL_1001)) Then
                                  .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface( _
                                                                      (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) - _
                                                                      (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))))
                              End If
                            End If
                            
                        Case enumPerfilAcesso.AdmGeral
                        
                            .SubItems(COL_AGDF_NET_VALR_SIST) = fgVlrXml_To_Interface(0)
                        
                            If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                                .SubItems(COL_AGDF_NET_VALR_DEFI) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                            Else
                                .SubItems(COL_AGDF_NET_VLDL_0004) = "-" & fgVlrXml_To_Interface(fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_FINC").Text))
                            End If
                        
                            If IsNumeric(.SubItems(COL_AGDF_NET_VALR_DEFI)) And IsNumeric(.SubItems(COL_AGDF_NET_VLDL_0004)) Then
                                
                                
                                .SubItems(COL_AGDF_NET_VLDL_PGRC) = fgVlrXml_To_Interface( _
                                                                    (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VALR_DEFI)))) - _
                                                                    (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VLDL_0004)))))
                                
                                If .Index = 1 Then
                                    If fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VLDL_PGRC))) >= 0 Then
                                        With Me.tlbComandos
                                            .Buttons("pagamentocontingencia").Enabled = False
                                            .Buttons("regularizacao").Enabled = False
                                        End With
                                    Else
                                        With Me.tlbComandos
                                            .Buttons("pagamentocontingencia").Enabled = True
                                            .Buttons("regularizacao").Enabled = True
                                        End With
                                    End If
                                End If
                            End If
                            
                        Case Else
                            
                            .SubItems(COL_BOAA_NET_VEIC_LEGA) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                            .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(0)
                            txtValorMensagem.Text = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    
                    End Select
                    
                End With

            Else
                With lvwNet.ListItems(strListItemKey)

                    If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito And _
                       objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                       
                        objDomNode.selectSingleNode("VA_FINC").Text = "-" & objDomNode.selectSingleNode("VA_FINC").Text
                        
                    End If
                    
                    Select Case PerfilAcesso
                        Case enumPerfilAcesso.AdmGeralPrevia
                        
                            If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                                .SubItems(COL_AGPV_NET_VLDL_0001) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                            Else
                                .SubItems(COL_AGPV_NET_VLDL_1001) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text * -1)
                            End If
                            
                            If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                                If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                                  If IsNumeric(.SubItems(COL_AGPV_NET_VLDL_0001)) And IsNumeric(.SubItems(COL_AGPV_NET_VLDL_1001)) Then
                                      .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface( _
                                                                          (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) - _
                                                                          (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))))
                                  End If
                                Else
                                  If IsNumeric(.SubItems(COL_AGPV_NET_VLDL_0001)) And IsNumeric(.SubItems(COL_AGPV_NET_VLDL_1001)) Then
                                      .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface( _
                                                                          (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) - _
                                                                          (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))))
                                  End If
                                
                                End If
                            Else
                                If IsNumeric(.SubItems(COL_AGPV_NET_VLDL_0001)) And IsNumeric(.SubItems(COL_AGPV_NET_VLDL_1001)) Then
                                    .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface( _
                                                                        (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) - _
                                                                        (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))))
                                End If
                            End If
                            
                        Case enumPerfilAcesso.AdmGeral
                            
                            .SubItems(COL_AGDF_NET_VALR_DEFI) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                            
                            If Not xmlDocLDL0004Previa Is Nothing Then
                                  If Not xmlDocLDL0004Previa.selectSingleNode("//Repet_LDL0001_ResultLiqd/*[IdentdPartCamr='" & Trim(objDomNode.selectSingleNode("CO_PARP_CAMR").Text) & "']") Is Nothing Then
                                      Set xmlNodeLDL0004P = xmlDocLDL0004Previa.selectSingleNode("//Repet_LDL0001_ResultLiqd/*[IdentdPartCamr='" & Trim(objDomNode.selectSingleNode("CO_PARP_CAMR").Text) & "']")
                                      .SubItems(COL_AGDF_NET_VLDL_0004) = "-" & fgVlrXml_To_Interface(fgVlrXml_To_Decimal(xmlNodeLDL0004P.selectSingleNode("VlrResultLiqdNLiqdant").Text))
                                  Else
                                      .SubItems(COL_AGDF_NET_VLDL_0004) = fgVlrXml_To_Interface(0)
                                  End If
                            Else
                                .SubItems(COL_AGDF_NET_VLDL_0004) = fgVlrXml_To_Interface(0)
                            End If
                        
                            If IsNumeric(.SubItems(COL_AGDF_NET_VALR_DEFI)) And IsNumeric(.SubItems(COL_AGDF_NET_VLDL_0004)) Then
                                
                                
                                .SubItems(COL_AGDF_NET_VLDL_PGRC) = fgVlrXml_To_Interface( _
                                                                    (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VALR_DEFI)))) - _
                                                                    (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VLDL_0004)))))
                                
                                If .Index = 1 Then
                                    If fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGDF_NET_VLDL_PGRC))) >= 0 Then
                                        With Me.tlbComandos
                                            .Buttons("pagamentocontingencia").Enabled = False
                                            .Buttons("regularizacao").Enabled = False
                                        End With
                                    Else
                                        With Me.tlbComandos
                                            .Buttons("pagamentocontingencia").Enabled = True
                                            .Buttons("regularizacao").Enabled = True
                                        End With
                                    End If
                                End If
                            End If
                            
                        Case Else
                            
                            txtValorMensagem.Text = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    
                    End Select
                
                End With
            End If

            If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                lvwNet.ListItems(strListItemKey).Tag = strListItemTag
            End If
proximo:
        Next

  
    End If

    Call fgClassificarListview(Me.lvwNet, lngIndexClassifListNet, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:

    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetMensagens", 0, 0, "Linha: " & Erl)

End Sub

'Carregar dados com NET de operações
Private Sub flCarregarListaNetOperacoes(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String

Dim dblValorAConcordar                      As Double
Dim dblValorConcordado                      As Double
Dim dblValorSistOrigem                      As Double
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(pstrFiltro, _
                                                     vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing

    Call xmlOperacoes.loadXML(strRetLeitura)

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetLeitura.loadXML(strRetLeitura)

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")

            strListItemKey = flMontarChaveItemListview(objDomNode)

            If Not fgExisteItemLvw(Me.lvwNet, strListItemKey) Then
            
                With lvwNet.ListItems.Add(, strListItemKey)

                    .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                    If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                        .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                    End If
                    
                    Select Case PerfilAcesso
                        Case enumPerfilAcesso.AdmGeralPrevia
                        
                            dblValorSistOrigem = flValorOperacoes(strListItemKey, enumValoresCalculados.SistemaOrigem)
                            
                            .SubItems(COL_AGPV_NET_VALR_SIST) = fgVlrXml_To_Interface(dblValorSistOrigem)
                            .SubItems(COL_AGPV_NET_VLDL_0001) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_AGPV_NET_VLDL_1001) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_AGPV_NET_VALR_PREV) = fgVlrXml_To_Interface(0)
                            
                        Case enumPerfilAcesso.AdmGeral
                        
                            dblValorSistOrigem = flValorOperacoes(strListItemKey, enumValoresCalculados.SistemaOrigem)
                            
                            .SubItems(COL_AGDF_NET_VALR_SIST) = fgVlrXml_To_Interface(dblValorSistOrigem)
                            .SubItems(COL_AGDF_NET_VALR_DEFI) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_AGDF_NET_VLDL_0004) = fgVlrXml_To_Interface(0)
                            .SubItems(COL_AGDF_NET_VLDL_PGRC) = fgVlrXml_To_Interface(0)
                            
                        Case Else
                            
                            dblValorAConcordar = flValorOperacoes(strListItemKey, enumValoresCalculados.AConcordar)
                            dblValorConcordado = flValorOperacoes(strListItemKey, enumValoresCalculados.Concordado)
                            
                            .SubItems(COL_BOAA_NET_VEIC_LEGA) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                            .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(dblValorAConcordar)
                            .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(dblValorConcordado)
                            .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(dblValorAConcordar + dblValorConcordado)
                            .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(0)
                    
                    End Select
                End With
            End If
        Next
    End If

    Call fgClassificarListview(Me.lvwNet, lngIndexClassifListNet, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes", 0)

End Sub

'Carregar dados com NET de operações
Private Sub flCarregarListaNetOperacoesPorArea(ByVal pstrFiltro As String)

Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodeJustif                        As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
Dim dblValorSistOrigem                      As Double

On Error GoTo ErrorHandler
    
    For Each objDomNode In xmlOperacoes.selectNodes("Repeat_DetalheOperacao/*")

        If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Then
    
            Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
            Call fgAppendNode(xmlLeitura, "", "Grupo_Parametros", "")
            Call fgAppendNode(xmlLeitura, "Grupo_Parametros", "CO_EMPR", fgObterCodigoCombo(Me.cboEmpresa.Text))
            Call fgAppendNode(xmlLeitura, "Grupo_Parametros", "TP_BKOF", 0)
            Call fgAppendNode(xmlLeitura, "Grupo_Parametros", "DT_LIQU", fgDate_To_DtXML(fgDataHoraServidor(DataAux)))
    
            Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A8LQS.clsJustificativaArea", xmlLeitura))
    
        End If
        
        strListItemKey = flMontarChaveItemListviewPorArea(objDomNode)

        If Not fgExisteItemLvw(Me.lvwDetalhe, strListItemKey) Then
        
            With lvwDetalhe.ListItems.Add(, strListItemKey)

                dblValorSistOrigem = flValorOperacoesPorArea(strListItemKey, enumValoresCalculados.SistemaOrigem)
                
                .Text = objDomNode.selectSingleNode("DE_BKOF").Text
                .SubItems(COL_ADMG_DET_VALR_SIST) = fgVlrXml_To_Interface(dblValorSistOrigem)
                .SubItems(COL_ADMG_DET_VALR_CAMR) = fgVlrXml_To_Interface(0)
                
                If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Then
                    For Each objDomNodeJustif In xmlLeitura.selectNodes("Repeat_JustificativaArea/Grupo_JustificativaArea[TP_BKOF='" & objDomNode.selectSingleNode("TP_BKOF").Text & "']")
                        .SubItems(COL_ADMG_DET_JUST_DIVG) = objDomNodeJustif.selectSingleNode("DE_JUST").Text
                    Next
                End If

            End With

        End If

    Next

    Call fgClassificarListview(Me.lvwDetalhe, lngIndexClassifListDet, True)
    
    Set xmlLeitura = Nothing

Exit Sub
ErrorHandler:
    Set xmlLeitura = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetOperacoesPorArea", 0)

End Sub

'Altera a exibição dos botões de acordo com o perfil do usuário
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)

On Error GoTo ErrorHandler

    With tlbComandos
        Select Case PerfilAcesso
            Case enumPerfilAcesso.AdmGeralPrevia
            
                .Buttons("concordancia").Visible = False
                .Buttons("discordancia").Visible = False
                .Buttons("retorno").Visible = True
                .Buttons("liberacao").Visible = True
                .Buttons("pagamentocontingencia").Visible = False
                .Buttons("regularizacao").Visible = False
                
                .Buttons("retorno").Enabled = True
                .Buttons("liberacao").Enabled = True
            
            Case enumPerfilAcesso.AdmGeral
            
                .Buttons("concordancia").Visible = True
                .Buttons("discordancia").Visible = True
                .Buttons("retorno").Visible = True
                .Buttons("liberacao").Visible = True
                .Buttons("pagamentocontingencia").Visible = True
                .Buttons("regularizacao").Visible = True
                
                .Buttons("concordancia").Enabled = True
                .Buttons("discordancia").Enabled = True
                .Buttons("retorno").Enabled = True
                .Buttons("liberacao").Enabled = True
                .Buttons("pagamentocontingencia").Enabled = True
                .Buttons("regularizacao").Enabled = True
                
            Case Else
                
                .Buttons("concordancia").Visible = True
                .Buttons("discordancia").Visible = False
                .Buttons("retorno").Visible = False
                .Buttons("liberacao").Visible = False
                .Buttons("pagamentocontingencia").Visible = False
                .Buttons("regularizacao").Visible = False
        
                .Buttons("concordancia").Enabled = True
        
        End Select
        
        .Refresh
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorPerfil", 0

End Sub

'Inicializa controles de tela e variáveis
Private Sub flInicializarFormulario()

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
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializarFormulario")
    End If

    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")

    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    fgRaiseError App.EXEName, TypeName(Me), "flInicializarFormulario", 0

End Sub

'Formata as colunas da lista de operações
Private Sub flInicializarlvwDetalhe()

On Error GoTo ErrorHandler

    With Me.lvwDetalhe.ColumnHeaders
        
        .Clear
        
        Select Case PerfilAcesso
            Case enumPerfilAcesso.AdmGeralPrevia
            
                lvwDetalhe.CheckBoxes = True
                
                .Add , , "Área", 2000
                .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
                .Add , , "Valor Câmara (Arquivo)", 2000, lvwColumnRight
                .Add , , "Diferença", 2000, lvwColumnRight
                .Add , , "Justificativa", 4000
                
            Case enumPerfilAcesso.AdmGeral
            
                lvwDetalhe.CheckBoxes = True
                
                .Add , , "Área", 2000
                .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
                .Add , , "Valor Câmara (Arquivo)", 2000, lvwColumnRight
                .Add , , "Diferença", 2000, lvwColumnRight
                
            Case Else
                
                lvwDetalhe.CheckBoxes = False
                
                .Add , , "Código Lançamento", 2000
                .Add , , "Código Operação", 2000
                .Add , , "D/C", 2000
                .Add , , "Valor Sistema Origem", 1800, lvwColumnRight
                .Add , , "Situação", 3600
        
        End Select
        
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarlvwDetalhe", 0

End Sub

'Formata as colunas da lista de mensagens
Private Sub flInicializarlvwNet()

On Error GoTo ErrorHandler

    lvwNet.CheckBoxes = (PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia)
    
    With Me.lvwNet.ColumnHeaders
        .Clear

        Select Case PerfilAcesso
            Case enumPerfilAcesso.AdmGeralPrevia
            
                lvwNet.CheckBoxes = False
                lvwNet.HideSelection = True
                
                .Add , , "Agente Compensação", 2500
                .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
                .Add , , "Valor Prévia (LDL0001 - LDL1001)", 3500, lvwColumnRight
                .Add , , "Diferença", 2000, lvwColumnRight
                .Add , , "Valor LDL0001", 2400, lvwColumnRight
                .Add , , "Valor LDL1001", 2400, lvwColumnRight
                
            Case enumPerfilAcesso.AdmGeral
            
                lvwNet.CheckBoxes = False
                lvwNet.HideSelection = True
                
                .Add , , "Agente Compensação", 2500
                .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
                .Add , , "Valor LDL0001 Definitiva", 2500, lvwColumnRight
                .Add , , "Diferença", 2000, lvwColumnRight
                .Add , , "Valor LDL0004 Prévia", 2400, lvwColumnRight
                .Add , , "Valor a Pagar ou a Receber", 3400, lvwColumnRight
                
            Case enumPerfilAcesso.AdmArea
            
                lvwNet.CheckBoxes = False
                lvwNet.HideSelection = False
                
                .Add , , "Agente Compensação", 2000
                .Add , , "Veículo Legal", 4000
                .Add , , "Valor a Concordar", 1800, lvwColumnRight
                .Add , , "Valor Concordado", 1800, lvwColumnRight
                .Add , , "Total Sistema Origem", 1800, lvwColumnRight
                .Add , , "Valor Câmara Prévia", 1800, lvwColumnRight
                .Add , , "Diferença", 1600, lvwColumnRight
        
            Case enumPerfilAcesso.BackOffice
                
                lvwNet.CheckBoxes = True
                lvwNet.HideSelection = False
                
                .Add , , "Agente Compensação", 2000
                .Add , , "Veículo Legal", 4000
                .Add , , "Valor a Concordar", 1800, lvwColumnRight
                .Add , , "Valor Concordado", 1800, lvwColumnRight
                .Add , , "Total Sistema Origem", 1800, lvwColumnRight
                .Add , , "Valor Câmara Prévia", 1800, lvwColumnRight
                .Add , , "Diferença", 1600, lvwColumnRight
        
        End Select
        
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarlvwNet", 0

End Sub

'Apaga o conteúdo das listas de mensagens e operações
Private Sub flLimparListas()
    Me.lvwNet.ListItems.Clear
    Me.lvwDetalhe.ListItems.Clear
End Sub

'Monta o conteúdo que será utilizado com a propriedade 'Key' dos itens do ListView
Private Function flMontarChaveItemListview(ByVal objDomNode As MSXML2.IXMLDOMNode)

Dim strListItemKey                          As String

On Error GoTo ErrorHandler

    Select Case PerfilAcesso
        Case enumPerfilAcesso.AdmGeralPrevia, enumPerfilAcesso.AdmGeral
        
            strListItemKey = "|" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text
            
        Case Else
            
            strListItemKey = "|" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text & _
                             "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                             "|" & objDomNode.selectSingleNode("TP_BKOF").Text
    
    End Select
    
    flMontarChaveItemListview = strListItemKey

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Monta o conteúdo que será utilizado com a propriedade 'Key' dos itens do ListView
Private Function flMontarChaveItemListviewPorArea(ByVal objDomNode As MSXML2.IXMLDOMNode)

Dim strListItemKey                          As String

On Error GoTo ErrorHandler

    strListItemKey = "|" & objDomNode.selectSingleNode("TP_BKOF").Text
            
    flMontarChaveItemListviewPorArea = strListItemKey

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListviewPorArea", 0

End Function

'Monta uma expressão XPath para seleção do conteúdo de um documento XML
Private Function flMontarCondicaoNavegacaoXMLOperacoes(ByVal strItemKey As String)

Dim strCondicao                             As String

On Error GoTo ErrorHandler

    Select Case PerfilAcesso
        Case enumPerfilAcesso.AdmGeralPrevia, enumPerfilAcesso.AdmGeral
        
            strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_ADMG_NET_AGEN_COMP) & "']"
            
        Case Else
            
            strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_BOAA_NET_AGEN_COMP) & "' " & _
                                                                  " and CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_BOAA_NET_VEIC_LEGA) & "' " & _
                                                                  " and TP_BKOF='" & Split(strItemKey, "|")(KEY_BOAA_NET_TIPO_BKOF) & "']"
    
    End Select
    
    flMontarCondicaoNavegacaoXMLOperacoes = strCondicao

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoes", 0

End Function

'Monta uma expressão XPath para seleção do conteúdo de um documento XML
Private Function flMontarCondicaoNavegacaoXMLOperacoesPorArea(ByVal strItemKey As String)

Dim strCondicao                             As String

On Error GoTo ErrorHandler

    strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[TP_BKOF='" & Split(strItemKey, "|")(KEY_ADMG_DET_TIPO_BKOF) & "']"
    flMontarCondicaoNavegacaoXMLOperacoesPorArea = strCondicao

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoesPorArea", 0

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String, ByVal intValor As enumValoresCalculados) As String

Dim strDebito                               As String
Dim strCredito                              As String

On Error GoTo ErrorHandler

    Select Case PerfilAcesso
        Case enumPerfilAcesso.AdmGeralPrevia, enumPerfilAcesso.AdmGeral
        
            strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                             " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_ADMG_NET_AGEN_COMP) & "' "
        
            strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                              " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_ADMG_NET_AGEN_COMP) & "' "
            
        Case Else
            
            strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                             " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_BOAA_NET_AGEN_COMP) & "' " & _
                                             " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_BOAA_NET_VEIC_LEGA) & "' " & _
                                             " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_BOAA_NET_TIPO_BKOF) & "' "
        
            strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                              " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_BOAA_NET_AGEN_COMP) & "' " & _
                                              " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_BOAA_NET_VEIC_LEGA) & "' " & _
                                              " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_BOAA_NET_TIPO_BKOF) & "' "
            
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                Select Case intValor
                    Case enumValoresCalculados.AConcordar
                        strDebito = strDebito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.EmSer & "' "
                        strCredito = strCredito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.EmSer & "' "
                    Case enumValoresCalculados.Concordado
                        strDebito = strDebito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackoffice & "' "
                        strCredito = strCredito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackoffice & "' "
                End Select
                
            Else
                Select Case intValor
                    Case enumValoresCalculados.AConcordar
                        strDebito = strDebito & " and (../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackoffice & "' or ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackofficeAutomatico & "')"
                        strCredito = strCredito & " and (../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackoffice & "' or ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackofficeAutomatico & "')"
                    Case enumValoresCalculados.Concordado
                        strDebito = strDebito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaAdmArea & "' "
                        strCredito = strCredito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaAdmArea & "' "
                End Select
                
            End If
    
    End Select
    
    strDebito = strDebito & "]) "
    strCredito = strCredito & "]) "

    flMontarExpressaoCalculoNetOperacoes = strCredito & " - " & strDebito

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoCalculoNetOperacoes", 0

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Private Function flMontarExpressaoCalculoNetOperacoesPorArea(ByVal strItemKey As String, ByVal intValor As enumValoresCalculados) As String

Dim strDebito                               As String
Dim strCredito                              As String

On Error GoTo ErrorHandler

    strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_ADMG_DET_TIPO_BKOF) & "' "

    strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                      " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_ADMG_DET_TIPO_BKOF) & "' "
    
    strDebito = strDebito & "])  "
    strCredito = strCredito & "]) "

    flMontarExpressaoCalculoNetOperacoesPorArea = strCredito & " - " & strDebito

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoCalculoNetOperacoesPorArea", 0

End Function

'Monta o XML com os dados de filtro para seleção de operações
Private Function flMontarXMLFiltroPesquisa() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.BMD)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ControleRepeticao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_ControleRepeticao", "ControleRepeticao", "> 1")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF)
    Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.NETEntradaManualMultilateralBMD)

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Multilateral)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    Select Case PerfilAcesso
        Case enumPerfilAcesso.AdmGeralPrevia
        
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLDL", "")
            Call fgAppendNode(xmlFiltros, "Grupo_TipoLDL", "TipoLDL", "P")
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
            Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL0001")
            Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL1001")
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_GrupoLancamentoFinanceiro", "")
            Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "6")
            Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "7")
            Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "38")
            Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "999")
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoAcao", "")
            'Call fgAppendNode(xmlFiltros, "Grupo_TipoAcao", "TipoAcao", enumTipoAcao.PreviaLiquidada)
            'Call fgAppendNode(xmlFiltros, "Grupo_TipoAcao", "TipoAcao", "0")
        
        Case enumPerfilAcesso.AdmGeral
        
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.MensagemLiquidada)
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_PerfilAcesso", "")
            Call fgAppendNode(xmlFiltros, "Grupo_PerfilAcesso", "PerfilAcesso", enumPerfilAcesso.AdmGeral)
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoAcao", "")
            'Call fgAppendNode(xmlFiltros, "Grupo_TipoAcao", "TipoAcao", enumTipoAcao.PreviaLiquidada)
            'Call fgAppendNode(xmlFiltros, "Grupo_TipoAcao", "TipoAcao", "0")
        
        Case enumPerfilAcesso.AdmArea
        
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeAutomatico)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.MensagemLiquidada)
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLDL", "")
            
            If optTipoMensagem(enumPreviaDefinitiva.Previa).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_TipoLDL", "TipoLDL", "P")
            
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_GrupoLancamentoFinanceiro", "")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "6")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "7")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "38")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "999")
            
            Else
                Call fgAppendNode(xmlFiltros, "Grupo_TipoLDL", "TipoLDL", "D")
            
            End If
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
            Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL0001")
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoAcao", "")
            Call fgAppendNode(xmlFiltros, "Grupo_TipoAcao", "TipoAcao", "0")
        
        Case enumPerfilAcesso.BackOffice
            
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
    
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLDL", "")
            
            If optTipoMensagem(enumPreviaDefinitiva.Previa).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_TipoLDL", "TipoLDL", "P")
            
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_GrupoLancamentoFinanceiro", "")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "6")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "7")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "38")
                Call fgAppendNode(xmlFiltros, "Grupo_GrupoLancamentoFinanceiro", "GrupoLancamentoFinanceiro", "999")
            
            Else
                Call fgAppendNode(xmlFiltros, "Grupo_TipoLDL", "TipoLDL", "D")
            
            End If
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
            Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL0001")
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoAcao", "")
            Call fgAppendNode(xmlFiltros, "Grupo_TipoAcao", "TipoAcao", "0")
        
    End Select
    
    flMontarXMLFiltroPesquisa = xmlFiltros.xml

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Monta XML com as chaves das operações que serão processadas
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer

Dim intStatusOper                           As Integer
Dim intStatusOper1                          As Integer
Dim intColunaValor                          As Integer
Dim strIndPreviaDefinitiva                  As String

On Error GoTo ErrorHandler

    intStatusOper = 0
    intStatusOper1 = 0
    
    Select Case PerfilAcesso
        Case enumPerfilAcesso.BackOffice
            intStatusOper = enumStatusOperacao.EmSer
            intColunaValor = COL_BOAA_NET_VALR_ACON
            
            If optTipoMensagem(enumPreviaDefinitiva.Previa).value Then
                strIndPreviaDefinitiva = "P"
            Else
                strIndPreviaDefinitiva = "D"
            End If

        Case enumPerfilAcesso.AdmArea
            intStatusOper = enumStatusOperacao.ConcordanciaBackoffice
            intStatusOper1 = enumStatusOperacao.ConcordanciaBackofficeAutomatico
            intColunaValor = COL_BOAA_NET_VALR_ACON

            If optTipoMensagem(enumPreviaDefinitiva.Previa).value Then
                strIndPreviaDefinitiva = "P"
            Else
                strIndPreviaDefinitiva = "D"
            End If

        Case enumPerfilAcesso.AdmGeralPrevia
            intStatusOper = enumStatusOperacao.ConcordanciaAdmArea
            intColunaValor = COL_AGPV_NET_VALR_PREV
            strIndPreviaDefinitiva = "P"

        Case enumPerfilAcesso.AdmGeral
            intStatusOper = enumStatusOperacao.ConcordanciaAdmArea
            intColunaValor = COL_AGDF_NET_VLDL_PGRC
            strIndPreviaDefinitiva = "D"

    
    End Select

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    For Each objListItem In Me.lvwNet.ListItems
        With objListItem

            If .Checked Then

                If (PerfilAcesso = enumPerfilAcesso.AdmGeral Or _
                   .SubItems(intColunaValor) <> fgVlrXml_To_Interface(0)) And _
                   .Key <> strChaveTotais Then

                    'If .SubItems(intColunaValor) < 0 Then
    
                        Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")
                        Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
                        
    
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "IN_ULTI_MESG", "N")
                                            
    
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                           "CO_EMPR", _
                                                           fgObterCodigoCombo(cboEmpresa.Text))
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                           "CO_PARP_CAMR", _
                                                           Split(.Key, "|")(KEY_ADMG_NET_AGEN_COMP))
                        
                        If (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) < 0 And _
                           (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))) < 0 Then
                        
                            If (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) > _
                               (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))) Then
                        
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                                   "VA_OPER_ATIV", _
                                                                   "0")
                            
                            Else
                            
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                                   "VA_OPER_ATIV", _
                                                                   fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(intColunaValor))))
                            
                            End If
                        
                        Else
                        
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "VA_OPER_ATIV", _
                                                               fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(intColunaValor))))
                        
                        End If
                        
                        
                        
                        
                        If Trim$(txtJustificativa.Text) <> vbNullString Then
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "TP_BKOF", _
                                                               Split(.Key, "|")(KEY_BOAA_NET_TIPO_BKOF))
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "DE_JUST", _
                                                               Trim$(txtJustificativa.Text))
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "DT_LIQU", _
                                                               fgDate_To_DtXML(fgDataHoraServidor(DataAux)))
                        End If
    
                        intIgnoraGradeHorario = 1
    
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                           "IgnoraGradeHorario", _
                                                           intIgnoraGradeHorario)
    
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                           "Repeat_Operacao", _
                                                           "")
    
                        If .Tag <> vbNullString Then
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "NU_CTRL_IF", _
                                                               Split(.Tag, "|")(TAG_MSG_NU_CTRL_IF))
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "DH_REGT_MESG_SPB", _
                                                               Split(.Tag, "|")(TAG_MSG_DH_REGT_MESG_SPB))
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "NU_SEQU_CNTR_REPE", _
                                                               Split(.Tag, "|")(TAG_MSG_NU_SEQU_CNTR_REPE))
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "DH_ULTI_ATLZ", _
                                                               Split(.Tag, "|")(TAG_MSG_DH_ULTI_ATLZ))
                            
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "NU_CTRL_CAMR", _
                                                               Split(.Tag, "|")(TAG_MSG_NU_CTRL_CAMR))
                                                                                        
                            If Split(.Tag, "|")(TAG_MSG_IN_DEBT_CRED) = "Débito" Then
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "IN_DEBT_CRED", "D")
                            Else
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "IN_DEBT_CRED", "C")
                            End If
                            
                            If intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia Then
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "TP_CONF_DIVG", "C")
                            ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Then
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "TP_CONF_DIVG", "D")
                            End If
                            
                            If (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) < 0 And _
                               (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))) < 0 Then
                            
                                If (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_0001)))) > _
                                   (fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_AGPV_NET_VLDL_1001)))) Then
                            
                                    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                                       "VA_FINC", _
                                                                       "0")
                                
                                Else
                                
                                    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                                       "VA_FINC", _
                                                                       fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(intColunaValor))))
                                
                                End If
                            
                            Else
                            
                                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                                   "VA_FINC", _
                                                                   fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(intColunaValor))))
                            
                            End If
                            
                            
                            
                            
                            
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "TP_INFO_LDL", _
                                                               strIndPreviaDefinitiva)
                            
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "VA_OPER_MESG_BATIDOS", _
                                                               IIf(.ListSubItems(COL_AGPV_NET_DIFE_VALR).ForeColor = vbRed, "N", "S"))
                            
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "NU_CTRL_IF_LDL1001", _
                                                               Split(.Tag, "|")(TAG_MSG_NU_CTRL_IF_LDL1001))
                                                               
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                               "DH_REGT_MESG_SPB_LDL1001", _
                                                               Split(.Tag, "|")(TAG_MSG_DH_REGT_MESG_SPB_LDL1001))
                                                               
                        Else
                            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "IN_DEBT_CRED", "D")
                        End If
                        
                        For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItem.Key))
                            If intStatusOper = Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) Or intStatusOper1 = Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) Then
                                Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
                                Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
                                Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                                   "NU_SEQU_OPER_ATIV", _
                                                                   objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                                Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                                   "DH_ULTI_ATLZ", _
                                                                   objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
                                
                                Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                                   "CO_PARP_CAMR", _
                                                                   objDomNode.selectSingleNode("CO_PARP_CAMR").Text)
                                
                                
                                
                                Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)
    
                                Set xmlItemOperacao = Nothing
                            End If
                        Next
    
                        Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
                        Set xmlItemEnvioMsg = Nothing
                    
                    'End If
                    
                End If

            End If

        End With
    Next

    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLProcessamento = vbNullString
    Else
        flMontarXMLProcessamento = xmlProcessamento.xml
    End If

    Set xmlProcessamento = Nothing

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLProcessamento", 0

End Function

'Monta XML com as chaves das operações que serão processadas
Private Function flMontarXMLRejeicao() As String

Dim objListItem                             As ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer

On Error GoTo ErrorHandler

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    For Each objListItem In Me.lvwDetalhe.ListItems
        With objListItem

            If .Checked Then

                Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")
                Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_EMPR", _
                                                   fgObterCodigoCombo(cboEmpresa.Text))
                
                intIgnoraGradeHorario = 1

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "IgnoraGradeHorario", _
                                                   intIgnoraGradeHorario)

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "Repeat_Operacao", _
                                                   "")
                 Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "IN_DEBT_CRED", "D")
                For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoesPorArea(objListItem.Key))
                    Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")

                    Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "NU_SEQU_OPER_ATIV", _
                                                       objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "DH_ULTI_ATLZ", _
                                                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)

                    Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)

                    Set xmlItemOperacao = Nothing
                Next

                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
                Set xmlItemEnvioMsg = Nothing
            
            End If

        End With
    Next

    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLRejeicao = vbNullString
    Else
        flMontarXMLRejeicao = xmlProcessamento.xml
    End If

    Set xmlProcessamento = Nothing

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLRejeicao", 0

End Function

'Exibe o resultado da última operação executada
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String, _
                               ByVal pintAcaoConciliacao As Integer)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRejeitar Then
            .strDescricaoOperacao = " rejeitados pelo admnistrador "
        Else
            .strDescricaoOperacao = " liquidados "
        End If
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'Monta a tela com os dados do filtro selecionado
Private Sub flPesquisar()

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

    strDocFiltros = flMontarXMLFiltroPesquisa
    
    Call flCarregarListaNetOperacoes(strDocFiltros)
    Call flCarregarListaNetArquivoCamara(strDocFiltros)

    If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Or _
       PerfilAcesso = enumPerfilAcesso.AdmGeral Or _
       txtValorMensagem.Visible Then
       
        Call flCarregarListaNetMensagens(strDocFiltros)
        
    End If

    If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Or _
       PerfilAcesso = enumPerfilAcesso.AdmGeral Then
       
        Call flCarregarListaNetOperacoesPorArea(strDocFiltros)
        Call flCarregarListaNetArquivoCamaraPorArea(strDocFiltros)
        Call flCalcularDiferencasListViewDetalhe
        
    End If

    Call flCalcularDiferencasListViewNet
    
    If lvwNet.ListItems.Count > 0 And _
      (PerfilAcesso = enumPerfilAcesso.AdmArea Or _
       PerfilAcesso = enumPerfilAcesso.BackOffice) Then
       
        lvwNet.ListItems(1).Selected = True
        Call lvwNet_ItemClick(lvwNet.ListItems(1))
        
    End If

    Call flCalcularTotais
    
    txtJustificativa.Text = vbNullString
    
    fgCursor

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPesquisar", 0)

End Sub

'Enviar itens de mensagem e operações para liquidação
Private Function flProcessar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLRetorno                           As String
Dim strXMLProc                              As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    If intAcaoProcessamento = AdmGeralRejeitar Then
        strXMLProc = flMontarXMLRejeicao
    Else
        strXMLProc = flMontarXMLProcessamento
    End If

    If strXMLProc <> vbNullString Then
        fgCursor True
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.ProcessarLoteLiquidacaoMultilateralBMF(intAcaoProcessamento, _
                                                                                   strXMLProc, _
                                                                                   vntCodErro, _
                                                                                   vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objOperacaoMensagem = Nothing
        fgCursor
    End If

    If strXMLRetorno <> vbNullString Then
        Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetornoErro.loadXML(strXMLRetorno)
    Else
        Set xmlRetornoErro = Nothing
    End If

    flProcessar = strXMLRetorno

Exit Function
ErrorHandler:
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flProcessar", Me.Caption

End Function

'Valida a seleção dos itens na tela, para posterior processamento
Private Function flValidarItensProcessamento(ByVal intAcao As enumAcaoConciliacao) As String

Dim intLinhas                               As Integer
Dim objListItem                             As ListItem

    If intAcao = enumAcaoConciliacao.AdmGeralRejeitar Then
        If fgItemsCheckedListView(Me.lvwDetalhe) = 0 Then
            flValidarItensProcessamento = "Selecione pelo menos um item da lista de operações por Área, antes de prosseguir com a rejeição de Nets."
            Exit Function
        End If
    Else
        If fgItemsCheckedListView(Me.lvwNet) = 0 Then
            flValidarItensProcessamento = "Selecione pelo menos um item da lista de Nets de Operações, antes de prosseguir com a esta solicitação."
            Exit Function
        End If
    End If
    
    Select Case PerfilAcesso
        Case enumPerfilAcesso.BackOffice
        
            For Each objListItem In lvwNet.ListItems
                
                If objListItem.Checked And _
                   objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> objListItem.SubItems(COL_BOAA_NET_VALR_CAMR) And _
                   objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> fgVlrXml_To_Interface(0) Then
                    
                    flValidarItensProcessamento = "Valor A concordar é diferente do valor enviado pela câmara. Deseja prosseguir com a operação ?"
                    Exit Function
                
                End If
            
            Next
            
        Case enumPerfilAcesso.AdmArea

            For Each objListItem In lvwNet.ListItems
                If objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> objListItem.SubItems(COL_BOAA_NET_VALR_CAMR) And _
                   objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> fgVlrXml_To_Interface(0) Then
                    
                    flValidarItensProcessamento = "Valor A concordar é diferente do valor enviado pela câmara. Deseja prosseguir com a operação ?"
                    Exit Function
                
                End If
            Next

            intLinhas = 0
            For Each objListItem In lvwNet.ListItems
                If objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> fgVlrXml_To_Interface(0) Then
                    intLinhas = intLinhas + 1
                    Exit For
                End If
            Next

            If intLinhas = 0 Then
                flValidarItensProcessamento = "Todos os valores pendentes de concordância, para o Adm. de Área, já foram processados."
                Exit Function
            End If

        Case enumPerfilAcesso.AdmGeralPrevia

            If intAcao <> enumAcaoConciliacao.AdmGeralRejeitar Then

                For Each objListItem In lvwNet.ListItems
                    If objListItem.Checked And _
                       objListItem.Tag = vbNullString Then
                        
                        
                        flValidarItensProcessamento = "Valor Prévia não encontrado em um ou mais itens selecionados. Solicitação não permitida."
                        Exit Function
                    
                    End If
                Next
                
                For Each objListItem In lvwNet.ListItems
                    If objListItem.Checked And _
                       objListItem.Tag <> strChaveTotais Then
                        
                        If fgVlrXml_To_Decimal(objListItem.ListSubItems(COL_AGPV_NET_VALR_SIST)) = 0 Then
                            flValidarItensProcessamento = "Valor do Sistema Origem não informado. Solicitação não permitida."
                            Exit Function
                        End If
                        
                    
                    End If
                Next
    
                For Each objListItem In lvwNet.ListItems
                    If objListItem.Checked And _
                       objListItem.ListSubItems(COL_AGPV_NET_DIFE_VALR).ForeColor = vbRed Then
                        
                        flValidarItensProcessamento = "Valor do Sistema Origem é diferente do Valor Prévia (LDL0001 - LDL1001). Deseja prosseguir com a operação ?"
                        Exit Function
                    
                    End If
                Next
    
            End If

        Case enumPerfilAcesso.AdmGeral

            If intAcao <> enumAcaoConciliacao.AdmGeralRejeitar Then

                For Each objListItem In lvwNet.ListItems
                    If objListItem.Key <> strChaveTotais Then
                        If objListItem.Tag = vbNullString Then
                            flValidarItensProcessamento = "Valor LDL0001 Definitiva não encontrado em um ou mais itens. Solicitação não permitida."
                            Exit Function
                        
                        Else
                            If intAcao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
                               intAcao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Then
                                
                                If Split(objListItem.Tag, "|")(TAG_MSG_TP_ACAO_MESG_SPB_EXEC) = enumTipoAcao.Liberacao Then
                                    flValidarItensProcessamento = "Confirmação já foi enviada para esta mensagem. Solicitação não permitida."
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
    
                If intAcao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Then
                
                    For Each objListItem In lvwNet.ListItems
                        If objListItem.Key <> strChaveTotais Then
                            If objListItem.ListSubItems(COL_AGDF_NET_DIFE_VALR).ForeColor <> vbRed Then
                                
                                flValidarItensProcessamento = "Valor do Sistema Origem é igual ao Valor LDL0001 Definitiva. Solicitação não permitida."
                                Exit Function
                            
                            End If
                        End If
                    Next
                
                ElseIf intAcao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Then
                
                ElseIf intAcao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
                       intAcao = enumAcaoConciliacao.AdmGeralPagamento Then
                    
                    For Each objListItem In lvwNet.ListItems
                        If objListItem.Checked And _
                           objListItem.Tag <> strChaveTotais Then
                            
                            If fgVlrXml_To_Decimal(objListItem.ListSubItems(COL_AGPV_NET_VALR_SIST)) = 0 Then
                                flValidarItensProcessamento = "Valor do Sistema Origem não informado para um ou mais  membros de compensação. Solicitação não permitida."
                                Exit Function
                            End If
                            
                        
                        End If
                    Next
                
                ElseIf intAcao = enumAcaoConciliacao.AdmGeralRegularizar Or _
                       intAcao = enumAcaoConciliacao.AdmAreaPagamento Then
                
                    For Each objListItem In lvwNet.ListItems
                        If objListItem.Key <> strChaveTotais Then
                            If objListItem.ListSubItems(COL_AGDF_NET_DIFE_VALR).ForeColor = vbRed Then
                                
                                flValidarItensProcessamento = "Valor do Sistema Origem é diferente do Valor LDL0001 Definitiva. Solicitação não permitida."
                                Exit Function
                            
                            End If
                        End If
                    Next
                
                End If
        
            End If

    End Select

End Function

'Calcula o valor da operações
Private Function flValorOperacoes(ByVal strItemKey As String, ByVal intValor As enumValoresCalculados) As Variant

Dim strExpression                           As String
Dim vntValor                                As Variant
Dim xmlAux                                  As MSXML2.DOMDocument40

    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAux = IIf(intValor = Camara, xmlArquivoCamara, xmlOperacoes)

    vntValor = 0
    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey, intValor)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlAux, strExpression))

    flValorOperacoes = vntValor
    
    Set xmlAux = Nothing

End Function

'Calcula o valor da operações
Private Function flValorOperacoesPorArea(ByVal strItemKey As String, ByVal intValor As enumValoresCalculados) As Variant

Dim strExpression                           As String
Dim vntValor                                As Variant
Dim xmlAux                                  As MSXML2.DOMDocument40

    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAux = IIf(intValor = Camara, xmlArquivoCamara, xmlOperacoes)

    vntValor = 0
    strExpression = flMontarExpressaoCalculoNetOperacoesPorArea(strItemKey, intValor)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlAux, strExpression))

    flValorOperacoesPorArea = vntValor
    
    Set xmlAux = Nothing

End Function

'Configura o perfil de acesso do usuário
Property Get PerfilAcesso() As enumPerfilAcesso
    PerfilAcesso = lngPerfil
End Property

'Configura o perfil de acesso do usuário
Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
    
    lngPerfil = pPerfil

    Select Case pPerfil
        Case enumPerfilAcesso.BackOffice
            
            Me.Caption = "BMD - Liquidação Multilateral (Backoffice)"
            fraTipoMensagem.Visible = True
        
            lblValorMensagem.Visible = (gintTipoBackoffice = enumTipoBackOffice.Corretoras)
            txtValorMensagem.Visible = (gintTipoBackoffice = enumTipoBackOffice.Corretoras)
            txtJustificativa.Enabled = False
        
        Case enumPerfilAcesso.AdmArea
            
            Me.Caption = "BMD - Liquidação Multilateral (Administrador de Área)"
            fraTipoMensagem.Visible = True
        
            lblValorMensagem.Visible = (gintTipoBackoffice = enumTipoBackOffice.Corretoras)
            txtValorMensagem.Visible = (gintTipoBackoffice = enumTipoBackOffice.Corretoras)
            txtJustificativa.Enabled = True
        
        Case enumPerfilAcesso.AdmGeralPrevia
            
            Me.Caption = "BMD - Liquidação Multilateral (Administrador Geral Prévia)"
            fraTipoMensagem.Visible = False
        
            lblValorMensagem.Visible = False
            txtValorMensagem.Visible = False
            txtJustificativa.Enabled = False
        
        Case enumPerfilAcesso.AdmGeral
            
            Me.Caption = "BMD - Liquidação Multilateral (Administrador Geral Definitiva)"
            fraTipoMensagem.Visible = False
    
            lblValorMensagem.Visible = False
            txtValorMensagem.Visible = False
            txtJustificativa.Enabled = False
        
    End Select

    Call flConfigurarBotoesPorPerfil(PerfilAcesso)
    Call flLimparListas
    Call flInicializarlvwDetalhe
    Call flInicializarlvwNet

    If cboEmpresa.ListIndex <> -1 Or cboEmpresa.Text <> vbNullString Then
        Call flPesquisar
    End If

End Property

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    Call flPesquisar
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboEmpresa_Click", Me.Caption

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(lvwNet, Retorno)
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
    Call flInicializarFormulario
    Call flInicializarlvwNet
    Call flInicializarlvwDetalhe
    fgCursor

    Set xmlOperacoes = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlArquivoCamara = CreateObject("MSXML2.DOMDocument.4.0")

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

        .lvwNet.Top = .cboEmpresa.Top + .cboEmpresa.Height + 240
        .lvwNet.Left = .cboEmpresa.Left
        .lvwNet.Height = .imgDummyH.Top - .lvwNet.Top
        .lvwNet.Width = .Width - 240

        .lvwDetalhe.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwDetalhe.Left = .cboEmpresa.Left
        .lvwDetalhe.Height = .Height - .lvwDetalhe.Top - 780 - .fraComplementos.Height
        .lvwDetalhe.Width = .Width - 240
    
        .fraComplementos.Top = .lvwDetalhe.Top + .lvwDetalhe.Height
        .fraComplementos.Left = .cboEmpresa.Left
        .fraComplementos.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set xmlArquivoCamara = Nothing
    Set frmLiquidacaoMultilateralBMF = Nothing
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
        If .imgDummyH.Top > (.Height - 3500) And (.Height - 3500) > 0 Then
            .imgDummyH.Top = .Height - 3500
        End If

        .lvwNet.Height = .imgDummyH.Top - .lvwNet.Top
        .lvwDetalhe.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwDetalhe.Height = .Height - .lvwDetalhe.Top - 780 - .fraComplementos.Height
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDummyH = False
End Sub

Private Sub lvwDetalhe_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwDetalhe, ColumnHeader.Index)
    lngIndexClassifListDet = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwDetalhe_ColumnClick", Me.Caption

End Sub

Private Sub lvwDetalhe_DblClick()

On Error GoTo ErrorHandler

    If PerfilAcesso <> enumPerfilAcesso.AdmArea And PerfilAcesso <> enumPerfilAcesso.BackOffice Then Exit Sub
    
    If Not lvwDetalhe.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Split(lvwDetalhe.SelectedItem.Key, "|")(KEY_DET_NU_SEQU_OPER_ATIV)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwDetalhe_DblClick", Me.Caption

End Sub

Private Sub lvwDetalhe_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim objListItem                             As ListItem

On Error GoTo ErrorHandler
    
    If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Then
        For Each objListItem In lvwNet.ListItems
            objListItem.Checked = False
        Next
        
        With Me.tlbComandos.Buttons
            
            If Item.Checked Then
                .Item("retorno").Enabled = True
                .Item("liberacao").Enabled = False
            Else
                .Item("retorno").Enabled = True
                .Item("liberacao").Enabled = True
            
            End If
            
        End With
        
        
        
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_ItemCheck", Me.Caption

End Sub

Private Sub lvwNet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwNet, ColumnHeader.Index)
    lngIndexClassifListNet = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_ColumnClick", Me.Caption

End Sub

Private Sub lvwNet_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim dblValorAConcordar                      As Double
Dim objListItem                             As ListItem
    
On Error GoTo ErrorHandler
    
    If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Then
        For Each objListItem In lvwDetalhe.ListItems
            objListItem.Checked = False
        Next
        
        With Me.tlbComandos.Buttons
            .Item("retorno").Enabled = False
            .Item("liberacao").Enabled = True
        End With
    End If
    
    If Item.Key = strChaveTotais Then Item.Checked = False
    
    If Item.Checked Then
        Select Case PerfilAcesso
            Case enumPerfilAcesso.BackOffice
                dblValorAConcordar = fgVlrXml_To_Decimal(fgVlr_To_Xml(Item.SubItems(COL_BOAA_NET_VALR_ACON)))
            Case enumPerfilAcesso.AdmGeralPrevia
                dblValorAConcordar = fgVlrXml_To_Decimal(fgVlr_To_Xml(Item.SubItems(COL_AGPV_NET_VALR_SIST)))
        End Select
        
        If dblValorAConcordar = 0 Then
            frmMural.Display = "Valor a ser concordado é igual a 0 (Zero). Seleção do item não permitida."
            frmMural.Show vbModal
            Item.Checked = False
        End If
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_ItemCheck", Me.Caption

End Sub

Private Sub lvwNet_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    If PerfilAcesso = enumPerfilAcesso.AdmArea Or PerfilAcesso = enumPerfilAcesso.BackOffice Then
        Call flCarregarListaDetalheOperacoes
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_ItemClick", Me.Caption

End Sub

Private Sub lvwNet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If PerfilAcesso = enumPerfilAcesso.AdmGeralPrevia Then
        If Button = vbRightButton Then
            ctlMenu1.ShowMenuMarcarDesmarcar
        End If
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_MouseDown", Me.Caption

End Sub

Private Sub optTipoMensagem_Click(Index As Integer)

On Error GoTo ErrorHandler
    
    If Me.cboEmpresa.Text = vbNullString Then Exit Sub
    
    Call flLimparListas
    Call flInicializarlvwNet
    Call flPesquisar
    DoEvents
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optNaturezaMovimento_Click", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
Dim strValidaProcessamento                  As String

On Error GoTo ErrorHandler

    Button.Enabled = False: DoEvents
    intAcaoProcessamento = 0

    Select Case Button.Key
        Case "refresh"
            Call flPesquisar

        Case "concordancia"
            Select Case PerfilAcesso
                Case enumPerfilAcesso.BackOffice
                    intAcaoProcessamento = enumAcaoConciliacao.BOConcordar
                Case enumPerfilAcesso.AdmArea
                    intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar
                Case enumPerfilAcesso.AdmGeralPrevia, enumPerfilAcesso.AdmGeral
                    intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia
            End Select

        Case "discordancia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarDiscordancia
        
        Case "retorno"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar
        
        Case "liberacao"
            Select Case PerfilAcesso
                Case enumPerfilAcesso.AdmGeralPrevia
                    intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento
                
                Case enumPerfilAcesso.AdmGeral
                    If fgVlrXml_To_Decimal(fgVlr_To_Xml(lvwNet.ListItems(1).SubItems(COL_AGDF_NET_VLDL_PGRC))) >= 0 Then
                        intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRecebimento
                    Else
                       intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento
                    End If
                
            End Select
        
        Case "pagamentocontingencia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoContingencia
        
        Case "regularizacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar
        
        Case gstrSair
            Unload Me

    End Select

    If intAcaoProcessamento <> 0 Then
        
        strValidaProcessamento = flValidarItensProcessamento(intAcaoProcessamento)
        
        If strValidaProcessamento <> vbNullString Then
            If InStr(1, strValidaProcessamento, "?") = 0 Then
                frmMural.Display = strValidaProcessamento
                frmMural.IconeExibicao = IconExclamation
                frmMural.Show vbModal
                GoTo ExitSub
                
            Else
                If MsgBox(strValidaProcessamento, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    GoTo ExitSub
                Else
                    If PerfilAcesso = enumPerfilAcesso.AdmArea Then
                        'If Trim$(txtJustificativa.Text) = vbNullString Then
                        '    frmMural.Display = "Informe a justificativa para a divergência de valores."
                        '    frmMural.IconeExibicao = IconExclamation
                        '    frmMural.Show vbModal
                        '
                        '    If txtJustificativa.Enabled Then txtJustificativa.SetFocus
                        '    GoTo ExitSub
                        'End If
                    End If
                End If
            
            End If
        Else
            If MsgBox("Confirma o processamento do(s) item(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        
        End If

        strResultadoOperacao = flProcessar
        
        If strResultadoOperacao <> vbNullString Then
            Call flMostrarResultado(strResultadoOperacao, intAcaoProcessamento)
            Call flPesquisar
        End If
    End If

ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
