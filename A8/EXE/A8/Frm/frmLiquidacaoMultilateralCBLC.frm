VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLiquidacaoMultilateralCBLC 
   Caption         =   "Liquidação de Corretoras"
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
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4350
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
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoMultilateralCBLC.frx":16D8
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
      Left            =   10110
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
         NumButtons      =   9
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
            Caption         =   "Rejeitar          "
            Key             =   "retorno"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Liberação Pag."
            Key             =   "liberacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pg. Conting.   "
            Key             =   "pagamentocontingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Regularizar     "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetalhe 
      Height          =   3525
      Left            =   60
      TabIndex        =   1
      Top             =   4140
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6218
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
Attribute VB_Name = "frmLiquidacaoMultilateralCBLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Liquidação Multilateral CBLC

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40
Private xmlLancamentosCamara                As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Net de Operações - Backoffice e Adm.Área
Private Const COL_BOAA_NET_AGEN_COMP        As Integer = 0
Private Const COL_BOAA_NET_VEIC_LEGA        As Integer = 1
Private Const COL_BOAA_NET_GRUP_LANC        As Integer = 2
Private Const COL_BOAA_NET_VALR_ACON        As Integer = 3
Private Const COL_BOAA_NET_VALR_CONC        As Integer = 4
Private Const COL_BOAA_NET_TOTL_SIST        As Integer = 5
Private Const COL_BOAA_NET_VALR_CAMR        As Integer = 6
Private Const COL_BOAA_NET_DIFE_VALR        As Integer = 7

'Constantes de Configuração de Colunas de Net de Operações - Adm.Geral
Private Const COL_ADMG_NET_AGEN_COMP        As Integer = 0
Private Const COL_ADMG_NET_AREA_RESP        As Integer = 1
Private Const COL_ADMG_NET_VALR_SIST        As Integer = 2
Private Const COL_ADMG_NET_VALR_CAMR        As Integer = 3
Private Const COL_ADMG_NET_DIFE_VALR        As Integer = 4
Private Const COL_ADMG_NET_VALR_LDL1        As Integer = 5
Private Const COL_ADMG_NET_VALR_LDL5        As Integer = 6

'Constantes de posicionamento de campos na propriedade Key do item de Net de Operações - Backoffice e Adm.Área
Private Const KEY_BOAA_NET_AGEN_COMP        As Integer = 1
Private Const KEY_BOAA_NET_VEIC_LEGA        As Integer = 2
Private Const KEY_BOAA_NET_GRUP_LANC        As Integer = 3

'Constantes de posicionamento de campos na propriedade Key do item de Net de Operações - Adm.Geral
Private Const KEY_ADMG_NET_AGEN_COMP        As Integer = 1
Private Const KEY_ADMG_NET_AREA_RESP        As Integer = 2

'Constantes de posicionamento de campos na propriedade Tag do item de Net de Operações - Adm.Geral
Private Const TAG_MSG_NU_CTRL_IF            As Integer = 1
Private Const TAG_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const TAG_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const TAG_MSG_DH_ULTI_ATLZ          As Integer = 4
Private Const TAG_MSG_CO_ULTI_SITU_PROC     As Integer = 5

'Constantes de Configuração de Colunas de Detalhes de Operação - Backoffice e Adm.Área
Private Const COL_BOAA_DET_CODI_OPER        As Integer = 0
Private Const COL_BOAA_DET_DEBT_CRED        As Integer = 1
Private Const COL_BOAA_DET_VALR_SIST        As Integer = 2
Private Const COL_BOAA_DET_STAT_OPER        As Integer = 3

'Constantes de Configuração de Colunas de Detalhes de Operação - Adm.Geral
Private Const COL_ADMG_DET_GRUP_LANC        As Integer = 0
Private Const COL_ADMG_DET_VALR_SIST        As Integer = 1
Private Const COL_ADMG_DET_VALR_CAMR        As Integer = 2
Private Const COL_ADMG_DET_DIFE_VALR        As Integer = 3

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const KEY_DET_NU_SEQU_OPER_ATIV     As Integer = 1

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmLiquidacaoMultilateralCBLC"

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

            .SubItems(COL_ADMG_DET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorMensagem - dblValorOperacao)

            If dblValorMensagem - dblValorOperacao <> 0 Then
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
            If PerfilAcesso = AdmGeral Then
                dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_NET_VALR_SIST)))
                dblValorCamara = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_NET_VALR_CAMR)))

                .SubItems(COL_ADMG_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorCamara - dblValorOperacao)

                If dblValorCamara - dblValorOperacao <> 0 Then
                    .ListSubItems(COL_ADMG_NET_DIFE_VALR).ForeColor = vbRed
                End If
            Else
                dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_TOTL_SIST)))
                dblValorCamara = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_CAMR)))

                .SubItems(COL_BOAA_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblValorCamara - dblValorOperacao)

                If dblValorCamara - dblValorOperacao <> 0 Then
                    .ListSubItems(COL_BOAA_NET_DIFE_VALR).ForeColor = vbRed
                End If
            End If

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
Dim dblCamara                               As Double
Dim dblSistemaOrigem                        As Double
Dim dblDiferenca                            As Double

On Error GoTo ErrorHandler

    dblAConcordar = 0
    dblConcordado = 0
    dblCamara = 0
    dblSistemaOrigem = 0
    dblDiferenca = 0
    
    For Each objListItem In lvwNet.ListItems
        With objListItem
            If .Key <> strChaveTotais Then
                If PerfilAcesso = AdmGeral Then
                    dblSistemaOrigem = dblSistemaOrigem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_NET_VALR_SIST)))
                    dblCamara = dblCamara + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_NET_VALR_CAMR)))
                    dblDiferenca = dblDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_ADMG_NET_DIFE_VALR)))
    
                Else
                    dblAConcordar = dblAConcordar + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_ACON)))
                    dblConcordado = dblConcordado + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_CONC)))
                    dblSistemaOrigem = dblSistemaOrigem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_TOTL_SIST)))
                    dblCamara = dblCamara + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_VALR_CAMR)))
                    dblDiferenca = dblDiferenca + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_BOAA_NET_DIFE_VALR)))
                
                End If
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
        
        If PerfilAcesso = AdmGeral Then
            .SubItems(COL_ADMG_NET_VALR_SIST) = fgVlrXml_To_Interface(dblSistemaOrigem)
            .SubItems(COL_ADMG_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblCamara)
            .SubItems(COL_ADMG_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblDiferenca)
            
            .ListSubItems(COL_ADMG_NET_VALR_SIST).Bold = True
            .ListSubItems(COL_ADMG_NET_VALR_CAMR).Bold = True
            .ListSubItems(COL_ADMG_NET_DIFE_VALR).Bold = True
        
        Else
            .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(dblAConcordar)
            .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(dblConcordado)
            .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(dblSistemaOrigem)
            .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblCamara)
            .SubItems(COL_BOAA_NET_DIFE_VALR) = fgVlrXml_To_Interface(dblDiferenca)
            
            .ListSubItems(COL_BOAA_NET_VALR_ACON).Bold = True
            .ListSubItems(COL_BOAA_NET_VALR_CONC).Bold = True
            .ListSubItems(COL_BOAA_NET_TOTL_SIST).Bold = True
            .ListSubItems(COL_BOAA_NET_VALR_CAMR).Bold = True
            .ListSubItems(COL_BOAA_NET_DIFE_VALR).Bold = True
        
        End If
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

On Error GoTo ErrorHandler

    If lvwNet.SelectedItem Is Nothing Then Exit Sub

    strItemKey = lvwNet.SelectedItem.Key
    lvwDetalhe.ListItems.Clear
    
    If strItemKey = strChaveTotais Then Exit Sub

    For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))

        With lvwDetalhe.ListItems.Add(, "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            If PerfilAcesso = AdmGeral Then
                dblValorCamara = flValorOperacoes(strItemKey, enumValoresCalculados.Camara)
            
                .Text = objDomNode.selectSingleNode("DE_GRUP_LANC_FINC").Text
                .SubItems(COL_ADMG_DET_VALR_SIST) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                .SubItems(COL_ADMG_DET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)

            Else
                .Text = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
                .SubItems(COL_BOAA_DET_DEBT_CRED) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_BOAA_DET_VALR_SIST) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                .SubItems(COL_BOAA_DET_STAT_OPER) = objDomNode.selectSingleNode("DE_SITU_PROC").Text

            End If

        End With

    Next

    Call fgClassificarListview(Me.lvwDetalhe, lngIndexClassifListDet, True)
    
    If PerfilAcesso = AdmGeral Then Call flCalcularDiferencasListViewDetalhe

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
    strRetLeitura = objRemessa.ObterDetalheRemessaCBLC(pstrFiltro, _
                                                       vntCodErro, _
                                                       vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objRemessa = Nothing

    Call xmlLancamentosCamara.loadXML(strRetLeitura)

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaNetArquivoCamara")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheRemessa/*")

            strListItemKey = flMontarChaveItemListview(objDomNode)
            dblValorCamara = flValorOperacoes(strListItemKey, enumValoresCalculados.Camara)

            If Not fgExisteItemLvw(Me.lvwNet, strListItemKey) Then
            
                With lvwNet.ListItems.Add(, strListItemKey)
                    
                    .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                    If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                        .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                    End If
                    
                    If PerfilAcesso = AdmGeral Then
                        .SubItems(COL_ADMG_NET_AREA_RESP) = objDomNode.selectSingleNode("DE_BKOF").Text
                        .SubItems(COL_ADMG_NET_VALR_SIST) = fgVlrXml_To_Interface(0)
                        .SubItems(COL_ADMG_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                        .SubItems(COL_ADMG_NET_VALR_LDL1) = " "
                        .SubItems(COL_ADMG_NET_VALR_LDL5) = " "
                        
                    Else
                        .SubItems(COL_BOAA_NET_VEIC_LEGA) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                        .SubItems(COL_BOAA_NET_GRUP_LANC) = objDomNode.selectSingleNode("DE_GRUP_LANC_FINC").Text
                        .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(0)
                        .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(0)
                        .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(0)
                        .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                    
                    End If
                    
                End With
                
            Else

                With lvwNet.ListItems(strListItemKey)
                    If PerfilAcesso = AdmGeral Then
                        .SubItems(COL_ADMG_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                    Else
                        .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(dblValorCamara)
                    End If
                End With
            End If
        Next
    End If

    Call fgClassificarListview(Me.lvwNet, lngIndexClassifListNet, True)
    
    Set xmlRetLeitura = Nothing

Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetArquivoCamara", 0)

End Sub

'Exibe os detalhes de lista de mensagens
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
Dim objListItem                             As ListItem
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Me.tlbComandos.Buttons.Item("regularizacao").Enabled = False
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(pstrFiltro, _
                                                     vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaNetMensagens")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            
            Set objListItem = Nothing
            
            If objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text <> "1" Then
                strListItemKey = "|" & _
                                 objDomNode.selectSingleNode("CO_PARP_CAMR").Text & "|" & _
                                 objDomNode.selectSingleNode("TP_BKOF").Text
            Else
                strListItemKey = strChaveTotais
            End If
            
            
            If Not fgExisteItemLvw(Me.lvwNet, strListItemKey) Then
                
                If objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text = "1" Then
                    Set objListItem = lvwNet.ListItems.Add(, strListItemKey)
                End If
                
            Else
                Set objListItem = lvwNet.ListItems(strListItemKey)
            End If
            
            If Not objListItem Is Nothing Then
                With objListItem
                    
                    If objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text = "1" Then
                        .Text = "Totais"
                        .Bold = True
                    End If
                    
                    If objDomNode.selectSingleNode("CO_MESG_SPB").Text = "LDL0001" Then
                        strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                                         "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                                         "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                                         "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                                         "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
                                     
                        .Tag = strListItemTag
                        
                        With Me.tlbComandos.Buttons
                            If objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text = "Crédito" Then
                                objListItem.SubItems(COL_ADMG_NET_VALR_LDL1) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                                
                                .Item("liberacao").Caption = "Liberação Rec."
                                .Item("pagamentocontingencia").Enabled = False
                                '.Item("regularizacao").Enabled = False
                            Else
                                objListItem.SubItems(COL_ADMG_NET_VALR_LDL1) = "-" & fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                                
                                .Item("liberacao").Caption = "Liberação Pag."
                                .Item("pagamentocontingencia").Enabled = True
                                .Item("regularizacao").Enabled = True
                            End If
                        End With
                        
                        .ListSubItems(COL_ADMG_NET_VALR_LDL1).Bold = True
                    
                    Else
                        .SubItems(COL_ADMG_NET_VALR_LDL5) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                        .ListSubItems(COL_ADMG_NET_VALR_LDL5).Bold = True
                        Me.tlbComandos.Buttons.Item("regularizacao").Enabled = True
                        
                    End If
                
                End With
            End If
        Next
    End If

    Call fgClassificarListview(Me.lvwNet, lngIndexClassifListNet, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetMensagens", 0)

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

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")

            strListItemKey = flMontarChaveItemListview(objDomNode)

            If Not fgExisteItemLvw(Me.lvwNet, strListItemKey) Then
            
                With lvwNet.ListItems.Add(, strListItemKey)

                    .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    
                    If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                        .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                    End If
                    
                    If PerfilAcesso = AdmGeral Then
                        dblValorSistOrigem = flValorOperacoes(strListItemKey, enumValoresCalculados.SistemaOrigem)
                        
                        .SubItems(COL_ADMG_NET_AREA_RESP) = objDomNode.selectSingleNode("DE_BKOF").Text
                        .SubItems(COL_ADMG_NET_VALR_SIST) = fgVlrXml_To_Interface(dblValorSistOrigem)
                        .SubItems(COL_ADMG_NET_VALR_CAMR) = fgVlrXml_To_Interface(0)
                        .SubItems(COL_ADMG_NET_VALR_LDL1) = " "
                        .SubItems(COL_ADMG_NET_VALR_LDL5) = " "
                        
                    Else
                        dblValorAConcordar = flValorOperacoes(strListItemKey, enumValoresCalculados.AConcordar)
                        dblValorConcordado = flValorOperacoes(strListItemKey, enumValoresCalculados.Concordado)
                        
                        .SubItems(COL_BOAA_NET_VEIC_LEGA) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                        .SubItems(COL_BOAA_NET_GRUP_LANC) = objDomNode.selectSingleNode("DE_GRUP_LANC_FINC").Text
                        .SubItems(COL_BOAA_NET_VALR_ACON) = fgVlrXml_To_Interface(dblValorAConcordar)
                        .SubItems(COL_BOAA_NET_VALR_CONC) = fgVlrXml_To_Interface(dblValorConcordado)
                        .SubItems(COL_BOAA_NET_TOTL_SIST) = fgVlrXml_To_Interface(dblValorAConcordar + dblValorConcordado)
                        .SubItems(COL_BOAA_NET_VALR_CAMR) = fgVlrXml_To_Interface(0)
                    
                    End If
                    
                End With

            End If

        Next
    End If

    Call fgClassificarListview(Me.lvwNet, lngIndexClassifListNet, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
        
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes", 0)

End Sub

'Altera a exibição dos botões de acordo com o perfil do usuário
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)

On Error GoTo ErrorHandler

    With tlbComandos
        .Buttons("concordancia").Visible = True
        .Buttons("retorno").Visible = (PerfilAcesso = enumPerfilAcesso.AdmGeral)
        .Buttons("liberacao").Visible = .Buttons("retorno").Visible
        .Buttons("pagamentocontingencia").Visible = .Buttons("retorno").Visible
        .Buttons("regularizacao").Visible = .Buttons("retorno").Visible
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
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flInicializarFormulario", 0

End Sub

'Formata as colunas da lista de operações
Private Sub flInicializarlvwDetalhe()

On Error GoTo ErrorHandler

    With Me.lvwDetalhe.ColumnHeaders
        
        .Clear
        
        If PerfilAcesso = AdmGeral Then
            .Add , , "Grupo Lançamento Financeiro", 4000
            .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
            .Add , , "Valor Câmara (Arquivo)", 2000, lvwColumnRight
            .Add , , "Diferença", 2000, lvwColumnRight
        Else
            .Add , , "Código Operação", 4000
            .Add , , "D/C", 2000
            .Add , , "Valor Sistema Origem", 1800, lvwColumnRight
            .Add , , "Situação", 3600
        End If
    
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarlvwDetalhe", 0

End Sub

'Formata as colunas da lista de mensagens
Private Sub flInicializarlvwNet()

On Error GoTo ErrorHandler

    lvwNet.CheckBoxes = (PerfilAcesso = AdmGeral)
    
    With Me.lvwNet.ColumnHeaders
        .Clear

        If PerfilAcesso = AdmGeral Then
            .Add , , "Agente Compensação", 2500
            .Add , , "Área", 1500
            .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
            .Add , , "Valor Câmara (Arquivo)", 2000, lvwColumnRight
            .Add , , "Diferença", 2000, lvwColumnRight
            .Add , , "Valor Câmara (LDL0001)", 2400, lvwColumnRight
            .Add , , "Valor Câmara (LDL0005R2)", 2400, lvwColumnRight
        Else
            .Add , , "Agente Compensação", 2000
            .Add , , "Veículo Legal", 2000
            .Add , , "Grupo Lanc. Financeiro", 2000
            .Add , , "Valor a Concordar", 1800, lvwColumnRight
            .Add , , "Valor Concordado", 1800, lvwColumnRight
            .Add , , "Total Sistema Origem", 1800, lvwColumnRight
            .Add , , "Valor Câmara (Arquivo)", 1800, lvwColumnRight
            .Add , , "Diferença", 1600, lvwColumnRight
        End If

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

    If PerfilAcesso = AdmGeral Then
        strListItemKey = "|" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text & _
                         "|" & objDomNode.selectSingleNode("TP_BKOF").Text
    Else
        strListItemKey = "|" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text & _
                         "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                         "|" & objDomNode.selectSingleNode("CO_GRUP_LANC_FINC").Text
    End If

    flMontarChaveItemListview = strListItemKey

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Monta uma expressão XPath para seleção do conteúdo de um documento XML
Private Function flMontarCondicaoNavegacaoXMLOperacoes(ByVal strItemKey As String)

Dim strCondicao                             As String

On Error GoTo ErrorHandler

    If PerfilAcesso = AdmGeral Then
        strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_ADMG_NET_AGEN_COMP) & "' " & _
                                                              " and TP_BKOF='" & Split(strItemKey, "|")(KEY_ADMG_NET_AREA_RESP) & "']"
    Else
        strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_BOAA_NET_AGEN_COMP) & "' " & _
                                                              " and CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_BOAA_NET_VEIC_LEGA) & "' " & _
                                                              " and CO_GRUP_LANC_FINC='" & Split(strItemKey, "|")(KEY_BOAA_NET_GRUP_LANC) & "']"
    End If

    flMontarCondicaoNavegacaoXMLOperacoes = strCondicao

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoes", 0

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String, ByVal intValor As enumValoresCalculados) As String

Dim strDebito                               As String
Dim strCredito                              As String

On Error GoTo ErrorHandler

    If PerfilAcesso = AdmGeral Then
        strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                         " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_ADMG_NET_AGEN_COMP) & "' " & _
                                         " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_ADMG_NET_AREA_RESP) & "' "
    
        strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                          " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_ADMG_NET_AGEN_COMP) & "' " & _
                                          " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_ADMG_NET_AREA_RESP) & "' "
    
    Else
        strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                         " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_BOAA_NET_AGEN_COMP) & "' " & _
                                         " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_BOAA_NET_VEIC_LEGA) & "' " & _
                                         " and ../CO_GRUP_LANC_FINC='" & Split(strItemKey, "|")(KEY_BOAA_NET_GRUP_LANC) & "' "
    
        strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                          " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_BOAA_NET_AGEN_COMP) & "' " & _
                                          " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_BOAA_NET_VEIC_LEGA) & "' " & _
                                          " and ../CO_GRUP_LANC_FINC='" & Split(strItemKey, "|")(KEY_BOAA_NET_GRUP_LANC) & "' "
        
        If PerfilAcesso = BackOffice Then
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
                    strDebito = strDebito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackoffice & "' "
                    strCredito = strCredito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaBackoffice & "' "
                Case enumValoresCalculados.Concordado
                    strDebito = strDebito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaAdmArea & "' "
                    strCredito = strCredito & " and ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.ConcordanciaAdmArea & "' "
            End Select
            
        End If
        
    End If

    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "

    flMontarExpressaoCalculoNetOperacoes = strDebito & strCredito

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoCalculoNetOperacoes", 0

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
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CLBCAcoes)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    If PerfilAcesso = BackOffice Then
        
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
    
    ElseIf PerfilAcesso = AdmArea Then
        
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
    
    ElseIf PerfilAcesso = AdmGeral Then
        
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.RecebimentoLib)
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ControleRepeticao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_ControleRepeticao", "ControleRepeticao", "> 0 ")
        
    End If

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralCBLC)
    Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.NETEntradaManualMultilateralCBLC)

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Multilateral)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL0001")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL0005R2")

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
Dim intColunaValor                          As Integer

On Error GoTo ErrorHandler

    intStatusOper = 0
    Select Case PerfilAcesso
        Case enumPerfilAcesso.BackOffice
            intStatusOper = enumStatusOperacao.EmSer
            intColunaValor = COL_BOAA_NET_VALR_ACON
        
        Case enumPerfilAcesso.AdmArea
            intStatusOper = enumStatusOperacao.ConcordanciaBackoffice
            intColunaValor = COL_BOAA_NET_VALR_ACON
    
        Case enumPerfilAcesso.AdmGeral
            intStatusOper = enumStatusOperacao.ConcordanciaAdmArea
            intColunaValor = COL_ADMG_NET_VALR_SIST
            
            If intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar Then
                If Me.tlbComandos.Buttons.Item("liberacao").Caption = "Liberação Rec." Then
                    intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar
                End If
            End If
        
    End Select
            
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    If PerfilAcesso = AdmGeral And intAcaoProcessamento <> enumAcaoConciliacao.AdmGeralRejeitar Then
        Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")

        Set objListItem = lvwNet.ListItems(strChaveTotais)
        With objListItem
            Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "NU_CTRL_IF", Split(.Tag, "|")(TAG_MSG_NU_CTRL_IF))
            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "DH_REGT_MESG_SPB", Split(.Tag, "|")(TAG_MSG_DH_REGT_MESG_SPB))
            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "NU_SEQU_CNTR_REPE", Split(.Tag, "|")(TAG_MSG_NU_SEQU_CNTR_REPE))
            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "DH_ULTI_ATLZ", Split(.Tag, "|")(TAG_MSG_DH_ULTI_ATLZ))
            Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "CO_ULTI_SITU_PROC", Split(.Tag, "|")(TAG_MSG_CO_ULTI_SITU_PROC))
        End With
    
    End If
    
    For Each objListItem In Me.lvwNet.ListItems
        With objListItem
            
            If (.Checked And intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar) Or _
                             intAcaoProcessamento <> enumAcaoConciliacao.AdmGeralRejeitar Then
                
                If (.SubItems(intColunaValor) <> fgVlrXml_To_Interface(0) Or _
                    intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizarRecebimento) And _
                    .Key <> strChaveTotais Then
    
                    If PerfilAcesso <> AdmGeral Or intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar Then
                        Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")
                        Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
                    End If
                    
                    If .Index = 1 Or PerfilAcesso <> AdmGeral Or _
                    intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar Then
                        
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "CO_EMPR", fgObterCodigoCombo(cboEmpresa.Text))
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "CO_PARP_CAMR", Split(.Key, "|")(KEY_ADMG_NET_AGEN_COMP))
                        intIgnoraGradeHorario = 1
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "IgnoraGradeHorario", intIgnoraGradeHorario)
                        Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", "Repeat_Operacao", "")
        
                    End If
                    
                    For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItem.Key))
                        
                        If intStatusOper = Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) Then
                            Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                            Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
                            Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", "NU_SEQU_OPER_ATIV", objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                            Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", "DH_ULTI_ATLZ", objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
                            Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)
                            Set xmlItemOperacao = Nothing
                        End If
                    Next
    
                End If
            
                If (.Index = lvwNet.ListItems.Count - 1 Or _
                    PerfilAcesso <> AdmGeral Or _
                    intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar) And _
                    Not xmlItemEnvioMsg Is Nothing Then
                    
                    Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
                    Set xmlItemEnvioMsg = Nothing
                    
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

    If PerfilAcesso = AdmGeral Then
        Call flCarregarListaNetMensagens(strDocFiltros)
    End If

    Call flCalcularDiferencasListViewNet
    
    If lvwNet.ListItems.Count > 0 Then
        lvwNet.ListItems(1).Selected = True
        Call lvwNet_ItemClick(lvwNet.ListItems(1))
    End If

    Call flCalcularTotais
    
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

    strXMLProc = flMontarXMLProcessamento

    If strXMLProc <> vbNullString Then
        fgCursor True
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.ProcessarLoteLiquidacaoMultilateralCBLC(intAcaoProcessamento, _
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
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
        
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flProcessar", Me.Caption

End Function

'Valida a seleção dos itens na tela, para posterior processamento
Private Function flValidarItensProcessamento(ByVal intAcao As enumAcaoConciliacao) As String

Dim intLinhas                               As Integer
Dim objListItem                             As ListItem
Dim intAcaoLDL0001                          As enumTipoAcao

    If PerfilAcesso = BackOffice Then
        intLinhas = 0
        For Each objListItem In lvwNet.ListItems
            If objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> objListItem.SubItems(COL_BOAA_NET_VALR_CAMR) And _
               objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> fgVlrXml_To_Interface(0) Then
                intLinhas = intLinhas + 1
                Exit For
            End If
        Next
        
        If intLinhas > 0 Then
            flValidarItensProcessamento = "Valor A concordar é diferente do valor enviado pela câmara. Deseja prosseguir com a operação ?"
            Exit Function
        End If
    
        intLinhas = 0
        For Each objListItem In lvwNet.ListItems
            If objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> fgVlrXml_To_Interface(0) Then
                intLinhas = intLinhas + 1
                Exit For
            End If
        Next
        
        If intLinhas = 0 Then
            flValidarItensProcessamento = "Todos os valores pendentes de concordância, para o Backoffice, já foram processados."
            Exit Function
        End If
    
    ElseIf PerfilAcesso = AdmArea Then
        intLinhas = 0
        For Each objListItem In lvwNet.ListItems
            If objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> objListItem.SubItems(COL_BOAA_NET_VALR_CAMR) And _
               objListItem.SubItems(COL_BOAA_NET_VALR_ACON) <> fgVlrXml_To_Interface(0) Then
                intLinhas = intLinhas + 1
                Exit For
            End If
        Next
        
        If intLinhas > 0 Then
            flValidarItensProcessamento = "Valor A concordar é diferente do valor enviado pela câmara. Deseja prosseguir com a operação ?"
            Exit Function
        End If
    
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
    
    ElseIf PerfilAcesso = AdmGeral Then
        Select Case intAcao
            
            Case enumAcaoConciliacao.AdmGeralRejeitar
                
                If fgItemsCheckedListView(Me.lvwNet) = 0 Then
                    flValidarItensProcessamento = "Selecione pelo menos um item da lista, antes de prosseguir com a operação desejada."
                    Exit Function
                End If
        
        
                If Me.tlbComandos.Buttons.Item("liberacao").Caption = "Liberação Rec." Then
                    If flVerificaAcaoLDL0001(enumTipoAcao.EnviadoRecebimento) Then
                        flValidarItensProcessamento = "Recebimento já foi Liberado. Operação não permitida."
                        Exit Function
                    End If
                End If
        
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                 enumAcaoConciliacao.AdmGeralPagamento, _
                 enumAcaoConciliacao.AdmGeralRegularizar
                 
                If Me.tlbComandos.Buttons.Item("liberacao").Caption = "Liberação Rec." Then
                    If flVerificaAcaoLDL0001(enumTipoAcao.EnviadoRecebimento) Then
                        flValidarItensProcessamento = "Recebimento já foi Liberado. Operação não permitida."
                        Exit Function
                    End If
                End If
                 
                If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) = vbNullString Then
                    flValidarItensProcessamento = "Valor da LDL0001 não encontrado. Operação não permitida."
                    Exit Function
                End If
                
                If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_SIST)) <> _
                   Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) Then
                    flValidarItensProcessamento = "Total do Sistema Origem difere do Total da LDL0001. Operação não permitida."
                    Exit Function
                End If
                
                
            Case enumAcaoConciliacao.AdmGeralRecebimento
                
                If flVerificaAcaoLDL0001(enumTipoAcao.EnviadoRecebimento) Then
                    flValidarItensProcessamento = "Recebimento já foi Liberado. Operação não permitida."
                    Exit Function
                Else
                    
                    If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) = vbNullString Then
                        flValidarItensProcessamento = "Valor da LDL0001 não encontrado. Operação não permitida."
                        Exit Function
                    End If
                    
                    If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_SIST)) <> _
                       Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) Then
                        flValidarItensProcessamento = "Total do Sistema Origem difere do Total da LDL0001. Operação não permitida."
                        Exit Function
                    End If
                End If
                
            
            Case enumAcaoConciliacao.AdmGeralRegularizarRecebimento
                
                If Not flVerificaAcaoLDL0001(enumTipoAcao.EnviadoRecebimento) Then
                    flValidarItensProcessamento = "Recebimento não foi Liberado. Operação não permitida."
                    Exit Function
                Else
                    If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL5)) <> _
                       Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) Then
                        flValidarItensProcessamento = "Valor da mensagem LDL0001 divergente da LDL0005R2. Regularização não permitida."
                        Exit Function
                    End If
                
                End If
                
            Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                 
                If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) = vbNullString Then
                    flValidarItensProcessamento = "Valor da LDL0001 não encontrado. Operação não permitida."
                    Exit Function
                End If
                
                If Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_SIST)) = _
                   Trim$(lvwNet.ListItems(strChaveTotais).SubItems(COL_ADMG_NET_VALR_LDL1)) Then
                    flValidarItensProcessamento = "Total do Sistema Origem é igual ao Total da LDL0001. Pagamento em contingência não permitido."
                    Exit Function
                End If
            
        End Select
        
    End If

End Function

'Calcula o valor da operações
Private Function flValorOperacoes(ByVal strItemKey As String, ByVal intValor As enumValoresCalculados) As Variant

Dim strExpression                           As String
Dim vntValor                                As Variant
Dim xmlAux                                  As MSXML2.DOMDocument40

    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAux = IIf(intValor = Camara, xmlLancamentosCamara, xmlOperacoes)

    vntValor = 0
    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey, intValor)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlAux, strExpression))

    flValorOperacoes = vntValor
    
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
            Me.Caption = "CBLC - Liquidação Multilateral (Backoffice)"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "CBLC - Liquidação Multilateral (Administrador de Área)"
        Case enumPerfilAcesso.AdmGeral
            Me.Caption = "CBLC - Liquidação Multilateral (Administrador Geral)"
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
    Set xmlLancamentosCamara = CreateObject("MSXML2.DOMDocument.4.0")

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
        .lvwDetalhe.Height = .Height - .lvwDetalhe.Top - 720
        .lvwDetalhe.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set xmlLancamentosCamara = Nothing
    Set frmLiquidacaoMultilateralCBLC = Nothing
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

        .lvwNet.Height = .imgDummyH.Top - .lvwNet.Top
        .lvwDetalhe.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwDetalhe.Height = .Height - .lvwDetalhe.Top - 720
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDummyH = False
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

On Error GoTo ErrorHandler
    
    If Item.Key = strChaveTotais Then Item.Checked = False
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_ItemCheck", Me.Caption

End Sub

Private Sub lvwNet_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call flCarregarListaDetalheOperacoes
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_ItemClick", Me.Caption

End Sub

Private Sub lvwNet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If PerfilAcesso = AdmGeral Then
        If Button = vbRightButton Then
            ctlMenu1.ShowMenuMarcarDesmarcar
        End If
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwNet_MouseDown", Me.Caption

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
            If PerfilAcesso = BackOffice Then
                intAcaoProcessamento = enumAcaoConciliacao.BOConcordar
            ElseIf PerfilAcesso = AdmArea Then
                intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar
            ElseIf PerfilAcesso = AdmGeral Then
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia
            End If

        Case "retorno"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar
        
        Case "liberacao"
            
            If Button.Caption = "Liberação Rec." Then
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRecebimento
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento
            End If
        
        Case "pagamentocontingencia"
            
            If Button.Caption = "Liberação Rec." Then
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRecebimento
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoContingencia
            End If
        
        Case "regularizacao"
            
            If Me.tlbComandos.Buttons.Item("liberacao").Caption = "Liberação Rec." Then
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizarRecebimento
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar
            End If
            
        
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

'KIDA - CBLC - 24/09/2008
Private Function flVerificaAcaoLDL0001(ByVal pintAcaoLDL0001 As enumTipoAcao) As Boolean

Dim objListItem                             As ListItem
Dim blnPagtoLiberado                        As Boolean
Dim xmlHistMesg                             As MSXML2.DOMDocument40
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strHistMesg                             As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim strCondicao                             As String
Dim blnAcaoPermitida                        As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If
    
    vntCodErro = "0"
    vntMensagemErro = ""
    
    Set objListItem = lvwNet.ListItems(strChaveTotais)
    
    If Not objListItem Is Nothing Then
    
        
        
        With objListItem
            If .Tag <> "" Then
                Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
                Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SequenciaOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_SequenciaOperacao", "SequenciaOperacao", "0")
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_NumeroCtrlIF", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_NumeroCtrlIF", "NumeroCtrlIF", Split(.Tag, "|")(TAG_MSG_NU_CTRL_IF))
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataRegistro", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_DataRegistro", "DataRegistro", Split(.Tag, "|")(TAG_MSG_DH_REGT_MESG_SPB)) 'fgDtHr_To_Xml(datDataRegistroMensagem))
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_NumeroSequenciaControleRepeticao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_NumeroSequenciaControleRepeticao", "NumeroSequenciaControleRepeticao", 1)
            
                Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
                strHistMesg = objMensagem.ObterHistoricoMensagem(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                Set objMensagem = Nothing
                
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
    
                If strHistMesg <> vbNullString Then
                    
                    Set xmlHistMesg = CreateObject("MSXML2.DOMDocument.4.0")
                    xmlHistMesg.loadXML strHistMesg
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & pintAcaoLDL0001 & "']"
                    blnAcaoPermitida = Not xmlHistMesg.selectNodes(strCondicao).length = 0
                
                End If
                
            Else
                blnAcaoPermitida = True
            End If
        End With
    End If

    flVerificaAcaoLDL0001 = blnAcaoPermitida

    Set xmlHistMesg = Nothing
    Set xmlDomFiltros = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlHistMesg = Nothing
    Set xmlDomFiltros = Nothing
    Set objMensagem = Nothing
    
    If vntCodErro <> "0" Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    fgRaiseError App.EXEName, TypeName(Me), "flVerificaAcaoLDL0001", 0

End Function


