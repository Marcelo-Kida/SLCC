VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLiquidacaoCorretoras 
   Caption         =   "Liquida��o de Corretoras"
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
   Begin VB.Frame fraControles 
      Height          =   1215
      Left            =   60
      TabIndex        =   8
      Top             =   7080
      Width           =   12855
      Begin VB.TextBox txtComentario 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1440
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   5925
      End
      Begin VB.ComboBox cboTipoJustificativa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   5940
      End
      Begin VB.Label lblComentario 
         AutoSize        =   -1  'True
         Caption         =   "Coment�rio"
         Height          =   165
         Left            =   540
         TabIndex        =   12
         Top             =   690
         Width           =   795
      End
      Begin VB.Label lblTipoJustificativa 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Justificativa"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   330
         Width           =   1185
      End
   End
   Begin VB.Frame fraNaturezaMovto 
      Caption         =   "Natureza Movimento"
      Height          =   570
      Left            =   4500
      TabIndex        =   5
      Top             =   60
      Width           =   2895
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "&Recebimento"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "Pa&gamento"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
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
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8295
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   582
      ButtonWidth     =   3440
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela             "
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Concordar                 "
            Key             =   "concordancia"
            Object.ToolTipText     =   "Concodar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Liberar                      "
            Key             =   "liberacao"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Liberar Contig�ncia  "
            Key             =   "liberacaocontingencia"
            Object.ToolTipText     =   "Liberar Contig�ncia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Rejeitar                    "
            Key             =   "retorno"
            Object.ToolTipText     =   "Retornar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                          "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   2295
      Left            =   60
      TabIndex        =   2
      Top             =   4140
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   4048
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Conciliar Opera��o"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "N�mero Comando"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "N�mero Comando Original"
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
            Picture         =   "frmLiquidacaoCorretoras.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoCorretoras.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3255
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin A8.ctlMenu ctlMenu1 
      Left            =   10110
      Top             =   0
      _ExtentX        =   2990
      _ExtentY        =   661
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
      TabIndex        =   4
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmLiquidacaoCorretoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Liquida��o de opera��es efetuadas em corretoras

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40

'Constantes de Configura��o de Colunas de Mensagem
Private Const COL_MSG_LOCA_LIQU             As Integer = 0
Private Const COL_MSG_ISPB_IF               As Integer = 1
Private Const COL_MSG_CNPJ_CORR             As Integer = 2
Private Const COL_MSG_NOME_CORR             As Integer = 3
Private Const COL_MSG_CNPT                  As Integer = 4
Private Const COL_MSG_AGEN                  As Integer = 5
Private Const COL_MSG_CNTA_CORR             As Integer = 6
Private Const COL_MSG_VALR                  As Integer = 7
Private Const COL_MSG_VALR_CAMR             As Integer = 7
Private Const COL_MSG_VALR_OPER             As Integer = 8
Private Const COL_MSG_STAT                  As Integer = 8
Private Const COL_MSG_DIFE                  As Integer = 9

'Perfil de Acesso Backofice
Private Const COL_MSG_ID_LOTE               As Integer = 9
Private Const COL_MSG_SG_SIST               As Integer = 10
Private Const COL_MSG_DATA_OPERACAO         As Integer = 11
Private Const COL_MSG_TP_DEB_CRED_LOTE      As Integer = 12
Private Const COL_MSG_VA_TOT_LOTE           As Integer = 13
Private Const COL_MSG_QT_OPER_LOTE          As Integer = 14
Private Const COL_MSG_QT_LOTE_APUR_SLCC     As Integer = 15

'Outros Perfis de Acesso
'Pagamento
Private Const COL_MSG_ID_LOTE_PGTO               As Integer = 8
Private Const COL_MSG_SG_SIST_PGTO               As Integer = 9
Private Const COL_MSG_DATA_OPERACAO_PGTO         As Integer = 10
Private Const COL_MSG_TP_DEB_CRED_LOTE_PGTO      As Integer = 11
Private Const COL_MSG_VA_TOT_LOTE_PGTO           As Integer = 12
Private Const COL_MSG_QT_OPER_LOTE_PGTO          As Integer = 13
Private Const COL_MSG_QT_LOTE_APUR_SLCC_PGTO     As Integer = 14

'Recebimento
Private Const COL_MSG_ID_LOTE_RECB               As Integer = 10
Private Const COL_MSG_SG_SIST_RECB               As Integer = 11
Private Const COL_MSG_DATA_OPERACAO_RECB         As Integer = 12
Private Const COL_MSG_TP_DEB_CRED_LOTE_RECB      As Integer = 13
Private Const COL_MSG_VA_TOT_LOTE_RECB           As Integer = 14
Private Const COL_MSG_QT_OPER_LOTE_RECB          As Integer = 15
Private Const COL_MSG_QT_LOTE_APUR_SLCC_RECB     As Integer = 16




'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const KEY_MSG_CO_LOCA_LIQU          As Integer = 1
Private Const KEY_MSG_CO_ISPB_CNPT          As Integer = 2
Private Const KEY_MSG_CO_CNPJ_CNPT          As Integer = 3
Private Const KEY_MSG_CO_AGEN_COTR          As Integer = 4
Private Const KEY_MSG_NU_CC_COTR            As Integer = 5
Private Const KEY_MSG_CO_ULTI_SITU_PROC     As Integer = 6
Private Const KEY_MSG_TP_IF_CRED_DEBT       As Integer = 7
Private Const KEY_MSG_ID_LOTE               As Integer = 8
Private Const KEY_MSG_SG_SIST               As Integer = 9
Private Const KEY_MSG_DATA_OPERACAO         As Integer = 10
Private Const KEY_MSG_NU_CTRL_IF            As Integer = 11

'Constantes de posicionamento de campos na propriedade Tag do item do ListView de Mensagens
Private Const TAG_MSG_NU_CTRL_IF            As Integer = 1
Private Const TAG_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const TAG_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const TAG_MSG_DH_ULTI_ATLZ          As Integer = 4

'Constantes de Configura��o de Colunas de Opera��o
Private Const COL_OP_CLIENTE                As Integer = 0
Private Const COL_OP_ID_TITULO              As Integer = 1
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 2
Private Const COL_OP_DC                     As Integer = 3
Private Const COL_OP_VALOR                  As Integer = 4
Private Const COL_OP_QUANTIDADE             As Integer = 5
Private Const COL_OP_PU                     As Integer = 6
Private Const COL_OP_STATUS                 As Integer = 7
Private Const COL_OP_NUMERO_COMANDO         As Integer = 8
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 9
Private Const COL_OP_DATA_OPERACAO          As Integer = 10
Private Const COL_OP_CODIGO                 As Integer = 11
Private Const COL_OP_ID_LOTE                As Integer = 12
Private Const COL_OP_SG_SIST                As Integer = 13
Private Const COL_OP_DT_LOTE                As Integer = 14
Private Const COL_OP_TP_DEB_CRED_LOTE       As Integer = 15
Private Const COL_OP_VA_TOT_LOTE            As Integer = 16
Private Const COL_OP_QT_OPER_LOTE           As Integer = 17


'Constantes de posicionamento de campos na propriedade Key do item do ListView de Opera��es
Private Const KEY_OP_NU_SEQU_OPER_ATIV      As Integer = 1

'Constante para o Mapa de Navega��o
Private Const strFuncionalidade             As String = "frmLiquidacaoCorretoras"
'------------------------------------------------------------------------------------------
'Fim declara��o constantes

Private Enum enumNaturezaMovimento
    Pagamento = 0
    Recebimento = 1
End Enum

Private Enum enumTipoPesquisa
    operacao = 0
    MENSAGEM = 1
End Enum

Private intAcaoProcessamento                As enumAcaoConciliacao

Private lngPerfil                           As Long
Private blnDummyH                           As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Calcula as diferen�as entre os valores de opera��o e mensagem
Private Sub flCalcularDiferencasListView()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

On Error GoTo ErrorHandler
    
    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            
            
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VALR_OPER)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VALR_CAMR)))
            
            .SubItems(COL_MSG_DIFE) = fgVlrXml_To_Interface(dblValorMensagem - dblValorOperacao)
            
            If dblValorMensagem - dblValorOperacao <> 0 Then
                .ListSubItems(COL_MSG_DIFE).ForeColor = vbRed
            End If
        End With
    Next

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListView", 0)

End Sub

'Mostra os campos de detalhes das opera��es
Private Sub flCarregarListaDetalheOperacoes()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strItemKey                              As String

On Error GoTo ErrorHandler
    
    If lvwMensagem.SelectedItem Is Nothing Then Exit Sub
    
    strItemKey = lvwMensagem.SelectedItem.Key
    lvwOperacao.ListItems.Clear

    For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))

        With lvwOperacao.ListItems.Add(, "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
            .SubItems(COL_OP_CODIGO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            .SubItems(COL_OP_DC) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
            .SubItems(COL_OP_ID_TITULO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
            .SubItems(COL_OP_QUANTIDADE) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("QT_ATIV_MERC").Text)
            .SubItems(COL_OP_PU) = fgVlrXml_To_InterfaceDecimais(objDomNode.selectSingleNode("PU_ATIV_MERC").Text, 8)
            .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
            .SubItems(COL_OP_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
            .SubItems(COL_OP_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
            
            If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
            End If
            
            If objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_LIQUIDACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
            End If
            
            If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
            End If
            .SubItems(COL_OP_ID_LOTE) = objDomNode.selectSingleNode("CD_LOTE").Text
            .SubItems(COL_OP_SG_SIST) = objDomNode.selectSingleNode("SG_SIST").Text
            .SubItems(COL_OP_DT_LOTE) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
            .SubItems(COL_OP_TP_DEB_CRED_LOTE) = objDomNode.selectSingleNode("TP_DEBT_CRED_LOTE").Text
            .SubItems(COL_OP_VA_TOT_LOTE) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VL_TOTL_LOTE").Text)
            .SubItems(COL_OP_QT_OPER_LOTE) = objDomNode.selectSingleNode("QT_TOTL_LOTE").Text
        End With

    Next

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaDetalheOperacoes", 0)

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
Dim strListItemTag
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

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
            strListItemKey = flMontarChaveItemListview(objDomNode)
                        
            strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
                             
            If Not flExisteItemLvwMensagem(Me.lvwMensagem, strListItemKey, Trim(objDomNode.selectSingleNode("CO_MESG_SPB").Text)) Then
                
                With lvwMensagem.ListItems.Add(, strListItemKey)
                    
                    .Text = objDomNode.selectSingleNode("SG_LOCA_LIQU").Text
                    .SubItems(COL_MSG_ISPB_IF) = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
                    .SubItems(COL_MSG_CNPJ_CORR) = objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text
                    .SubItems(COL_MSG_NOME_CORR) = objDomNode.selectSingleNode("NO_CNPT_CNCL").Text
                    .SubItems(COL_MSG_AGEN) = objDomNode.selectSingleNode("CO_AGEN_COTR").Text
                    .SubItems(COL_MSG_CNTA_CORR) = objDomNode.selectSingleNode("NU_CC_COTR").Text
                    .SubItems(COL_MSG_VALR_CAMR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    
                End With
                
            Else
                With lvwMensagem.ListItems(strListItemKey)
                    
                    'verifica se j� existe uma Mensagem SPB conciliada � esta opera��o, se Sim, ent�o acrescenta uma linha nova para a MensagemSPB atual
                    'If .Tag = "" Then 'n�o havia Mensagem conciliada
                        
                        .SubItems(COL_MSG_VALR_CAMR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                        
                        'atualiza a key
                        strListItemKey = strListItemKey & Trim(objDomNode.selectSingleNode("NU_CTRL_IF").Text)
                        .Key = strListItemKey
                    
                    'Else 'j� h� outra Mensagem conciliada
                        
'                        With lvwMensagem.ListItems.Add(, strListItemKey)
'
'                            .Text = objDomNode.selectSingleNode("SG_LOCA_LIQU").Text
'                            .SubItems(COL_MSG_ISPB_IF) = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
'                            .SubItems(COL_MSG_CNPJ_CORR) = objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text
'                            .SubItems(COL_MSG_NOME_CORR) = objDomNode.selectSingleNode("NO_CNPT_CNCL").Text
'                            .SubItems(COL_MSG_AGEN) = objDomNode.selectSingleNode("CO_AGEN_COTR").Text
'                            .SubItems(COL_MSG_CNTA_CORR) = objDomNode.selectSingleNode("NU_CC_COTR").Text
'                            .SubItems(COL_MSG_VALR_CAMR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
'
'                        End With
                    
                    'End If
                    
                End With
            End If
            
            lvwMensagem.ListItems(strListItemKey).Tag = strListItemTag
            
        Next
    End If
    
    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetMensagens", 0)

End Sub

'Verifica se existe o item no list view mensagem
Public Function flExisteItemLvwMensagem(ByRef objListView As MSComctlLib.ListView, _
                                        ByRef pstrKeyMensagem As String, _
                                        ByVal pMensagemSPB As String) As Boolean
Dim objListItem                             As ListItem
Dim strKeyOperacaoOriginal                  As String
Dim strKeyOperacao                          As String
Dim strKeyMensagemSemNumCtrlIF              As String

On Error GoTo ErrorHandler
                                
    For Each objListItem In objListView.ListItems
        
        'para compara��o da key da opera��o e da mensagem, considerar apenas os 7 primeiros campos das keys (se mensagemSPB for STR008R2 ou PAG0108R2), ou se
        'visto que do 8� campo em diante (CD_LOTE, SG_SIST, DT_OPER_ATIV, NU_CTRL_IF) n�o fazem parte da chave de concilia��o da opera��o com mensagem R2
        strKeyOperacaoOriginal = objListItem.Key
        
        If pMensagemSPB = "STR0006R2" Then
        
            strKeyOperacao = "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_LOCA_LIQU) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_ISPB_CNPT) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_CNPJ_CNPT) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_AGEN_COTR) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "|||"
    
            'montar key da mensagem sem o numerocontroleif no final
            strKeyMensagemSemNumCtrlIF = "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_LOCA_LIQU) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_ISPB_CNPT) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_CNPJ_CNPT) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_AGEN_COTR) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "|||"
        
        ElseIf pMensagemSPB = "STR0007R2" Then
        
            strKeyOperacao = "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_LOCA_LIQU) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_ISPB_CNPT) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "|||"
    
            'montar key da mensagem sem o numerocontroleif no final
            strKeyMensagemSemNumCtrlIF = "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_LOCA_LIQU) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_ISPB_CNPT) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "|||"
        
        
        Else 'pMensagemSPB = "STR0008R2" Or pMensagemSPB = "PAG0108R2" Then
        
            strKeyOperacao = "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_LOCA_LIQU) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_ISPB_CNPT) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_CNPJ_CNPT) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_AGEN_COTR) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_NU_CC_COTR) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & _
                             "|" & Split(strKeyOperacaoOriginal, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "|||"
    
            'montar key da mensagem sem o numerocontroleif no final
            strKeyMensagemSemNumCtrlIF = "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_LOCA_LIQU) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_ISPB_CNPT) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_CNPJ_CNPT) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_AGEN_COTR) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_NU_CC_COTR) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & _
                                         "|" & Split(pstrKeyMensagem, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "|||"
        
        
        End If

        If strKeyOperacao = strKeyMensagemSemNumCtrlIF Then
        
            If Split(strKeyOperacaoOriginal, "|")(KEY_MSG_NU_CTRL_IF) = "" Then 'garante que n�o h� outra MensagemSPB conciliada a este lote
                flExisteItemLvwMensagem = True
                pstrKeyMensagem = strKeyOperacaoOriginal
                Exit Function
            End If
            
        End If
            
    Next
                                
Exit Function
ErrorHandler:

    flExisteItemLvwMensagem = False
    Err.Clear

End Function
'Carregar dados com NET de opera��es
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
Dim dblValorOperacao                        As Double
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim dbOperacoesLote                         As Double

On Error GoTo ErrorHandler

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(pstrFiltro, _
                                                     vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes")
        End If
        
        Call xmlOperacoes.loadXML(strRetLeitura)
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")
            
            strListItemKey = flMontarChaveItemListview(objDomNode)
                    
            If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                dblValorOperacao = flValorOperacoes(strListItemKey)
                
                dbOperacoesLote = flQtdeOperLote(strListItemKey)
            
                If (dblValorOperacao < 0 And optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value) Or _
                   (dblValorOperacao >= 0 And optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value) Then
                    
                    With lvwMensagem.ListItems.Add(, strListItemKey)
                        
                        .Text = objDomNode.selectSingleNode("SG_LOCA_LIQU").Text
                        .SubItems(COL_MSG_ISPB_IF) = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
                        .SubItems(COL_MSG_CNPJ_CORR) = objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text
                        .SubItems(COL_MSG_NOME_CORR) = objDomNode.selectSingleNode("NO_CNPT").Text
                        .SubItems(COL_MSG_AGEN) = objDomNode.selectSingleNode("CO_AGEN_COTR").Text
                        .SubItems(COL_MSG_CNTA_CORR) = objDomNode.selectSingleNode("NU_CC_COTR").Text
                        
                        Select Case Val(objDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text)
                            Case 1
                                .SubItems(COL_MSG_CNPT) = "Externa"
                            Case 2
                                .SubItems(COL_MSG_CNPT) = "Interna"
                        End Select
                        
                        If PerfilAcesso = BackOffice Then
                            .SubItems(COL_MSG_STAT) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                            .SubItems(COL_MSG_VALR) = fgVlrXml_To_Interface(dblValorOperacao)
                            .SubItems(COL_MSG_ID_LOTE) = objDomNode.selectSingleNode("CD_LOTE").Text
                            .SubItems(COL_MSG_SG_SIST) = objDomNode.selectSingleNode("SG_SIST").Text
                            .SubItems(COL_MSG_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                            .SubItems(COL_MSG_TP_DEB_CRED_LOTE) = objDomNode.selectSingleNode("TP_DEBT_CRED_LOTE").Text
                            .SubItems(COL_MSG_VA_TOT_LOTE) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VL_TOTL_LOTE").Text)
                            .SubItems(COL_MSG_QT_OPER_LOTE) = objDomNode.selectSingleNode("QT_TOTL_LOTE").Text
                            'RATS 1212 - 20/09/2012
                            .SubItems(COL_MSG_QT_LOTE_APUR_SLCC) = dbOperacoesLote
                        Else
                            If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                                .SubItems(COL_MSG_VALR) = fgVlrXml_To_Interface(dblValorOperacao)
                                .SubItems(COL_MSG_ID_LOTE_PGTO) = objDomNode.selectSingleNode("CD_LOTE").Text
                                .SubItems(COL_MSG_SG_SIST_PGTO) = objDomNode.selectSingleNode("SG_SIST").Text
                                .SubItems(COL_MSG_DATA_OPERACAO_PGTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                                .SubItems(COL_MSG_TP_DEB_CRED_LOTE_PGTO) = objDomNode.selectSingleNode("TP_DEBT_CRED_LOTE").Text
                                .SubItems(COL_MSG_VA_TOT_LOTE_PGTO) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VL_TOTL_LOTE").Text)
                                .SubItems(COL_MSG_QT_OPER_LOTE_PGTO) = objDomNode.selectSingleNode("QT_TOTL_LOTE").Text
                                'RATS 1212 - 20/09/2012
                                .SubItems(COL_MSG_QT_LOTE_APUR_SLCC_PGTO) = dbOperacoesLote
                            Else
                                .SubItems(COL_MSG_VALR_CAMR) = " "
                                .SubItems(COL_MSG_VALR_OPER) = fgVlrXml_To_Interface(dblValorOperacao)
                                .SubItems(COL_MSG_ID_LOTE_RECB) = objDomNode.selectSingleNode("CD_LOTE").Text
                                .SubItems(COL_MSG_SG_SIST_RECB) = objDomNode.selectSingleNode("SG_SIST").Text
                                .SubItems(COL_MSG_DATA_OPERACAO_RECB) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                                .SubItems(COL_MSG_TP_DEB_CRED_LOTE_RECB) = objDomNode.selectSingleNode("TP_DEBT_CRED_LOTE").Text
                                .SubItems(COL_MSG_VA_TOT_LOTE_RECB) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VL_TOTL_LOTE").Text)
                                .SubItems(COL_MSG_QT_OPER_LOTE_RECB) = objDomNode.selectSingleNode("QT_TOTL_LOTE").Text
                                'RATS 1212 - 20/09/2012
                                .SubItems(COL_MSG_QT_LOTE_APUR_SLCC_RECB) = dbOperacoesLote
                            End If
                        End If
                        

                    End With
                    
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
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes", 0)

End Sub

'Altera a exibi��o dos bot�es de acordo com o perfil do usu�rio
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)
    
On Error GoTo ErrorHandler
    
    With tlbComandos
        .Buttons("concordancia").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice)
        .Buttons("liberacao").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("liberacaocontingencia").Visible = .Buttons("liberacao").Visible
        .Buttons("retorno").Visible = .Buttons("liberacao").Visible
        
        tlbComandos.Buttons("liberacaocontingencia").Enabled = IIf(PerfilAcesso = enumPerfilAcesso.AdmArea, True, False)
        cboTipoJustificativa.Enabled = IIf(PerfilAcesso = enumPerfilAcesso.AdmArea, True, False)
        txtComentario.Enabled = IIf(PerfilAcesso = enumPerfilAcesso.AdmArea, True, False)
        
        .Refresh
    End With
    
Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorPerfil", 0

End Sub

'Inicializa controles de tela e vari�veis
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
    Call fgCarregarCombos(Me.cboTipoJustificativa, xmlMapaNavegacao, "TipoJustificativa", "TP_JUST_CNCL", "NO_TIPO_JUST_CNCL")
    
    Set objMIU = Nothing
    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarFormulario", 0

End Sub

'Formata as colunas da lista de mensagens
Private Sub flInicializarLvwMensagem()

On Error GoTo ErrorHandler
    
    With Me.lvwMensagem.ColumnHeaders
        .Clear
        
        If PerfilAcesso = BackOffice Then
            .Add , , "Local Liquida��o", 1455
            .Add , , "ISPB IF", 1050
            .Add , , "CNPJ Contraparte", 1600
            .Add , , "Nome Contraparte", 3670
            .Add , , "Contraparte", 1000
            .Add , , "Ag�ncia", 860
            .Add , , "Conta Corrente", 1560
            .Add , , "Valor Apurado SLCC", 1800, lvwColumnRight
            .Add , , "Status", 2000
            .Add , , "Id Lote", 1000
            .Add , , "Sigla de Sistema", 1500
            .Add , , "Data Lote", 1600
            .Add , , "D/C Lote", 1000
            .Add , , "Valor Total Lote", 2000
            .Add , , "Quantidade Opera��es Lote", 3000
            'RATS 1211 - Inclus�o de Campos na Concilia��o
            .Add , , "Quantidade Apurada SLCC", 3000
        Else
            If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                .Add , , "Local Liquida��o", 1455
                .Add , , "ISPB IF", 1050
                .Add , , "CNPJ Contraparte", 1600
                .Add , , "Nome Contraparte", 5670
                .Add , , "Contraparte", 1000
                .Add , , "Ag�ncia", 860
                .Add , , "Conta Corrente", 1560
                .Add , , "Valor Apurado SLCC", 1800, lvwColumnRight
                .Add , , "Id Lote", 1000
                .Add , , "Sigla de Sistema", 1500
                .Add , , "Data Lote", 1600
                .Add , , "D/C Lote", 1000
                .Add , , "Valor Total Lote", 2000
                .Add , , "Quantidade Opera��es Lote", 3000
                'RATS 1211 - Inclus�o de Campos na Concilia��o
                .Add , , "Quantidade Apurada SLCC", 3000
                
            Else
                .Add , , "Local Liquida��o", 1455
                .Add , , "ISPB IF", 1050
                .Add , , "CNPJ Contraparte", 1600
                .Add , , "Nome Contraparte", 2470
                .Add , , "Contraparte", 1000
                .Add , , "Ag�ncia", 860
                .Add , , "Conta Corrente", 1560
                .Add , , "Valor Contraparte", 1600, lvwColumnRight
                .Add , , "Valor Opera��o Apurado SLCC", 2600, lvwColumnRight
                .Add , , "Diferen�a", 1600, lvwColumnRight
                .Add , , "Id Lote", 1000
                .Add , , "Sigla de Sistema", 1500
                .Add , , "Data Lote", 1600
                .Add , , "D/C Lote", 1000
                .Add , , "Valor Total Lote", 2000
                .Add , , "Quantidade Opera��es Lote", 3000
                'RATS 1211 - Inclus�o de Campos na Concilia��o
                .Add , , "Quantidade Apurada SLCC", 3000
            End If
        End If
    End With
    
Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwMensagem", 0

End Sub

'Formata as colunas da lista de opera��es
Private Sub flInicializarLvwOperacao()

On Error GoTo ErrorHandler
    
    With Me.lvwOperacao.ColumnHeaders
        .Clear
        .Add , , "Cliente", 2700
        .Add , , "ID Ativo", 2000
        .Add , , "Data Vencimento", 1600
        .Add , , "D/C", 800
        .Add , , "Valor", 1695, lvwColumnRight
        .Add , , "Quantidade", 1440, lvwColumnRight
        .Add , , "PU", 1243, lvwColumnRight
        .Add , , "Status", 2000
        .Add , , "N�mero Comando", 1700
        .Add , , "Data Liquida��o", 1600
        .Add , , "Data Opera��o", 1600
        .Add , , "C�digo", 2000
        .Add , , "Id Lote", 1000
        .Add , , "Sigla de Sistema", 1500
        .Add , , "Data Lote", 1600
        .Add , , "D/C Lote", 1000
        .Add , , "Valor Total Lote", 2000
        .Add , , "Quantidade Opera��es Lote", 3000
        'RATS 1211 - Inclus�o de Campos na Concilia��o
        '.Add , , "Quantidade Apurada SLCC", 3000
    End With
    
Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwOperacao", 0

End Sub

' Limpa o controle de justificativa
Private Sub flLimparControleJustificativa(Optional ByVal pblnVisaoPorQuantidade As Boolean = True)
    cboTipoJustificativa.ListIndex = -1
    txtComentario.Text = vbNullString
End Sub

'Apaga o conte�do das listas de mensagens e opera��es
Private Sub flLimparListas()
    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear
End Sub

'Exibe de forma diferenciada os itens que tenham sido rejeitados por motivo de grade de hor�rio
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim intCont                                 As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='3095' or CodigoErro='3023']")
            For Each objListItem In lvwMensagem.ListItems
                With objListItem
                    If Split(.Key, "|")(KEY_MSG_CO_LOCA_LIQU) = objDomNode.selectSingleNode("CO_LOCA_LIQU").Text And _
                       Split(.Key, "|")(KEY_MSG_CO_ISPB_CNPT) = objDomNode.selectSingleNode("CO_ISPB_CNPT").Text And _
                       Split(.Key, "|")(KEY_MSG_CO_CNPJ_CNPT) = objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text And _
                       Split(.Key, "|")(KEY_MSG_TP_IF_CRED_DEBT) = objDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text And _
                       Split(.Key, "|")(KEY_MSG_CO_AGEN_COTR) = objDomNode.selectSingleNode("CO_AGEN_COTR").Text And _
                       Split(.Key, "|")(KEY_MSG_NU_CC_COTR) = objDomNode.selectSingleNode("NU_CC_COTR").Text Then
                        
                        For intCont = 1 To .ListSubItems.Count
                            .ListSubItems(intCont).ForeColor = vbRed
                        Next
                        
                        .Text = "Hor�rio Excedido"
                        .ToolTipText = "Hor�rio limite p/envio da mensagem excedido"
                        .ForeColor = vbRed
                        
                        Exit For
                    
                    End If
                End With
            Next
        Next
    End If

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0

End Sub

'Monta o conte�do que ser� utilizado com a propriedade 'Key' dos itens do ListView
Private Function flMontarChaveItemListview(ByVal objDomNode As MSXML2.IXMLDOMNode)
                
Dim strListItemKey                          As String
    
On Error GoTo ErrorHandler
    
    strListItemKey = "|" & objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & _
                     "|" & objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text & _
                     "|" & objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text & _
                     "|" & objDomNode.selectSingleNode("CO_AGEN_COTR").Text & _
                     "|" & objDomNode.selectSingleNode("NU_CC_COTR").Text
                     

    If PerfilAcesso = BackOffice Then
        If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
            strListItemKey = strListItemKey & _
                             "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & _
                             "|" & objDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text
        Else
            strListItemKey = strListItemKey & _
                             "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & _
                             "|"
        End If
    Else
        If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
            strListItemKey = strListItemKey & _
                             "|" & _
                             "|" & objDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text
        Else
            strListItemKey = strListItemKey & _
                             "|" & _
                             "|"
        End If
    End If
    
    If Not objDomNode.selectSingleNode("CD_LOTE") Is Nothing Then
        strListItemKey = strListItemKey & _
                        "|" & objDomNode.selectSingleNode("CD_LOTE").Text
    Else
        strListItemKey = strListItemKey & "|"
    End If

    strListItemKey = strListItemKey & "|" & objDomNode.selectSingleNode("SG_SIST").Text
    
    If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> "00:00:00" Then
        strListItemKey = strListItemKey & "|" & objDomNode.selectSingleNode("DT_OPER_ATIV").Text
    Else
        strListItemKey = strListItemKey & "|"
    End If
    
    If PerfilAcesso = AdmArea Then
        If Not objDomNode.selectSingleNode("NU_CTRL_IF") Is Nothing Then
            strListItemKey = strListItemKey & "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text
        Else
            strListItemKey = strListItemKey & "|"
        End If
    Else
        strListItemKey = strListItemKey & "|"
    End If
    
    flMontarChaveItemListview = strListItemKey
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Monta uma express�o XPath para sele��o do conte�do de um documento XML
Private Function flMontarCondicaoNavegacaoXMLOperacoes(ByVal strItemKey As String)
                
Dim strCondicao                             As String
    
On Error GoTo ErrorHandler
    
    
    strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_MSG_CO_LOCA_LIQU) & "' " & _
                                                          " and CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_ISPB_CNPT) & "' " & _
                                                          " and CO_CNPJ_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_CNPJ_CNPT) & "' " & _
                                                          " and CO_AGEN_COTR='" & Split(strItemKey, "|")(KEY_MSG_CO_AGEN_COTR) & "' " & _
                                                          " and NU_CC_COTR='" & Split(strItemKey, "|")(KEY_MSG_NU_CC_COTR) & "' " & _
                                                          " and CD_LOTE='" & Split(strItemKey, "|")(KEY_MSG_ID_LOTE) & "' " & _
                                                          " and SG_SIST='" & Split(strItemKey, "|")(KEY_MSG_SG_SIST) & "' " & _
                                                          " and DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DATA_OPERACAO) & "' "
    
    
    If PerfilAcesso = BackOffice Then
        If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
            strCondicao = strCondicao & _
                        " and CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' " & _
                        " and TP_IF_CRED_DEBT='" & Split(strItemKey, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "' "
        Else
            strCondicao = strCondicao & _
                        " and CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' "
        End If
    Else
        If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
            strCondicao = strCondicao & _
                        " and TP_IF_CRED_DEBT='" & Split(strItemKey, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "' "
        End If
    End If
    
    strCondicao = strCondicao & "]"
    flMontarCondicaoNavegacaoXMLOperacoes = strCondicao
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoes", 0

End Function

'Monta uma express�o XPath para a somat�ria dos valores de opera��es
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String)
                
Dim strDebito                               As String
Dim strCredito                              As String
    
On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_MSG_CO_LOCA_LIQU) & "' " & _
                                     " and ../CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_ISPB_CNPT) & "' " & _
                                     " and ../CO_CNPJ_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_CNPJ_CNPT) & "' " & _
                                     " and ../CO_AGEN_COTR='" & Split(strItemKey, "|")(KEY_MSG_CO_AGEN_COTR) & "' " & _
                                     " and ../NU_CC_COTR='" & Split(strItemKey, "|")(KEY_MSG_NU_CC_COTR) & "' " & _
                                     " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_MSG_ID_LOTE) & "' " & _
                                     " and ../SG_SIST='" & Split(strItemKey, "|")(KEY_MSG_SG_SIST) & "' " & _
                                     " and ../DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DATA_OPERACAO) & "' "
    
    strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                      " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_MSG_CO_LOCA_LIQU) & "' " & _
                                      " and ../CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_ISPB_CNPT) & "' " & _
                                      " and ../CO_CNPJ_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_CNPJ_CNPT) & "' " & _
                                      " and ../CO_AGEN_COTR='" & Split(strItemKey, "|")(KEY_MSG_CO_AGEN_COTR) & "' " & _
                                      " and ../NU_CC_COTR='" & Split(strItemKey, "|")(KEY_MSG_NU_CC_COTR) & "' " & _
                                      " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_MSG_ID_LOTE) & "' " & _
                                      " and ../SG_SIST='" & Split(strItemKey, "|")(KEY_MSG_SG_SIST) & "' " & _
                                      " and ../DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DATA_OPERACAO) & "' "
    
    If PerfilAcesso = BackOffice Then
        If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
            strDebito = strDebito & _
                        " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' " & _
                        " and ../TP_IF_CRED_DEBT='" & Split(strItemKey, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "' "
        
            strCredito = strCredito & _
                        " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' " & _
                        " and ../TP_IF_CRED_DEBT='" & Split(strItemKey, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "' "
        
        Else
            strDebito = strDebito & _
                        " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' "
        
            strCredito = strCredito & _
                        " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' "
        
        End If
    Else
        If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
            strDebito = strDebito & _
                        " and ../TP_IF_CRED_DEBT='" & Split(strItemKey, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "' "
        
            strCredito = strCredito & _
                        " and ../TP_IF_CRED_DEBT='" & Split(strItemKey, "|")(KEY_MSG_TP_IF_CRED_DEBT) & "' "
        
        End If
    End If

    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "
    
    flMontarExpressaoCalculoNetOperacoes = strDebito & strCredito
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoCalculoNetOperacoes", 0

End Function

'Monta uma express�o XPath para a somat�ria dos valores de opera��es
Private Function flMontarExpressaoNetQtdeOperLote(ByVal strItemKey As String)
                
Dim strQtdeOperLote                               As String
    
On Error GoTo ErrorHandler
    
    strQtdeOperLote = "count(//CD_LOTE[../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_MSG_CO_LOCA_LIQU) & "' " & _
                                     " and ../CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_ISPB_CNPT) & "' " & _
                                     " and ../CO_CNPJ_CNPT='" & Split(strItemKey, "|")(KEY_MSG_CO_CNPJ_CNPT) & "' " & _
                                     " and ../CO_AGEN_COTR='" & Split(strItemKey, "|")(KEY_MSG_CO_AGEN_COTR) & "' " & _
                                     " and ../NU_CC_COTR='" & Split(strItemKey, "|")(KEY_MSG_NU_CC_COTR) & "' " & _
                                     " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_MSG_ID_LOTE) & "' " & _
                                     " and ../SG_SIST='" & Split(strItemKey, "|")(KEY_MSG_SG_SIST) & "' " & _
                                     " and ../DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DATA_OPERACAO) & "']) "
    
    flMontarExpressaoNetQtdeOperLote = strQtdeOperLote
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoNetQtdeOperLote", 0

End Function


'Monta o XML com os dados de filtro para sele��o de opera��es
Private Function flMontarXMLFiltroPesquisa(ByVal intTipoPesquisa As enumTipoPesquisa) As String
    
Dim xmlFiltros                              As MSXML2.DOMDocument40
    
On Error GoTo ErrorHandler
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux)) & "000000"))
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux)) & "235959"))
    
    If intTipoPesquisa = operacao Then
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        If PerfilAcesso = BackOffice Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ManualEmSer)
        Else
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ManualEmSer)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeAutomatico)
        End If
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
        Call fgAppendNode(xmlFiltros, "Grupo_LayoutEntrada", "LayoutEntrada", "50")
            
    Else
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "STR0008R2")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "STR0006R2")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "STR0007R2")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "PAG0108R2")
        
    End If
    
    flMontarXMLFiltroPesquisa = xmlFiltros.xml
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Monta XML com as chaves das opera��es que ser�o processadas
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer
Dim strItemKey                              As String
    
On Error GoTo ErrorHandler
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")
    
    For Each objListItem In Me.lvwMensagem.ListItems
        With objListItem
            If .Checked Then
                
                Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_EMPR", _
                                                   fgObterCodigoCombo(cboEmpresa.Text))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_LOCA_LIQU", _
                                                   Split(.Key, "|")(KEY_MSG_CO_LOCA_LIQU))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_ISPB_CNPT", _
                                                   Split(.Key, "|")(KEY_MSG_CO_ISPB_CNPT))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_CNPJ_CNPT", _
                                                   Split(.Key, "|")(KEY_MSG_CO_CNPJ_CNPT))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "NO_CNPT", _
                                                   .SubItems(COL_MSG_NOME_CORR))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "TP_IF_CRED_DEBT", _
                                                   Split(.Key, "|")(KEY_MSG_TP_IF_CRED_DEBT))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_AGEN_COTR", _
                                                   Split(.Key, "|")(KEY_MSG_CO_AGEN_COTR))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "NU_CC_COTR", _
                                                   Split(.Key, "|")(KEY_MSG_NU_CC_COTR))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "ID_PGTO_RECB_GRUP", _
                                                   IIf(optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value, "P", "R"))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "TipoJustificativa", _
                                                   fgObterCodigoCombo(Me.cboTipoJustificativa))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "Comentario", _
                                                   Trim$(txtComentario.Text))
                
                If PerfilAcesso = AdmArea And optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value Then
                    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                       "VA_LANC", _
                                                       Replace(fgVlr_To_Xml(.SubItems(COL_MSG_VALR_OPER)), "-", vbNullString))
                Else
                    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                       "VA_LANC", _
                                                       Replace(fgVlr_To_Xml(.SubItems(COL_MSG_VALR)), "-", vbNullString))
                End If
                
                intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)
                
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "IgnoraGradeHorario", _
                                                   intIgnoraGradeHorario)
            
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
                End If
                
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "Repeat_Operacao", _
                                                   "")
                
                strItemKey = objListItem.Key
                
                For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))
                    Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
                    Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "NU_SEQU_OPER_ATIV", _
                                                       objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "DH_ULTI_ATLZ", _
                                                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "CO_ULTI_SITU_PROC", _
                                                       objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text)
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "ID_PGTO_RECB", _
                                                       IIf(optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value, "P", "R"))
                    
                    Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)
                    
                    Set xmlItemOperacao = Nothing
                Next
            
                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
                Set xmlItemEnvioMsg = Nothing
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

'Exibe o resultado da �ltima opera��o executada
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " liquidados "
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
    Call flLimparControleJustificativa
    
    If Me.cboEmpresa.ListIndex = -1 Or Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    fgCursor True
    
    strDocFiltros = flMontarXMLFiltroPesquisa(operacao)
    Call flCarregarListaNetOperacoes(strDocFiltros)
    
    If PerfilAcesso = AdmArea And optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value Then
        strDocFiltros = flMontarXMLFiltroPesquisa(MENSAGEM)
        Call flCarregarListaNetMensagens(strDocFiltros)
        Call flCalcularDiferencasListView
    End If
    
    If lvwMensagem.ListItems.Count > 0 Then
        lvwMensagem.ListItems(1).Selected = True
        Call lvwMensagem_ItemClick(lvwMensagem.ListItems(1))
    End If

    Call flConfigurarBotoesPorPerfil(PerfilAcesso)

    fgCursor
    
Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPesquisar", 0)

End Sub

'Enviar itens de mensagem e opera��es para liquida��o
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
        strXMLRetorno = objOperacaoMensagem.LiquidarCorretoras(intAcaoProcessamento, _
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
Private Function flAlterarStatusLotes(ByVal pEnumSituacaoLote As enumSituacaoLote)
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim objConciliacaoLote                      As A8MIU.clsOperacaoMensagem
Dim lngIdLote                               As Long
Dim strSiglaSistema                         As String
Dim dtDataLote                              As Double
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim datDataLote                             As Date
Dim strDataLote                             As String

On Error GoTo ErrorHandler

Set objConciliacaoLote = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
    For Each objListItem In Me.lvwMensagem.ListItems
        If objListItem.Checked Then
            If PerfilAcesso = BackOffice Then
                lngIdLote = CLng(objListItem.SubItems(COL_MSG_ID_LOTE))
                strSiglaSistema = CStr(objListItem.SubItems(COL_MSG_SG_SIST))
                datDataLote = CDate(objListItem.SubItems(COL_MSG_DATA_OPERACAO))
            Else 'Administrador da �rea
                If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                    lngIdLote = CLng(objListItem.SubItems(COL_MSG_ID_LOTE_PGTO))
                    strSiglaSistema = CStr(objListItem.SubItems(COL_MSG_SG_SIST_PGTO))
                    datDataLote = CDate(objListItem.SubItems(COL_MSG_DATA_OPERACAO_PGTO))
                Else
                    lngIdLote = CLng(objListItem.SubItems(COL_MSG_ID_LOTE_RECB))
                    strSiglaSistema = CStr(objListItem.SubItems(COL_MSG_SG_SIST_RECB))
                    datDataLote = CDate(objListItem.SubItems(COL_MSG_DATA_OPERACAO_RECB))
                End If
            End If
            strDataLote = CStr(Format(datDataLote, "DD/MM/YYYY"))
            objConciliacaoLote.AlterarStatusLote lngIdLote, strSiglaSistema, strDataLote, pEnumSituacaoLote, vntCodErro, vntMensagemErro
        End If
    Next
    
Exit Function
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flLiberarLotes", Me.Caption

End Function
Private Function flValidarNetComLote() As Boolean
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem

Dim intTotalOper                            As Integer
Dim intTotalOperLote                        As Integer
Dim dblValorNet                             As Double
Dim dblValorTotalLote                       As Double




    For Each objListItem In Me.lvwMensagem.ListItems
        If objListItem.Checked Then
            
            intTotalOper = 0
            For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItem.Key))
                intTotalOper = intTotalOper + 1
            Next
            
            'Validar Itens de Acordo com o Lote
            If PerfilAcesso = enumPerfilAcesso.AdmArea Then
                If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                    If objListItem.SubItems(COL_MSG_VALR_CAMR) <> "" Then
                        dblValorNet = CDbl(objListItem.SubItems(COL_MSG_VALR_CAMR))
                    End If
                Else
                    If objListItem.SubItems(COL_MSG_VALR_OPER) <> "" Then
                        dblValorNet = CDbl(objListItem.SubItems(COL_MSG_VALR_OPER))
                    End If
                End If
            Else
                If objListItem.SubItems(COL_MSG_VALR_CAMR) <> "" Then
                    dblValorNet = CDbl(objListItem.SubItems(COL_MSG_VALR_CAMR))
                End If
            End If
            
            If PerfilAcesso = BackOffice Then
                dblValorTotalLote = CDbl(objListItem.SubItems(COL_MSG_VA_TOT_LOTE))
                intTotalOperLote = CInt(objListItem.SubItems(COL_MSG_QT_OPER_LOTE))
            Else
                If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                    dblValorTotalLote = CDbl(objListItem.SubItems(COL_MSG_VA_TOT_LOTE_PGTO))
                    intTotalOperLote = CInt(objListItem.SubItems(COL_MSG_QT_OPER_LOTE_PGTO))
                Else
                    If objListItem.SubItems(COL_MSG_VA_TOT_LOTE_RECB) <> "" Then
                        dblValorTotalLote = CDbl(objListItem.SubItems(COL_MSG_VA_TOT_LOTE_RECB))
                    End If
                    If objListItem.SubItems(COL_MSG_QT_OPER_LOTE_RECB) <> "" Then
                        intTotalOperLote = CInt(objListItem.SubItems(COL_MSG_QT_OPER_LOTE_RECB))
                    End If
                End If
            End If
            
            If ((dblValorTotalLote - dblValorNet) <> 0) Or (dblValorTotalLote = 0) Then
                If dblValorTotalLote = 0 Then
                    MsgBox "Valor do Lote n�o apurado pelo SLCC. N�o � poss�vel efetuar a Libera��o.", vbExclamation + vbOKOnly, Me.Caption
                    Exit Function
                ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
                    If MsgBox("Valor do Lote diferente do Valor apurado pelo SLCC. Deseja liberar o lote assim mesmo?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                        Exit Function
                    Else
                        If cboTipoJustificativa.Text = Empty Or txtComentario.Text = Empty Then
                            MsgBox "� necess�rio informar a Justificativa e o Coment�rio.", vbExclamation + vbOKOnly, Me.Caption
                            Exit Function
                        Else
                            flValidarNetComLote = True
                            Exit Function
                        End If
                    End If
                ElseIf PerfilAcesso = enumPerfilAcesso.BackOffice Then
                    If MsgBox("Valor do Lote diferente do Valor apurado pelo SLCC. Deseja liberar o lote pelo valor apurado pelo SLCC?", vbExclamation + vbYesNo, Me.Caption) = vbNo Then
                        Exit Function
                    Else
                        flValidarNetComLote = True
                        Exit Function
                    End If
                ElseIf PerfilAcesso = enumPerfilAcesso.AdmArea Then
                    MsgBox "Valor do Lote diferente do Valor apurado pelo SLCC. O lote s� poder� ser liberado pela Conting�ncia.", vbExclamation + vbOKOnly, Me.Caption
                    Exit Function
                End If
            End If
                
            If intTotalOperLote <> intTotalOper Then
            
                If intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
                    If MsgBox("Quantidade de Opera��es do Lote diferente da quantidade apurada pelo SLCC. Deseja liberar o lote assim mesmo?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                        Exit Function
                    Else
                        If cboTipoJustificativa.Text = vbNullString Or txtComentario.Text = vbNullString Then
                            MsgBox "� necess�rio informar a Justificativa e o Coment�rio", vbExclamation + vbOKOnly, Me.Caption
                            Exit Function
                        Else
                            flValidarNetComLote = True
                        End If
                    End If
                ElseIf PerfilAcesso = enumPerfilAcesso.BackOffice Then
                    If MsgBox("Quantidade de Opera��es do Lote diferente da quantidade de Opera��es apuradas pelo SLCC. Deseja liberar o lote pelo valor apurado pelo SLCC?", vbExclamation + vbYesNo, Me.Caption) = vbNo Then
                        Exit Function
                    End If
                ElseIf PerfilAcesso = enumPerfilAcesso.AdmArea Then
                    MsgBox "Quantidade de Opera��es do Lote diferente da quantidade de Opera��es apuradas pelo SLCC. O lote s� poder� ser liberado por Conting�ncia", vbExclamation + vbOKOnly, Me.Caption
                    Exit Function
                End If
            End If
        End If
    Next
    
    flValidarNetComLote = True

End Function
'Valida a sele��o dos itens na tela, para posterior processamento
Private Function flValidarItensProcessamento(ByVal intAcao As enumAcaoConciliacao) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem


    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flValidarItensProcessamento = "Selecione pelo menos um item da lista, antes de prosseguir com a opera��o desejada."
        Exit Function
    End If
    
    If intAcao = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
        If Trim$(cboTipoJustificativa.Text) = vbNullString Then
            flValidarItensProcessamento = "Selecione o Tipo de Justificativa."
            If cboTipoJustificativa.Enabled Then cboTipoJustificativa.SetFocus
            Exit Function
        End If
    
        If Trim$(txtComentario.Text) = vbNullString Then
            flValidarItensProcessamento = "Informe o motivo da libera��o por conting�ncia."
            If txtComentario.Enabled Then txtComentario.SetFocus
            Exit Function
        End If
    End If
    
    For Each objListItem In Me.lvwMensagem.ListItems
        If objListItem.Checked Then
            If PerfilAcesso = AdmArea Then
                If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                    If intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar Then
                        
                        For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItem.Key))
                            If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.ConcordanciaBackoffice Then
                                
                                flValidarItensProcessamento = "Existem uma ou mais opera��es com Status diferente de 'Concord�ncia Backoffice'. Libera��o n�o permitida."
                                Exit Function
                                
                            End If
                        Next
                    
                    End If
                    
                Else
                    If intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar Then
                        
                        If Trim$(objListItem.SubItems(COL_MSG_VALR_OPER)) = vbNullString And _
                           objListItem.SubItems(COL_MSG_CNPT) = "Externa" Then
                            'flValidarItensProcessamento = "Valor de mensagem n�o encontrado em um ou mais itens selecionados para processamento. Libera��o n�o permitida."
                            flValidarItensProcessamento = "Valor do Lote n�o apurado pelo SLCC. N�o � poss�vel efetuar a Libera��o."
                            Exit Function
                            
                        ElseIf objListItem.ListSubItems(COL_MSG_DIFE).ForeColor = vbRed And _
                               objListItem.SubItems(COL_MSG_CNPT) = "Externa" Then
                            flValidarItensProcessamento = "Libera��o n�o permitida: um ou mais item n�o batido. Se necess�rio, utilize o processo Liberar em Conting�ncia"
                            Exit Function
                            
                        End If
                        
                        For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItem.Key))
                            If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.ConcordanciaBackoffice And Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.ConcordanciaBackofficeAutomatico Then
                                
                                flValidarItensProcessamento = "Existem uma ou mais opera��es com Status diferente de 'Concord�ncia Backoffice'. Libera��o n�o permitida."
                                Exit Function
                                
                            End If
                        Next
                    
                    ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
                        
                        If Trim$(objListItem.SubItems(COL_MSG_VALR_OPER)) = vbNullString Then
                            'flValidarItensProcessamento = "Valor de mensagem encontrado em um ou mais itens selecionados para processamento. Libera��o por conting�ncia n�o permitida."
                            flValidarItensProcessamento = "Valor do Lote n�o apurado pelo SLCC. N�o � poss�vel efetuar a Libera��o por Conting�ncia."
                            Exit Function
                        End If
                        
'                        If objListItem.SubItems(COL_MSG_CNPT) = "Interna" Then
'                            flValidarItensProcessamento = "Libera��o por conting�ncia s� � permitida para Contraparte Externa."
'                            Exit Function
'                        End If
                        
                        For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItem.Key))
                            If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.ConcordanciaBackoffice Then
                                
                                flValidarItensProcessamento = "Existem uma ou mais opera��es com Status diferente de 'Concord�ncia Backoffice'. Libera��o por conting�ncia n�o permitida."
                                Exit Function
                                
                            End If
                        Next
                    End If
                End If
            End If
        End If
    Next
    
End Function

'Calcula o valor da opera��es
Private Function flValorOperacoes(ByVal strItemKey As String)
    
Dim strExpression                   As String
Dim vntValor                        As Variant
    
    vntValor = 0
    
    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlOperacoes, strExpression))
    
    flValorOperacoes = vntValor

End Function

'Calcula a quantidade de opera��es do lote
Private Function flQtdeOperLote(ByVal strItemKey As String)
    
Dim strExpression                   As String
Dim vntValor                        As Variant
    
    vntValor = 0
    
    strExpression = flMontarExpressaoNetQtdeOperLote(strItemKey)
    vntValor = flQtdeOperLote + Val(fgFuncaoXPath(xmlOperacoes, strExpression))
    
    flQtdeOperLote = vntValor

End Function


'Configura o perfil de acesso do usu�rio
Property Get PerfilAcesso() As enumPerfilAcesso
    PerfilAcesso = lngPerfil
End Property

'Configura o perfil de acesso do usu�rio
Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
    
    lngPerfil = pPerfil
    
    Select Case pPerfil
        Case enumPerfilAcesso.BackOffice
            Me.Caption = "Libera��o - Liquida��o com Corretoras (Backoffice)"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "Libera��o - Liquida��o com Corretoras (Administrador de �rea)"
    End Select
    
    Call flConfigurarBotoesPorPerfil(PerfilAcesso)
    Call flLimparListas
    Call flLimparControleJustificativa
    
    If cboEmpresa.ListIndex <> -1 Or cboEmpresa.Text <> vbNullString Then
        Call flInicializarLvwMensagem
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
    Call flInicializarFormulario
    Call flInicializarLvwMensagem
    Call flInicializarLvwOperacao
    fgCursor
    
    Set xmlOperacoes = CreateObject("MSXML2.DOMDocument.4.0")
    
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

        .lvwMensagem.Top = .fraNaturezaMovto.Top + .fraNaturezaMovto.Height + 120
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .imgDummyH.Top = .lvwMensagem.Top + .lvwMensagem.Height
        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 780 - .fraControles.Height
        .lvwOperacao.Width = .Width - 240
    
        .fraControles.Top = .lvwOperacao.Top + .lvwOperacao.Height
        .fraControles.Left = .cboEmpresa.Left
        .fraControles.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set frmLiquidacaoCorretoras = Nothing
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
        If .imgDummyH.Top > (.Height - 3000) And (.Height - 3000) > 0 Then
            .imgDummyH.Top = .Height - 3000
        End If

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 780 - .fraControles.Height
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

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call flCarregarListaDetalheOperacoes
    
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
            .SequenciaOperacao = Split(lvwOperacao.SelectedItem.Key, "|")(KEY_OP_NU_SEQU_OPER_ATIV)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

End Sub

Private Sub optNaturezaMovimento_Click(Index As Integer)

On Error GoTo ErrorHandler
    
    Call flLimparListas
    Call flLimparControleJustificativa
    DoEvents
    
    tlbComandos.Buttons("liberacaocontingencia").Enabled = IIf(PerfilAcesso = enumPerfilAcesso.AdmArea, True, False)
                                                               
    cboTipoJustificativa.Enabled = tlbComandos.Buttons("liberacaocontingencia").Enabled
    txtComentario.Enabled = tlbComandos.Buttons("liberacaocontingencia").Enabled
    
 
    
    Call flInicializarLvwMensagem
    Call flPesquisar
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optNaturezaMovimento_Click", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
Dim strValidaProcessamento                  As String
Dim strValidaNetComLote                     As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    intAcaoProcessamento = 0
    
    Select Case Button.Key
        Case "refresh"
            Call flPesquisar
            
        Case "concordancia"
            intAcaoProcessamento = enumAcaoConciliacao.BOConcordar
            
        Case "liberacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar
            
        Case "liberacaocontingencia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberarContingencia
            
        Case "retorno"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaRejeitar
            
        Case gstrSair
            Unload Me
            
    End Select
    
    If intAcaoProcessamento <> 0 Then
        strValidaProcessamento = flValidarItensProcessamento(intAcaoProcessamento)
        If strValidaProcessamento <> vbNullString Then
            frmMural.Display = strValidaProcessamento
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            GoTo ExitSub
        End If
        
        'Validar Net com o Lote
        If intAcaoProcessamento <> enumAcaoConciliacao.AdmAreaRejeitar Then
            If Not flValidarNetComLote() Then GoTo ExitSub
        End If
        
        If intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar Then
            If MsgBox("Confirma a libera��o do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
            If MsgBox("Confirma a libera��o por conting�ncia do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        End If
        
        strResultadoOperacao = flProcessar
        
        If strResultadoOperacao <> vbNullString Then
            If Button.Key = "liberacao" Or Button.Key = "liberacaocontingencia" Then
                Call flAlterarStatusLotes(enumSituacaoLote.Liberado)
            ElseIf Button.Key = "retorno" Then
                Call flAlterarStatusLotes(enumSituacaoLote.Rejeitado)
            End If
            Call flMostrarResultado(strResultadoOperacao)
            Call flPesquisar
        End If
    
        Call flMarcarRejeitadosPorGradeHorario
    End If
    
ExitSub:
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
