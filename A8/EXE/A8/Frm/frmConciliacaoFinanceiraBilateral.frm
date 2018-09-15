VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConciliacaoFinanceiraBilateral 
   Caption         =   "CBLC - Liquidação Bruta"
   ClientHeight    =   8640
   ClientLeft      =   975
   ClientTop       =   855
   ClientWidth     =   12975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   12975
   WindowState     =   2  'Maximized
   Begin VB.Frame fraComplementos 
      Height          =   765
      Left            =   60
      TabIndex        =   4
      Top             =   7500
      Width           =   12855
      Begin VB.TextBox txtTotalMensagem 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         Width           =   2235
      End
      Begin VB.TextBox txtTotalOperacao 
         Alignment       =   1  'Right Justify
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
         Left            =   5955
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   2235
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Mensagens"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Operações"
         Height          =   195
         Index           =   1
         Left            =   4695
         TabIndex        =   7
         Top             =   360
         Width           =   1185
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
      Left            =   0
      Top             =   7050
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
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoFinanceiraBilateral.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   3255
      Left            =   60
      TabIndex        =   1
      Top             =   4080
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
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3285
      Left            =   60
      TabIndex        =   2
      Top             =   690
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5794
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
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   8310
      Width           =   12975
      _ExtentX        =   22886
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
Attribute VB_Name = "frmConciliacaoFinanceiraBilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Liquidação Bilateral de operações CETIP

Option Explicit

Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_M_NO_VEIC_LEGA            As Integer = 0
Private Const COL_M_CO_PARP_CAMR            As Integer = 1
Private Const COL_M_CO_ISPB_PART_CAMR       As Integer = 2
Private Const COL_M_CO_ISPB_BANC_LIQU_CNPT  As Integer = 3
Private Const COL_M_VA_FINC                 As Integer = 4
Private Const COL_M_TP_ACAO_MESG_SPB_EXEC   As Integer = 5
Private Const COL_M_NU_CTRL_CAMR            As Integer = 6
Private Const COL_M_DE_SITU_PROC            As Integer = 7

'Constantes de Configuração de Colunas de Operação
Private Const COL_O_NO_VEIC_LEGA            As Integer = 0
Private Const COL_O_CO_PARP_CAMR            As Integer = 1
Private Const COL_O_NO_CNPT                 As Integer = 2
Private Const COL_O_VA_OPER_ATIV            As Integer = 3
Private Const COL_O_CO_CNPT_CAMR            As Integer = 4
Private Const COL_O_DE_SITU_PROC            As Integer = 5
Private Const COL_O_NR_CNPJ_CPF_COMITENTE   As Integer = 6


'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const KEY_MSG_NU_CTRL_IF            As Integer = 1
Private Const KEY_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const KEY_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const KEY_MSG_CO_ULTI_SITU_PROC     As Integer = 4

'Constantes de posicionamento de campos na propriedade Tag do item do ListView de Mensagens
Private Const TAG_MSG_NU_SEQU_CNCL          As Integer = 1

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const KEY_OP_CO_PARP_CAMR           As Integer = 1
Private Const KEY_OP_CO_ISPB_BANC_LIQU_CNPT As Integer = 2
Private Const KEY_OP_CO_CNPT_CAMR           As Integer = 3
Private Const KEY_OP_CO_ULTI_SITU_PROC      As Integer = 4

'Constantes de posicionamento de campos na propriedade Tag do item do ListView de Operações
Private Const TAG_OP_NU_SEQU_CNCL           As Integer = 1
'------------------------------------------------------------------------------------------
'Fim declaração constantes

Private lngPerfil                           As Long
Private intAcaoProcessamento                As enumAcaoConciliacao
Private blnDummyH                           As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Calcular totais de mensagens e nets de operações
Private Sub flCalcularTotais()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

    On Error GoTo ErrorHandler

    dblValorMensagem = 0
    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            If .Checked Then
                dblValorMensagem = dblValorMensagem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_M_VA_FINC)))
            End If
        End With
    Next

    txtTotalMensagem.Text = fgVlrXml_To_Interface(dblValorMensagem)

    dblValorOperacao = 0
    For Each objListItem In lvwOperacao.ListItems
        With objListItem
            If .Checked Then
                dblValorOperacao = dblValorOperacao + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_O_VA_OPER_ATIV)))
            End If
        End With
    Next

    txtTotalOperacao.Text = fgVlrXml_To_Interface(dblValorOperacao)

    If lvwMensagem.ListItems.Count > 0 Or lvwOperacao.ListItems.Count > 0 Then
        If dblValorMensagem - dblValorOperacao = 0 Then
            txtTotalMensagem.ForeColor = vbBlack
            txtTotalOperacao.ForeColor = vbBlack
            
            With tlbComandos
                .Buttons("concordancia").Enabled = True
                .Buttons("discordancia").Enabled = False
                .Buttons("pagamento").Enabled = True
                .Buttons("pagamentocontingencia").Enabled = False
                .Buttons("regularizacao").Enabled = True
                .Buttons("retorno").Enabled = True
                .Buttons("pagamentostr").Enabled = True
                .Buttons("pagamentobacen").Enabled = True
            End With
        Else
            txtTotalMensagem.ForeColor = vbRed
            txtTotalOperacao.ForeColor = vbRed
            
            With tlbComandos
                .Buttons("concordancia").Enabled = False
                .Buttons("discordancia").Enabled = True
                .Buttons("pagamento").Enabled = False
                .Buttons("pagamentocontingencia").Enabled = True
                .Buttons("regularizacao").Enabled = False
                .Buttons("retorno").Enabled = True
                .Buttons("pagamentostr").Enabled = False
                .Buttons("pagamentobacen").Enabled = False
            End With
        End If
    Else
        txtTotalMensagem.ForeColor = vbBlack
        txtTotalOperacao.ForeColor = vbBlack
    End If
    
    Exit Sub

ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularTotais", 0)

End Sub

'Carregar lista de mensagens
Private Sub flCarregarListaMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objSubItem                              As MSComctlLib.ListSubItem
Dim strListItemKey                          As String

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

    On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

    Call xmlRetLeitura.loadXML(objMensagem.ObterDetalheMensagem(pstrFiltro, vntCodErro, vntMensagemErro))

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If xmlRetLeitura.xml <> vbNullString Then
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text

            With lvwMensagem.ListItems.Add(, strListItemKey)

                .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text

                .SubItems(COL_M_CO_ISPB_PART_CAMR) = objDomNode.selectSingleNode("CO_ISPB_PART_CAMR").Text
                .SubItems(COL_M_CO_ISPB_BANC_LIQU_CNPT) = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
                .SubItems(COL_M_CO_PARP_CAMR) = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                .SubItems(COL_M_VA_FINC) = "-" & fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                .SubItems(COL_M_DE_SITU_PROC) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                .SubItems(COL_M_TP_ACAO_MESG_SPB_EXEC) = fgDescricaoTipoAcao(Val(objDomNode.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text))
                .SubItems(COL_M_NU_CTRL_CAMR) = objDomNode.selectSingleNode("NU_CTRL_CAMR").Text

                If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.ConcordanciaBackoffice Or _
                   Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.PagamentoBackoffice Then
                    .ForeColor = vbBlue
                    For Each objSubItem In .ListSubItems
                        objSubItem.ForeColor = vbBlue
                    Next
                ElseIf Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.DiscordanciaBackoffice Then
                    .ForeColor = vbRed
                    For Each objSubItem In .ListSubItems
                        objSubItem.ForeColor = vbRed
                    Next
                End If
                
                If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.AConciliar Then
                    .Tag = "|0"
                Else
                    .Tag = "|" & objDomNode.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text
                End If
                
            End With
        Next
    End If

    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)

    Set xmlRetLeitura = Nothing
    Set objMensagem = Nothing

    Exit Sub

ErrorHandler:
    Set xmlRetLeitura = Nothing
    Set objMensagem = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaMensagens", 0)

End Sub

'Carregar lista de operações
Private Sub flCarregarListaOperacoes(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim dblValorOperacao                        As Double
Dim objSubItem                              As MSComctlLib.ListSubItem

    On Error GoTo ErrorHandler

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

    Call xmlRetLeitura.loadXML(objOperacao.ObterDetalheOperacao(pstrFiltro, vntCodErro, vntMensagemErro))

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Call xmlOperacoes.loadXML(xmlRetLeitura.xml)

    If xmlRetLeitura.xml <> vbNullString Then
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")

            strListItemKey = flMontarChaveItemListview(objDomNode)

            If Not fgExisteItemLvw(Me.lvwOperacao, strListItemKey) Then
                dblValorOperacao = flValorOperacoes(strListItemKey)

                If dblValorOperacao < 0 Then
                    With lvwOperacao.ListItems.Add(, strListItemKey)
        
                        .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
        
                        .SubItems(COL_O_CO_PARP_CAMR) = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                        .SubItems(COL_O_NO_CNPT) = objDomNode.selectSingleNode("NO_CNPT").Text
                        .SubItems(COL_O_VA_OPER_ATIV) = fgVlrXml_To_Interface(dblValorOperacao)
                        .SubItems(COL_O_CO_CNPT_CAMR) = objDomNode.selectSingleNode("CO_CNPT_CAMR").Text
                        .SubItems(COL_O_DE_SITU_PROC) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                        .SubItems(COL_O_NR_CNPJ_CPF_COMITENTE) = objDomNode.selectSingleNode("NR_CNPJ_CPJ").Text
                        
                        If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConcordanciaBackoffice Or _
                           Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.PagamentoBackoffice Then
                            .ForeColor = vbBlue
                            For Each objSubItem In .ListSubItems
                                objSubItem.ForeColor = vbBlue
                            Next
                        ElseIf Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.DiscordanciaBackoffice Then
                            .ForeColor = vbRed
                            For Each objSubItem In .ListSubItems
                                objSubItem.ForeColor = vbRed
                            Next
                        End If
                        
                        If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Registrada Or _
                           Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.RegistradaAutomatica Then
                            .Tag = "|0"
                        Else
                            .Tag = "|" & objDomNode.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text
                        End If
                        
                    End With
                End If
            End If
            
        Next
    End If

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)

    Set xmlRetLeitura = Nothing
    Set objOperacao = Nothing

    Exit Sub

ErrorHandler:
    Set xmlRetLeitura = Nothing
    Set objOperacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaOperacoes", 0)

End Sub

'Altera a exibição dos botões de acordo com o perfil do usuário
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)

    On Error GoTo ErrorHandler

    With tlbComandos
        .Buttons("concordancia").Visible = True
        .Buttons("discordancia").Visible = True
        .Buttons("pagamento").Visible = True
        .Buttons("pagamentocontingencia").Visible = True
        .Buttons("regularizacao").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("retorno").Visible = .Buttons("regularizacao").Visible
        .Buttons("pagamentostr").Visible = .Buttons("regularizacao").Visible
        .Buttons("pagamentobacen").Visible = .Buttons("regularizacao").Visible
        .Refresh
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorPerfil", 0

End Sub

'Inicializa controles de tela e variáveis
Private Sub flInicializarFormulario()

Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlLeitura, vbNullString, "Repeat_Filtro", vbNullString)
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsEmpresa", xmlLeitura))
    Call fgCarregarCombos(Me.cboEmpresa, xmlLeitura, "Empresa", "CO_EMPR", "NO_REDU_EMPR")

    Call flInicializarLvwMensagem
    Call flInicializarLvwOperacao
    Call flLimparListas
    Call flCalcularTotais
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarFormulario", 0

End Sub

'Formata as colunas da lista de mensagens
Private Sub flInicializarLvwMensagem()

    On Error GoTo ErrorHandler

    With Me.lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Veículo Legal Mensagem", 3000
        .Add , , "Identificador Part. Câmara", 2100
        .Add , , "ISPB Debitado", 1400
        .Add , , "ISPB Creditado", 1300
        .Add , , "Valor Mensagem", 1400, lvwColumnRight
        .Add , , "Ação", 2000
        .Add , , "Número Controle LTR", 1800
        .Add , , "Status", 1800
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwMensagem", 0

End Sub

'Formata as colunas da lista de operações
Private Sub flInicializarLvwOperacao()

    On Error GoTo ErrorHandler

    With Me.lvwOperacao.ColumnHeaders
        .Clear
        .Add , , "Veículo Legal Operação", 3000
        .Add , , "Identificador Part. Câmara", 2100
        .Add , , "Contraparte", 2700
        .Add , , "Net Operações", 1400, lvwColumnRight
        .Add , , "Identificador Part. Câmara Contraparte", 3800
        .Add , , "Status", 1800
        .Add , , "CNPJ/CPF Comitente", 2100
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwOperacao", 0

End Sub

'Apaga o conteúdo das listas de mensagens e operações
Private Sub flLimparListas()

    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear

End Sub

'Monta o conteúdo que será utilizado com a propriedade 'Key' dos itens do ListView
Private Function flMontarChaveItemListview(ByVal objDomNode As MSXML2.IXMLDOMNode)

Dim strListItemKey                          As String
Dim intStatus                               As Integer

    On Error GoTo ErrorHandler
    
    intStatus = Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text)
    If intStatus = enumStatusOperacao.RegistradaAutomatica Then
        intStatus = enumStatusOperacao.Registrada
    End If
    
    strListItemKey = "|" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text & _
                     "|" & objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text & _
                     "|" & objDomNode.selectSingleNode("CO_CNPT_CAMR").Text & _
                     "|" & intStatus

    flMontarChaveItemListview = strListItemKey

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Monta uma expressão XPath para seleção do conteúdo de um documento XML
Private Function flMontarCondicaoNavegacaoXMLOperacoes(ByVal strItemKey As String)

Dim strCondicao                             As String
Dim intStatus                               As Integer

    On Error GoTo ErrorHandler

    strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_OP_CO_PARP_CAMR) & "' " & _
                                                          " and CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_OP_CO_ISPB_BANC_LIQU_CNPT) & "' " & _
                                                          " and CO_CNPT_CAMR='" & Split(strItemKey, "|")(KEY_OP_CO_CNPT_CAMR) & "' "
    
    intStatus = Val(Split(strItemKey, "|")(KEY_OP_CO_ULTI_SITU_PROC))
    If intStatus = enumStatusOperacao.Registrada Then
        strCondicao = strCondicao & _
                      " and (CO_ULTI_SITU_PROC='" & enumStatusOperacao.Registrada & "' " & _
                      " or   CO_ULTI_SITU_PROC='" & enumStatusOperacao.RegistradaAutomatica & "') "
    Else
        strCondicao = strCondicao & _
                      " and  CO_ULTI_SITU_PROC='" & intStatus & "' "
    End If
    
    strCondicao = strCondicao & "]"

    flMontarCondicaoNavegacaoXMLOperacoes = strCondicao

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoes", 0

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String)

Dim strDebito                               As String
Dim strCredito                              As String
Dim intStatus                               As Integer

    On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_OP_CO_PARP_CAMR) & "' " & _
                                     " and ../CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_OP_CO_ISPB_BANC_LIQU_CNPT) & "' " & _
                                     " and ../CO_CNPT_CAMR='" & Split(strItemKey, "|")(KEY_OP_CO_CNPT_CAMR) & "' "

    strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                      " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_OP_CO_PARP_CAMR) & "' " & _
                                      " and ../CO_ISPB_BANC_LIQU_CNPT='" & Split(strItemKey, "|")(KEY_OP_CO_ISPB_BANC_LIQU_CNPT) & "' " & _
                                      " and ../CO_CNPT_CAMR='" & Split(strItemKey, "|")(KEY_OP_CO_CNPT_CAMR) & "' "

    intStatus = Val(Split(strItemKey, "|")(KEY_OP_CO_ULTI_SITU_PROC))
    If intStatus = enumStatusOperacao.Registrada Then
        strDebito = strDebito & _
                    " and (../CO_ULTI_SITU_PROC='" & enumStatusOperacao.Registrada & "' " & _
                    " or   ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.RegistradaAutomatica & "') "
    
        strCredito = strCredito & _
                     " and (../CO_ULTI_SITU_PROC='" & enumStatusOperacao.Registrada & "' " & _
                     " or   ../CO_ULTI_SITU_PROC='" & enumStatusOperacao.RegistradaAutomatica & "') "
    Else
        strDebito = strDebito & _
                    " and  ../CO_ULTI_SITU_PROC='" & intStatus & "' "
    
        strCredito = strCredito & _
                    " and  ../CO_ULTI_SITU_PROC='" & intStatus & "' "
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
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CETIP)

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")

    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.Registrada)
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica)
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    
    If PerfilAcesso = AdmArea Then
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.DiscordanciaBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.PagamentoBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.PagamentoBackofficeAutomatico)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.DiscordanciaBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.PagamentoBackoffice)
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.PagamentoBackofficeAutomatico)
    End If

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bilateral)

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoContraparte", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo)

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ISPBContraparte", "")
    Call fgAppendNode(xmlFiltros, "Grupo_ISPBContraparte", "ISPBContraparte", "")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CanalOperacaoInternaOperacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CanalOperacaoInternaOperacao", "CanalOperacaoInternaOperacao", "CONDIÇÃO (A.CO_CNAL_OPER_INTE = 'B' OR A.CO_CNAL_OPER_INTE IS NULL)")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CanalOperacaoInternaMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CanalOperacaoInternaMensagem", "CanalOperacaoInternaMensagem", "CONDIÇÃO (F.CO_CNAL_OPER_INTE = 'B' OR F.CO_CNAL_OPER_INTE IS NULL)")

    flMontarXMLFiltroPesquisa = xmlFiltros.xml

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Monta XML com as chaves das operações e mensagens a serem processadas
Private Function flMontarXMLProcessamento() As String

Dim objListItemMesg                         As ListItem
Dim objListItemOper                         As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvio                            As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim xmlDadosSTR                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strFinalidadeIF                         As String

    On Error GoTo ErrorHandler

    If intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoSTR Then
        strFinalidadeIF = flObterDominioFinalidadeMsgSTR
    End If
    
    Set xmlDadosSTR = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    For Each objListItemMesg In Me.lvwMensagem.ListItems
        With objListItemMesg
            If .Checked Then

                'Entrada Manual de dados complementares para o Pagamento STR
                If intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoSTR Then
                    With frmComplementoMensagensConciliacao
                        
                        Set .objSelectedItem = Nothing
                        .strComboFinalidade = strFinalidadeIF
                        .txtVeiculoLegal.Text = objListItemMesg.Text
                        .txtISPBIF.Text = objListItemMesg.SubItems(3)
                        .txtValor.Text = Replace$(objListItemMesg.SubItems(4), "-", vbNullString)
                        
                        For Each objListItemOper In Me.lvwOperacao.ListItems
                            If objListItemOper.Checked Then
                                .txtContraparte.Text = objListItemOper.SubItems(2)
                            End If
                        Next
                                
                        .Show vbModal
                        Call xmlDadosSTR.loadXML(.xmlComplemento.xml)
                        Set .xmlComplemento = Nothing
                        
                        If xmlDadosSTR.xml = vbNullString Then
                            With frmMural
                                .Display = "Informações complementares para o Pagamento STR não foram registradas. A seleção deste item será desfeita."
                                .IconeExibicao = IconCritical
                                .Show vbModal
                                
                                objListItemMesg.Checked = False
                                GoTo ProximoItem
                            End With
                        End If
                    End With
                End If
            
                Set xmlItemEnvio = CreateObject("MSXML2.DOMDocument.4.0")

                Call fgAppendNode(xmlItemEnvio, "", "Grupo_Envio", "")
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "NU_CTRL_IF", _
                                                Split(.Key, "|")(KEY_MSG_NU_CTRL_IF))
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "DH_REGT_MESG_SPB", _
                                                Split(.Key, "|")(KEY_MSG_DH_REGT_MESG_SPB))
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "NU_SEQU_CNTR_REPE", _
                                                Split(.Key, "|")(KEY_MSG_NU_SEQU_CNTR_REPE))
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "CO_ULTI_SITU_PROC", _
                                                Split(.Key, "|")(KEY_MSG_CO_ULTI_SITU_PROC))
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "CO_MESG_SPB", _
                                                "LTR0001")
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "CO_ISPB_BANC_LIQU_CNPT", _
                                                .SubItems(COL_M_CO_ISPB_BANC_LIQU_CNPT))
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "CONCILIACAO_FINANCEIRA_BILATERAL", _
                                                "")
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "IgnoraGradeHorario", _
                                                enumIndicadorSimNao.Sim)
                
                If xmlDadosSTR.xml <> vbNullString Then
                    For Each objDomNode In xmlDadosSTR.selectNodes("Repeat_Conciliacao/Grupo_Mensagem/*")
                        Call fgAppendXML(xmlItemEnvio, "Grupo_Envio", objDomNode.xml)
                    Next
                End If
                
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "Repeat_DetalheOperacao", _
                                                "")

                For Each objListItemOper In Me.lvwOperacao.ListItems
                    If objListItemOper.Checked Then
            
                        For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(objListItemOper.Key))
                            Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        
                            Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
                            Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                               "NU_SEQU_OPER_ATIV", _
                                                               objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                            Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                               "CO_ULTI_SITU_PROC", _
                                                               objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text)
        
                            Call fgAppendXML(xmlItemEnvio, "Repeat_DetalheOperacao", xmlItemOperacao.xml)
        
                            Set xmlItemOperacao = Nothing
                        Next
                        
                        Exit For
                    
                    End If
                Next

                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvio.xml)

                Set xmlItemEnvio = Nothing
                
                Exit For

            End If
        End With
    
ProximoItem:
    Next

    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLProcessamento = vbNullString
    Else
        flMontarXMLProcessamento = xmlProcessamento.xml
    End If

    Set xmlProcessamento = Nothing
    Set xmlDadosSTR = Nothing

    Exit Function

ErrorHandler:
    Set xmlProcessamento = Nothing
    Set xmlDadosSTR = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLProcessamento", 0

End Function

'Exibe o resultado da última operação executada
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

    On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " processados "
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'Obter o dominio da finalidade de mensagens STR,
'através da classe controladora de caso de uso MIU, método A8MIU.clsOperacaoMensagem.ObterDominioFinalidadeMsgSTR
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

'Monta a tela com os dados do filtro selecionado
Private Sub flPesquisar()

Dim strDocFiltros                           As String

    On Error GoTo ErrorHandler

    Call flLimparListas
    Call flCalcularTotais

    If Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If

    fgCursor True

    strDocFiltros = flMontarXMLFiltroPesquisa()
    Call flCarregarListaOperacoes(strDocFiltros)
    Call flCarregarListaMensagens(strDocFiltros)

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
        
        strXMLRetorno = objOperacaoMensagem.ConciliarCamaraLote(enumTipoConciliacao.Bilateral, _
                                                                intAcaoProcessamento, _
                                                                strXMLProc, _
                                                                vntCodErro, _
                                                                vntMensagemErro)

        Set objOperacaoMensagem = Nothing

        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

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

Dim objListItem                             As MSComctlLib.ListItem

Dim strOperIdentPartCamara                  As String
Dim strMesgIdentPartCamara                  As String

    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flValidarItensProcessamento = "Selecione um item da lista de mensagens, antes de prosseguir com a operação desejada."
        Exit Function
    End If

    If PerfilAcesso = BackOffice Then
        If intAcao <> BODiscordar And intAcao <> AdmGeralPagamentoContingencia Then
            For Each objListItem In Me.lvwOperacao.ListItems
                If objListItem.Checked Then
                    strOperIdentPartCamara = objListItem.SubItems(COL_O_CO_PARP_CAMR)
                    Exit For
                End If
            Next
    
            For Each objListItem In Me.lvwMensagem.ListItems
                If objListItem.Checked Then
                    strMesgIdentPartCamara = objListItem.SubItems(COL_M_CO_PARP_CAMR)
                    Exit For
                End If
            Next
    
            If strOperIdentPartCamara <> strMesgIdentPartCamara Then
                flValidarItensProcessamento = "Identificador de Participante Câmara do Net de Operações e Mensagens são diferentes. Operação não permitida."
                Exit Function
            End If
        End If
    End If

End Function

'Calcula o valor da operações
Private Function flValorOperacoes(ByVal strItemKey As String)

Dim strExpression                           As String
Dim vntValor                                As Variant

    vntValor = 0

    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlOperacoes, strExpression))

    flValorOperacoes = vntValor

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
            Me.Caption = "Conciliação Financeira Bilateral (Backoffice)"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "Conciliação Financeira Bilateral (Administrador de Área)"
    End Select

    Call flConfigurarBotoesPorPerfil(PerfilAcesso)
    Call flLimparListas

    If cboEmpresa.Text <> vbNullString Then
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

        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width

        .lvwMensagem.Top = .cboEmpresa.Top + .cboEmpresa.Height + 120
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwMensagem.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 780 - .fraComplementos.Height
        .lvwOperacao.Width = .Width - 240
        
        .fraComplementos.Top = .lvwOperacao.Top + .lvwOperacao.Height - 30
        .fraComplementos.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set frmConciliacaoFinanceiraBilateral = Nothing
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
        If .imgDummyH.Top > (.Height - 2500) And (.Height - 2500) > 0 Then
            .imgDummyH.Top = .Height - 2500
        End If

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 810 - .fraComplementos.Height
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

Dim objListItem                             As ListItem
Dim strSeqConciliacao                       As String

    On Error GoTo ErrorHandler
    
    If Item.Checked Then
        For Each objListItem In lvwMensagem.ListItems
            If objListItem.Key <> Item.Key Then
                objListItem.Checked = False
                objListItem.Selected = False
            End If
        Next
        Item.Selected = True
    End If
    
    If PerfilAcesso = enumPerfilAcesso.AdmArea Then
        strSeqConciliacao = Split(Item.Tag, "|")(TAG_MSG_NU_SEQU_CNCL)
        If Val(strSeqConciliacao) <> 0 Then
            For Each objListItem In lvwOperacao.ListItems
                If strSeqConciliacao = Split(objListItem.Tag, "|")(TAG_MSG_NU_SEQU_CNCL) Then
                    objListItem.Checked = Item.Checked
                    objListItem.Selected = True
                    objListItem.EnsureVisible
                Else
                    objListItem.Checked = False
                End If
            Next
        Else
            For Each objListItem In lvwOperacao.ListItems
                strSeqConciliacao = Split(objListItem.Tag, "|")(TAG_OP_NU_SEQU_CNCL)
                If Val(strSeqConciliacao) <> 0 Then
                    objListItem.Checked = False
                End If
            Next
        End If
    End If
    
    Call flCalcularTotais
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemCheck", Me.Caption

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
    
    If lvwOperacao.SelectedItem Is Nothing Then Exit Sub
    
    With frmComposicaoNetOperacoes
        .strXmlOperacoes = xmlOperacoes.xml
        .strCondicaoNavegacaoXml = flMontarCondicaoNavegacaoXMLOperacoes(lvwOperacao.SelectedItem.Key)
        
        .txtVeiculoLegal.Text = lvwOperacao.SelectedItem.Text
        .txtContraparte.Text = lvwOperacao.SelectedItem.SubItems(COL_O_NO_CNPT)
        .txtIdentPartCamara.Text = lvwOperacao.SelectedItem.SubItems(COL_O_CO_PARP_CAMR)
        .txtIdentPartCamaraContraparte.Text = lvwOperacao.SelectedItem.SubItems(COL_O_CO_CNPT_CAMR)
        
        .Show vbModal
    End With
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

End Sub

Private Sub lvwOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
Dim objListItem                             As ListItem
Dim strSeqConciliacao                       As String

    On Error GoTo ErrorHandler
    
    If Item.Checked Then
        For Each objListItem In lvwOperacao.ListItems
            If objListItem.Key <> Item.Key Then
                objListItem.Checked = False
                objListItem.Selected = False
            End If
        Next
        Item.Selected = True
    End If
    
    If PerfilAcesso = enumPerfilAcesso.AdmArea Then
        strSeqConciliacao = Split(Item.Tag, "|")(TAG_OP_NU_SEQU_CNCL)
        If Val(strSeqConciliacao) <> 0 Then
            For Each objListItem In lvwMensagem.ListItems
                If strSeqConciliacao = Split(objListItem.Tag, "|")(TAG_OP_NU_SEQU_CNCL) Then
                    objListItem.Checked = Item.Checked
                    objListItem.Selected = True
                    objListItem.EnsureVisible
                Else
                    objListItem.Checked = False
                End If
            Next
        Else
            For Each objListItem In lvwMensagem.ListItems
                strSeqConciliacao = Split(objListItem.Tag, "|")(TAG_MSG_NU_SEQU_CNCL)
                If Val(strSeqConciliacao) <> 0 Then
                    objListItem.Checked = False
                End If
            Next
        End If
    End If
    
    Call flCalcularTotais
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemCheck", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
Dim strValidaProcessamento                  As String
Dim strConfirmacao                          As String

    On Error GoTo ErrorHandler

    Button.Enabled = False: DoEvents
    intAcaoProcessamento = 0

    Select Case Button.Key
        Case "refresh"
            Call flPesquisar

        Case "concordancia"
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoProcessamento = enumAcaoConciliacao.BOConcordar
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia
                strConfirmacao = "a Concordância"
            End If
            
        Case "discordancia"
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoProcessamento = enumAcaoConciliacao.BODiscordar
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                strConfirmacao = "a Discordância"
            End If
            
        Case "retorno"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRejeitar
            strConfirmacao = "a Rejeição"
            
        Case "pagamento"
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoProcessamento = enumAcaoConciliacao.BOPagamento
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento
                strConfirmacao = "o Pagamento"
            End If
            
        Case "pagamentostr"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoSTR
            strConfirmacao = "o Pagamento via STR"
            
        Case "pagamentobacen"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoBACEN
            strConfirmacao = "o Pagamento via BACEN"
            
        Case "pagamentocontingencia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoContingencia
            strConfirmacao = "o Pagamento em Contingência"
            
        Case "regularizacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar
            strConfirmacao = "a Regularização"
            
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

        If strConfirmacao <> vbNullString Then
            If MsgBox("Confirma " & strConfirmacao & " do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        End If

        strResultadoOperacao = flProcessar

        If strResultadoOperacao <> vbNullString Then
            Call flMostrarResultado(strResultadoOperacao)
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
