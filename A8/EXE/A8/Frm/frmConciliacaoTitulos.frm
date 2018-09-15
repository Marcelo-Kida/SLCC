VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConciliacaoTitulos 
   Caption         =   "Ferramentas - Conciliação e Liquidação de Títulos (BMA)"
   ClientHeight    =   9915
   ClientLeft      =   225
   ClientTop       =   690
   ClientWidth     =   14205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboGrupoVeiculoLegal 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4350
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   9585
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   582
      ButtonWidth     =   2355
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Caption         =   "Concordar    "
            Key             =   "concordancia"
            Object.ToolTipText     =   "Concodar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Discordar    "
            Key             =   "discordancia"
            Object.ToolTipText     =   "Discordar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Liberar       "
            Key             =   "liberacao"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rejeitar      "
            Key             =   "retorno"
            Object.ToolTipText     =   "Retornar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contingência"
            Key             =   "transferencia"
            Object.ToolTipText     =   "Transf"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Regularizar  "
            Key             =   "regularizacao"
            Object.ToolTipText     =   "Regularizar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair             "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   5430
      Left            =   60
      TabIndex        =   2
      Top             =   4170
      Width           =   14085
      _ExtentX        =   24844
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
         Text            =   "Empresa"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   9480
      Top             =   60
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
            Picture         =   "frmConciliacaoTitulos.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoTitulos.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3255
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
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
   Begin A8.ctlMenu ctlMenu1 
      Left            =   12300
      Top             =   240
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin VB.Label lblConciliacao 
      AutoSize        =   -1  'True
      Caption         =   "Grupo Veículo Legal"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   1470
   End
   Begin VB.Image imgDummyH 
      Height          =   60
      Left            =   60
      MousePointer    =   7  'Size N S
      Top             =   4050
      Width           =   14040
   End
End
Attribute VB_Name = "frmConciliacaoTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Objeto responsavel pela conciliação de títulos,
' através da camada de controle de caso de uso MIU.
'
Option Explicit

'Constantes da chave da Mensagem
Private Const POS_TITULO                    As Integer = 1
Private Const POS_NUMERO_COMANDO            As Integer = 2
Private Const POS_CONTA                     As Integer = 3
Private Const POS_DTVENC                    As Integer = 4

Private Const COD_ERRO_NEGOCIO_GRADE        As Integer = 3023

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_AREA                  As Integer = 1
Private Const COL_MSG_TITULO                As Integer = 2
Private Const COL_MSG_DATA_VENCIMENTO       As Integer = 3
Private Const COL_MSG_CONTA                 As Integer = 4
Private Const COL_MSG_NUMERO_COMANDO        As Integer = 5
Private Const COL_MSG_QTDE_TOTAL            As Integer = 6
Private Const COL_MSG_QTDE_CAMARA           As Integer = 7
Private Const COL_MSG_DIFERENCA             As Integer = 8
Private Const COL_MSG_STATUS                As Integer = 9

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_VEICULO_LEGAL          As Integer = 0
Private Const COL_OP_CONTA                  As Integer = 1
Private Const COL_OP_DEBITO_CREDITO         As Integer = 2
Private Const COL_OP_TITULO                 As Integer = 3
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 4
Private Const COL_OP_QUANTIDADE             As Integer = 5
Private Const COL_OP_PU                     As Integer = 6
Private Const COL_OP_VALOR                  As Integer = 7
Private Const COL_OP_NUMERO_COMANDO         As Integer = 8
Private Const COL_OP_DATA_OPERACAO          As Integer = 9
Private Const COL_OP_EMPRESA                As Integer = 10

'Constantes de posicionamento de campos na propriedade Tag do item do ListView
Private Const POS_NU_CTRL_IF                As Integer = 1
Private Const POS_DH_REGT_MESG_SPB          As Integer = 2
Private Const POS_NU_SEQU_CNTR_REPE         As Integer = 3
Private Const POS_DH_ULTI_ATLZ              As Integer = 4
Private Const POS_CO_MESG_SPB               As Integer = 5

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmConciliacaoTitulos"

'Fim declaração constantes
'==============================================================================

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private fblnDummyH                          As Boolean
Private lngPerfil                           As Long
Private intAcaoConciliacao                  As enumAcaoConciliacao
Private intTipoConciliacao                  As enumTipoConciliacao

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Calcular a diferença dos valores das operações e mensagens SPB.

Private Sub flCalcularDiferencasListView()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double
Dim dblDiferenca                            As Double

On Error GoTo ErrorHandler

    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_QTDE_TOTAL)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_QTDE_CAMARA)))
            
            'Verfica se os dois valores tem o mesmo sinal
            If dblValorOperacao / Abs(IIf(dblValorOperacao = 0, 1, dblValorOperacao)) = dblValorMensagem / Abs(IIf(dblValorMensagem = 0, 1, dblValorMensagem)) Then
                dblDiferenca = Abs(dblValorOperacao - dblValorMensagem)
            Else
                dblDiferenca = Abs(dblValorOperacao) + Abs(dblValorMensagem)
            End If
            .SubItems(COL_MSG_DIFERENCA) = fgVlrXml_To_Interface(dblDiferenca)
            
            If dblDiferenca <> 0 Then
                .ListSubItems(COL_MSG_DIFERENCA).ForeColor = vbRed
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

    If Me.cboGrupoVeiculoLegal.ListIndex = -1 Or Me.cboGrupoVeiculoLegal.Text = vbNullString Then
        frmMural.Display = "Selecione o Grupo de Veículo Legal."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboGrupoVeiculoLegal.SetFocus
        Exit Sub
    End If

    fgCursor True

    strDocFiltros = flMontarXmlFiltro
    Call flCarregarMensagens(strDocFiltros)
    Call flCalcularDiferencasListView

    fgCursor

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarLista", 0)

End Sub

' Carrega o NET das da liquidações físicas e preencher a interface com os mesmos,
' através da classe controladora de caso de uso MIU, método A8MIU.clsMensagem.ObterNetLiquidacaoFisica

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
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterNetLiquidacaoFisica(pstrFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarMensagens")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_NetLiquidacaoFisica/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_ATIV_MERC").Text & _
                             "|" & objDomNode.selectSingleNode("NU_COMD_OPER").Text & _
                             "|" & objDomNode.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text & _
                             "|" & objDomNode.selectSingleNode("DT_VENC_ATIV").Text & "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text


            strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                             "|" & objDomNode.selectSingleNode("CO_MESG_SPB").Text
                             
            If strListItemTag = "||00:00:00|0|00:00:00|" Then
                strListItemTag = vbNullString
            End If

            With lvwMensagem.ListItems.Add(, strListItemKey)

                .SubItems(COL_MSG_AREA) = objDomNode.selectSingleNode("DE_BKOF").Text
                .SubItems(COL_MSG_TITULO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
                If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_VENCIMENTO) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                End If
                .SubItems(COL_MSG_CONTA) = objDomNode.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                .SubItems(COL_MSG_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_MSG_QTDE_TOTAL) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("QT_OPERACAO").Text)
                .SubItems(COL_MSG_QTDE_CAMARA) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("QT_CAMARA").Text)
                .SubItems(COL_MSG_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text

                'Adiciona chave da mensagem à propriedade Tag
                .Tag = strListItemTag

            End With

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

' Carrega as operações e preencher a interface com os mesmos,
' através da classe controladora de caso de uso MIU, método A8MIU.clsOperacao.ObterDetalheOperacao

Private Sub flCarregarOperacoes()

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim strFiltro                               As String
Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim vntValorTotal                           As Variant
Dim dblDiferenca                            As Double
Dim objListItem                             As ListItem

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String

Dim strTitulo                               As String
Dim strNumeroComando                        As String
Dim strConta                                As String
Dim strDataVencimento                       As String

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objListItem = lvwMensagem.SelectedItem

    If objListItem Is Nothing Then
        Exit Sub
    End If

    With objListItem
        strTitulo = Split(.Key, "|")(POS_TITULO)
        strNumeroComando = Split(.Key, "|")(POS_NUMERO_COMANDO)
        strConta = Split(.Key, "|")(POS_CONTA)
        strDataVencimento = Split(.Key, "|")(POS_DTVENC)
    End With

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlFiltro, "", "Repeat_Filtros", "")

    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Titulo", ""
    fgAppendNode xmlFiltro, "Grupo_Titulo", "Titulo", strTitulo

    If Not Trim$(strNumeroComando) = vbNullString Then
        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
        fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumeroComando", strNumeroComando
    End If

    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Status", ""
    If PerfilAcesso = enumPerfilAcesso.BackOffice Then
        fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.AConciliar
    Else
        fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice
        fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeAutomatico
        fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.DiscordanciaBackoffice
    End If

    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ContaSELIC", ""
    fgAppendNode xmlFiltro, "Grupo_ContaSELIC", "ContaSELIC", strConta

    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_DataVencimento", ""
    fgAppendNode xmlFiltro, "Grupo_DataVencimento", "DataIni", fgDtXML_To_Oracle(strDataVencimento)
    fgAppendNode xmlFiltro, "Grupo_DataVencimento", "DataFim", fgDtXML_To_Oracle(strDataVencimento)

    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux) + 1, "YYYYMMDD") & "000000"))

    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoOperacao", ""
    fgAppendNode xmlFiltro, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA

    strFiltro = xmlFiltro.xml
    Set xmlFiltro = Nothing

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(strFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing

    lvwOperacao.ListItems.Clear

    vntValorTotal = 0

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarOperacoes")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text

            With lvwOperacao.ListItems.Add(, strListItemKey)

                .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_OP_CONTA) = objDomNode.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                .SubItems(COL_OP_DEBITO_CREDITO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_OP_TITULO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
                .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                .SubItems(COL_OP_QUANTIDADE) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("QT_ATIV_MERC").Text)
                .SubItems(COL_OP_PU) = fgVlrXml_To_InterfaceDecimais(objDomNode.selectSingleNode("PU_ATIV_MERC").Text, 8)
                .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                .SubItems(COL_OP_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                .SubItems(COL_OP_EMPRESA) = objDomNode.selectSingleNode("CO_VEIC_LEGA").Text

'                If objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text = "Débito" Then
'                    vntValorTotal = vntValorTotal - Val(objDomNode.selectSingleNode("QT_ATIV_MERC").Text)
'                Else
'                    vntValorTotal = vntValorTotal + Val(objDomNode.selectSingleNode("QT_ATIV_MERC").Text)
'                End If

            End With

        Next
    End If

'    Me.lvwMensagem.SelectedItem.SubItems(COL_MSG_QTDE_TOTAL) = fgVlrXml_To_Interface(Abs(vntValorTotal))
'    dblDiferenca = fgVlrXml_To_Interface(Me.lvwMensagem.SelectedItem.SubItems(COL_MSG_QTDE_TOTAL) - Me.lvwMensagem.SelectedItem.SubItems(COL_MSG_QTDE_CAMARA))
'    Me.lvwMensagem.SelectedItem.SubItems(COL_MSG_DIFERENCA) = fgVlrXml_To_Interface(dblDiferenca)
'    Me.lvwMensagem.SelectedItem.SubItems(COL_MSG_DEBITO_CREDITO) = IIf(vntValorTotal >= 0, "Crédito", "Débito")

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
    Set xmlRetLeitura = Nothing

Exit Sub
ErrorHandler:
    
    Set objOperacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarOperacoes", 0)

End Sub

' Executar a conciliação efetuado através da camada controladora de casos de uso
' MIU, método A8MIU.clsOperacaoMensagem.ConciliarCamaraLote

Private Function flConciliar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem     As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem     As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLRetorno               As String
Dim strConciliacaoMsg           As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    strConciliacaoMsg = flMontarXMLConciliacao

    If InStr(1, strConciliacaoMsg, "NU_CTRL_IF") <> 0 Then
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
    
    Set objOperacaoMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flConciliar", Me.Caption

End Function

'Configurar os botões da tela conforme o perfil do usuário liberando ou não utilização
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)
    With tlbComandos
        .Buttons("concordancia").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice)
        .Buttons("discordancia").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice)
        .Buttons("liberacao").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("retorno").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("transferencia").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("regularizacao").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Refresh
    End With
End Sub

'Configurar os listview de mensagem e operação
Private Sub flFormatarListas()

    With lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Conciliar", 900
        .Add , , "Área", 1500
        .Add , , "Título", 1500
        .Add , , "Data Vencimento", 1440
        .Add , , "Conta", 1440
        .Add , , "Número Comando", 1500
        .Add , , "Qtde. Total", 1440, lvwColumnRight
        .Add , , "Qtde. Câmara", 1440, lvwColumnRight
        .Add , , "Diferença", 1440, lvwColumnRight
        .Add , , "Status", 2000
    End With

    With lvwOperacao.ColumnHeaders
        .Clear
        .Add , , "Veículo Legal", 1600
        .Add , , "Conta", 1600
        .Add , , "D/C", 1000
        .Add , , "Título", 1500
        .Add , , "Data Vencimento", 1440
        .Add , , "Quantidade", 1440, lvwColumnRight
        .Add , , "PU", 1440, lvwColumnRight
        .Add , , "Valor", 1440, lvwColumnRight
        .Add , , "Número Comando", 1500
        .Add , , "Data Operação", 1440
        .Add , , "Empresa", 0
    End With

End Sub

' Carrega as propriedades necessárias a interface frmCompromissadaGenerica, através da
' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

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
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If

    Call fgCarregarCombos(Me.cboGrupoVeiculoLegal, xmlMapaNavegacao, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA")

    Set objMIU = Nothing

Exit Function
ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Function

'Limpar os list view de mensagem e de operação
Private Sub flLimparListas()
    
    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear

End Sub

'Montar o xml para conciliação das informações apresentadas na interface

Private Function flMontarXMLConciliacao() As String

Dim objListItem                             As ListItem
Dim xmlConciliacaoMsg                       As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer

Dim strValidaConciliacao                    As String

On Error GoTo ErrorHandler

    strValidaConciliacao = flVerificarItensConciliacao
    If strValidaConciliacao <> vbNullString Then
        frmMural.Display = strValidaConciliacao
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Function
    End If

    Set xmlConciliacaoMsg = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlConciliacaoMsg, "", "Repeat_Conciliacao", "")

    For Each objListItem In Me.lvwMensagem.ListItems
        With objListItem
            If .Checked Then

                Call fgAppendNode(xmlConciliacaoMsg, "Repeat_Conciliacao", "Grupo_Mensagem", "")
                Call fgAppendNode(xmlConciliacaoMsg, "Grupo_Mensagem", _
                                                     "NU_CTRL_IF", _
                                                     Split(.Tag, "|")(POS_NU_CTRL_IF), _
                                                     "Repeat_Conciliacao")
                Call fgAppendNode(xmlConciliacaoMsg, "Grupo_Mensagem", _
                                                     "DH_REGT_MESG_SPB", _
                                                     Split(.Tag, "|")(POS_DH_REGT_MESG_SPB), _
                                                     "Repeat_Conciliacao")
                Call fgAppendNode(xmlConciliacaoMsg, "Grupo_Mensagem", _
                                                     "NU_SEQU_CNTR_REPE", _
                                                     Split(.Tag, "|")(POS_NU_SEQU_CNTR_REPE), _
                                                     "Repeat_Conciliacao")
                Call fgAppendNode(xmlConciliacaoMsg, "Grupo_Mensagem", _
                                                     "DH_ULTI_ATLZ", _
                                                     Split(.Tag, "|")(POS_DH_ULTI_ATLZ), _
                                                     "Repeat_Conciliacao")
                Call fgAppendNode(xmlConciliacaoMsg, "Grupo_Mensagem", _
                                                     "CO_MESG_SPB", _
                                                     Split(.Tag, "|")(POS_CO_MESG_SPB), _
                                                     "Repeat_Conciliacao")

                intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)

                Call fgAppendNode(xmlConciliacaoMsg, "Grupo_Mensagem", _
                                                     "IgnoraGradeHorario", _
                                                     intIgnoraGradeHorario, _
                                                     "Repeat_Conciliacao")

            End If
        End With
    Next

    flMontarXMLConciliacao = xmlConciliacaoMsg.xml

    Set xmlConciliacaoMsg = Nothing

Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLConciliacao", 0

End Function

'Montar o xml de filtro para pesquisa pelo perfil do usuário
Private Function flMontarXmlFiltro() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim strSelecaoFiltroOper                    As String
Dim strSelecaoFiltroMsg                     As String
Dim strDocFiltrosAux                        As String
Dim lngCont                                 As Long

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    fgAppendNode xmlFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", ""
    Call fgAppendNode(xmlFiltros, "Grupo_GrupoVeiculoLegal", "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeiculoLegal.Text))

    Select Case PerfilAcesso
        Case enumPerfilAcesso.BackOffice
            strSelecaoFiltroOper = enumStatusOperacao.AConciliar
            strSelecaoFiltroMsg = enumStatusMensagem.AConciliar

        Case enumPerfilAcesso.AdmArea
            strSelecaoFiltroOper = enumStatusOperacao.ConcordanciaBackoffice & ";" & _
                                   enumStatusOperacao.DiscordanciaBackoffice
            strSelecaoFiltroMsg = enumStatusMensagem.ConcordanciaBackoffice & ";" & _
                                  enumStatusMensagem.DiscordanciaBackoffice
                              
        Case enumPerfilAcesso.AdmGeral
            strSelecaoFiltroOper = enumStatusOperacao.ConcordanciaAdmArea & ";" & _
                                   enumStatusOperacao.DiscordanciaAdmArea
            strSelecaoFiltroMsg = enumStatusMensagem.ConcordanciaAdmArea & ";" & _
                                  enumStatusMensagem.DiscordanciaAdmArea

    End Select

    fgAppendNode xmlFiltros, "Repeat_Filtros", "Grupo_StatusOperacao", ""
    'Captura o filtro cumulativo para operação
    For lngCont = LBound(Split(strSelecaoFiltroOper, ";")) To UBound(Split(strSelecaoFiltroOper, ";"))
        Call fgAppendNode(xmlFiltros, "Grupo_StatusOperacao", "Status", Split(strSelecaoFiltroOper, ";")(lngCont))
    Next

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)))
    Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataFim", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)))


    fgAppendNode xmlFiltros, "Repeat_Filtros", "Grupo_StatusMensagem", ""
    'Captura o filtro cumulativo para mensagem
    For lngCont = LBound(Split(strSelecaoFiltroMsg, ";")) To UBound(Split(strSelecaoFiltroMsg, ";"))
        Call fgAppendNode(xmlFiltros, "Grupo_StatusMensagem", "Status", Split(strSelecaoFiltroMsg, ";")(lngCont))
    Next

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacaoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_DataOperacaoMensagem", "DataIni", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)))
    Call fgAppendNode(xmlFiltros, "Grupo_DataOperacaoMensagem", "DataFim", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data) + 1))

    fgAppendNode xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", ""
    Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Multilateral)

    fgAppendNode xmlFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", ""
    Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA)

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

' Retorna uma String referente a um preenchimento incorreto na interface. Se
' todos os campos estiverem preenchidos corretamente, retorna vbNullString
Private Function flVerificarItensConciliacao() As String

Dim objListItem                             As ListItem
Dim dblValorConsist                         As Double

    For Each objListItem In Me.lvwMensagem.ListItems
        If objListItem.Checked Then
            dblValorConsist = fgVlrXml_To_Decimal(fgVlr_To_Xml(objListItem.SubItems(COL_MSG_DIFERENCA)))

'            If objListItem.Tag = vbNullString Then
'                flVerificarItensConciliacao = "Não existe mensagem da BMA para conciliar, ação cancelada."
'                Exit Function
'            End If

            Select Case intAcaoConciliacao
                Case enumAcaoConciliacao.BOConcordar
                    If objListItem.Tag = vbNullString Then
                        flVerificarItensConciliacao = "A conciliação não pode ser efetuada pois existem um ou mais registros que não possuem mensagem BMA."
                        Exit Function
                    End If
                    
                    If dblValorConsist <> 0 Then
                        flVerificarItensConciliacao = "Existem um ou mais registros, com diferença entre valor de mensagem, e operação, diferentes de zero. Solicitação de Concordância não permitida."
                        Exit Function
                    End If

                Case enumAcaoConciliacao.BODiscordar
                    If objListItem.Tag = vbNullString Then
                        flVerificarItensConciliacao = "A dicordância não pode ser efetuada pois existem um ou mais registros que não possuem mensagem BMA."
                        Exit Function
                    End If
                    
                    If dblValorConsist = 0 Then
                        flVerificarItensConciliacao = "Existem um ou mais registros, com diferença entre valor de mensagem, e operação, iguais a zero. Solicitação de Discordância não permitida."
                        Exit Function
                    End If
                
                Case enumAcaoConciliacao.AdmAreaLiberar
                    If dblValorConsist <> 0 Then
                        flVerificarItensConciliacao = "Existem um ou mais registros, com diferença entre valor de mensagem, e operação, diferentes de zero. Liberação não permitida."
                        Exit Function
                    End If

                Case enumAcaoConciliacao.AdmAreaRegularizar
                    If dblValorConsist <> 0 Then
                        flVerificarItensConciliacao = "Existem um ou mais registros, com diferença entre valor de mensagem, e operação, diferentes de zero. Regularização não permitida."
                        Exit Function
                    End If

            End Select
        End If
    Next

End Function

Property Get PerfilAcesso() As enumPerfilAcesso
    'Retorna o perfil de acesso do usuário

    PerfilAcesso = lngPerfil

End Property

Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
    'Controla o perfil de acesso do usuário

    lngPerfil = pPerfil
    flConfigurarBotoesPorPerfil PerfilAcesso

End Property

Private Sub cboGrupoVeiculoLegal_Click()

On Error GoTo ErrorHandler

    If cboGrupoVeiculoLegal.Text <> vbNullString Then
        Call flCarregarLista
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboGrupoVeiculoLegal_Click", Me.Caption

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

    Call fgCenterMe(Me)
    Set Me.Icon = mdiLQS.Icon

    fgCursor True

    Call flInicializar
    Call flFormatarListas

    fgCursor

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
        If .imgDummyH.Top > (.Height - 3500) And (.Height - 3500) > 0 Then
            .imgDummyH.Top = .Height - 3500
        End If

        .lvwMensagem.Top = .cboGrupoVeiculoLegal.Top + .cboGrupoVeiculoLegal.Height + 120
        .lvwMensagem.Left = .cboGrupoVeiculoLegal.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .imgDummyH.Top = .lvwMensagem.Top + .lvwMensagem.Height
        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboGrupoVeiculoLegal.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
        .lvwOperacao.Width = .Width - 240
    End With

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

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fblnDummyH = False
End Sub

Private Sub lvwMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwMensagem, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_DblClick", Me.Caption

End Sub

Private Sub lvwMensagem_DblClick()

Dim strChave                                As String

On Error GoTo ErrorHandler

    If Not lvwMensagem.SelectedItem Is Nothing Then
        strChave = lvwMensagem.SelectedItem.Tag
        If strChave = vbNullString Then
            frmMural.Display = "Não existe mensagem para exibição do detalhe."
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            Exit Sub
        Else
            With frmDetalheOperacao
                .NumeroControleIF = Split(strChave, "|")(POS_NU_CTRL_IF)
                .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(strChave, "|")(POS_DH_REGT_MESG_SPB))
                .NumeroSequenciaRepeticao = Split(strChave, "|")(POS_NU_SEQU_CNTR_REPE)
                .Show vbModal
            End With
        End If
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwMensagem_DblClick"
End Sub

Private Sub lvwMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Item.Selected = True
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemCheck", Me.Caption

End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    fgCursor True
    Call flCarregarOperacoes
    fgCursor

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemClick", Me.Caption

End Sub

Private Sub lvwMensagem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
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
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

End Sub

Private Sub optNaturezaMovimento_Click(Index As Integer)

On Error GoTo ErrorHandler

    If cboGrupoVeiculoLegal.Text <> vbNullString Then
        Call flCarregarLista
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optNaturezaMovimento_Click", Me.Caption

End Sub

Private Sub lvwOperacao_DblClick()

On Error GoTo ErrorHandler

    If Not lvwOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .CodigoEmpresa = lvwOperacao.SelectedItem.ListSubItems(COL_OP_EMPRESA)
            .SequenciaOperacao = Mid(lvwOperacao.SelectedItem.Key, 2)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwOperacao_DblClick"
End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String

On Error GoTo ErrorHandler

    Button.Enabled = False: DoEvents
    fgCursor True

    intTipoConciliacao = enumTipoConciliacao.MultilateralTitulos
    intAcaoConciliacao = 0

    Select Case Button.Key
        Case "refresh"
            Call flCarregarLista

        Case "concordancia"
            intAcaoConciliacao = enumAcaoConciliacao.BOConcordar

        Case "discordancia"
            intAcaoConciliacao = enumAcaoConciliacao.BODiscordar

        Case "liberacao"
            intAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberar

        Case "retorno"
            intAcaoConciliacao = enumAcaoConciliacao.AdmAreaRejeitar

        Case "transferencia"
            intAcaoConciliacao = enumAcaoConciliacao.AdmAreaPagamentoContingencia

        Case "regularizacao"
            intAcaoConciliacao = enumAcaoConciliacao.AdmAreaRegularizar

        Case gstrSair
            Unload Me
            
    End Select

    If intAcaoConciliacao <> 0 Then
        strResultadoOperacao = flConciliar

        If strResultadoOperacao <> vbNullString Then
            Call flMostrarResultado(strResultadoOperacao)
            Call flCarregarLista
            Call flMarcarRejeitadosPorGradeHorario
        End If
    End If

Exit Sub
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub

'Marcar as mensagens SPB que deveriam ser enviadas mas retornaram rejeitadas por grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlRetorno                              As MSXML2.DOMDocument40
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            With lvwMensagem.ListItems
                For lngCont = 1 To .Count
                    If .Item(lngCont).Tag <> vbNullString Then
                        If UCase(Split(.Item(lngCont).Tag, "|")(1)) = UCase(objDomNode.selectSingleNode("NumeroControleIF").Text) _
                        And UCase(Split(.Item(lngCont).Tag, "|")(2)) = UCase(objDomNode.selectSingleNode("DTRegistroMensagemSPB").Text) _
                        And UCase(Split(.Item(lngCont).Tag, "|")(3)) = UCase(objDomNode.selectSingleNode("SequenciaRepeticao").Text) Then
    
                            For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                                .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                            Next
                            
                            .Item(lngCont).Text = "Horário Excedido"
                            .Item(lngCont).ToolTipText = "Horário limite excedido. Comande a operação novamente para ignorar grade."
                            .Item(lngCont).ForeColor = vbRed
                            
                            Exit For
                        End If
                    End If
                Next
            End With
        Next
    End If

Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0
    Exit Sub
    Resume

End Sub

