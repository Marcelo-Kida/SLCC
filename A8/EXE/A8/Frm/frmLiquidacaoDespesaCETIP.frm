VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLiquidacaoDespesaCETIP 
   Caption         =   "Liquidação Despesa CETIP - Administrador Geral"
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
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4350
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   7290
      Top             =   810
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
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoDespesaCETIP.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   2625
      Left            =   60
      TabIndex        =   1
      Top             =   4740
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   4630
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
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
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
         NumButtons      =   9
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
            Caption         =   "Concordar   "
            Key             =   "concordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Discordar    "
            Key             =   "discordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagamento "
            Key             =   "pagamento"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pg. Conting."
            Key             =   "pagamentocontingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Regularizar  "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair            "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3885
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6853
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
   Begin VB.Frame fraComplementos 
      Height          =   765
      Left            =   60
      TabIndex        =   5
      Top             =   7530
      Width           =   12855
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   2235
      End
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
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total NET Operações"
         Height          =   195
         Index           =   1
         Left            =   4290
         TabIndex        =   9
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total LTR0001"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   7
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Image imgDummyH 
      Height          =   60
      Left            =   30
      MousePointer    =   7  'Size N S
      Top             =   4650
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
Attribute VB_Name = "frmLiquidacaoDespesaCETIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Liquidação Despesas CETIP

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_NO_VEIC_LEGA          As Integer = 0
Private Const COL_MSG_NU_COMD_OPER          As Integer = 1
Private Const COL_MSG_VA_FINC               As Integer = 2
Private Const COL_MSG_NU_CTRL_CAMR          As Integer = 3

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const KEY_MSG_NU_CTRL_IF            As Integer = 1
Private Const KEY_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const KEY_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const KEY_MSG_DH_ULTI_ATLZ          As Integer = 4
Private Const KEY_MSG_CO_MESG_SPB           As Integer = 5
Private Const KEY_MSG_NU_COMD_OPER          As Integer = 6

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_DE_BKOF                As Integer = 0
Private Const COL_OP_VA_OPER_ATIV           As Integer = 1
Private Const COL_OP_CNPJ_CPF_COMITENTE     As Integer = 2

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const KEY_OP_TP_BKOF                As Integer = 1

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmLiquidacaoBruta"

'Constantes de erros de negócio específicos
Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3095
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
                dblValorMensagem = dblValorMensagem + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VA_FINC)))
            End If
        End With
    Next

    txtTotalMensagem.Text = fgVlrXml_To_Interface(dblValorMensagem)

    dblValorOperacao = 0
    For Each objListItem In lvwOperacao.ListItems
        With objListItem
            If .Checked Then
                dblValorOperacao = dblValorOperacao + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_OP_VA_OPER_ATIV)))
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

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
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
        Call xmlRetLeitura.loadXML(strRetLeitura)
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                             "|" & objDomNode.selectSingleNode("CO_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_COMD_OPER").Text
                             
            With lvwMensagem.ListItems.Add(, strListItemKey)
                
                .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_MSG_NU_COMD_OPER) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_MSG_VA_FINC) = "-" & fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                .SubItems(COL_MSG_NU_CTRL_CAMR) = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                
            End With
                
        Next
    End If
    
    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaMensagens", 0)

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
Dim dblValorOperacao                        As Double
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

            If Not fgExisteItemLvw(Me.lvwOperacao, strListItemKey) Then
                dblValorOperacao = flValorOperacoes(strListItemKey) * -1

                With lvwOperacao.ListItems.Add(, strListItemKey)

                    .Text = objDomNode.selectSingleNode("DE_BKOF").Text
                    .SubItems(COL_OP_VA_OPER_ATIV) = fgVlrXml_To_Interface(dblValorOperacao)
                    .SubItems(COL_OP_CNPJ_CPF_COMITENTE) = objDomNode.selectSingleNode("NR_CNPJ_CPF").Text

                End With
            End If
        Next
    End If

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes", 0)

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

'Formata as colunas da lista de mensagens
Private Sub flInicializarLvwMensagem()

On Error GoTo ErrorHandler
    
    With Me.lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Cliente", 4800
        .Add , , "Número Comando", 1600
        .Add , , "Valor", 1600, lvwColumnRight
        .Add , , "Ident. Part. Câmara", 2000
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
        .Add , , "Área Responsável", 4800
        .Add , , "Valor", 1600, lvwColumnRight
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

'Marcar as mensagens SPB que deveriam ser enviadas mas retornaram rejeitadas por grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            With lvwMensagem.ListItems
                For lngCont = 1 To .Count
                    If Split(.Item(lngCont).Key, "|")(KEY_MSG_NU_CTRL_IF) = objDomNode.selectSingleNode("NumeroControleIF").Text And _
                       Split(.Item(lngCont).Key, "|")(KEY_MSG_DH_REGT_MESG_SPB) = objDomNode.selectSingleNode("DTRegistroMensagemSPB").Text And _
                       Split(.Item(lngCont).Key, "|")(KEY_MSG_NU_SEQU_CNTR_REPE) = objDomNode.selectSingleNode("SequenciaRepeticao").Text Then
                        
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
        Next
    End If

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0

End Sub

'Monta o conteúdo que será utilizado com a propriedade 'Key' dos itens do ListView
Private Function flMontarChaveItemListview(ByVal objDomNode As MSXML2.IXMLDOMNode)

Dim strListItemKey                          As String

On Error GoTo ErrorHandler

    strListItemKey = "|" & objDomNode.selectSingleNode("TP_BKOF").Text
    
    flMontarChaveItemListview = strListItemKey

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String)

Dim strDebito                               As String
Dim strCredito                              As String

On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_OP_TP_BKOF) & "' "

    strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                      " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_OP_TP_BKOF) & "' "

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
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.Registrada)
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica)
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LayoutEntrada", "LayoutEntrada", "88")
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bruta)
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoContraparte", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo)
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")
    
    flMontarXMLFiltroPesquisa = xmlFiltros.xml
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Monta XML com as chaves das operações que serão processadas
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer

On Error GoTo ErrorHandler

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    For Each objListItem In Me.lvwMensagem.ListItems
        With objListItem
            If .Checked Then

                Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")

                Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_Mensagem", "")
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "CO_EMPR", _
                                                   fgObterCodigoCombo(cboEmpresa.Text))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "NU_CTRL_IF", _
                                                   Split(.Key, "|")(KEY_MSG_NU_CTRL_IF))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "DH_REGT_MESG_SPB", _
                                                   Split(.Key, "|")(KEY_MSG_DH_REGT_MESG_SPB))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "NU_SEQU_CNTR_REPE", _
                                                   Split(.Key, "|")(KEY_MSG_NU_SEQU_CNTR_REPE))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "DH_ULTI_ATLZ", _
                                                   Split(.Key, "|")(KEY_MSG_DH_ULTI_ATLZ))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "CO_MESG_SPB", _
                                                   Split(.Key, "|")(KEY_MSG_CO_MESG_SPB))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "NU_COMD_OPER", _
                                                   Split(.Key, "|")(KEY_MSG_NU_COMD_OPER))

                intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "IgnoraGradeHorario", _
                                                   intIgnoraGradeHorario)

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_Mensagem", _
                                                   "Repeat_Operacao", _
                                                   "")

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

'Exibe o resultado da última operação executada
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
    Call flCalcularTotais
    
    If Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    fgCursor True
    
    strDocFiltros = flMontarXMLFiltroPesquisa
    Call flCarregarListaMensagens(strDocFiltros)
    Call flCarregarListaNetOperacoes(strDocFiltros)
    
    With tlbComandos
        .Buttons("concordancia").Enabled = True
        .Buttons("discordancia").Enabled = True
        .Buttons("pagamento").Enabled = True
        .Buttons("pagamentocontingencia").Enabled = True
        .Buttons("regularizacao").Enabled = True
        .Refresh
    End With
    
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
        strXMLRetorno = objOperacaoMensagem.ConciliarCamaraLote(enumTipoConciliacao.Bruta, _
                                                                intAcaoProcessamento, _
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

    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flValidarItensProcessamento = "Selecione um item da lista de mensagens, antes de prosseguir com a operação desejada."
        Exit Function
    End If

End Function

'Calcula o valor da operações
Private Function flValorOperacoes(ByVal strItemKey As String)

    Dim strExpression                   As String
    Dim vntValor                        As Variant

    vntValor = 0

    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlOperacoes, strExpression))

    flValorOperacoes = vntValor

End Function

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
    Call flInicializarLvwMensagem
    Call flInicializarLvwOperacao
    Call flLimparListas
    Call flCalcularTotais
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
        
        .lvwMensagem.Top = .cboEmpresa.Top + .cboEmpresa.Height + 240
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 780 - .fraComplementos.Height
        .lvwOperacao.Width = .Width - 240
    
        .fraComplementos.Top = .lvwOperacao.Top + .lvwOperacao.Height
        .fraComplementos.Left = .cboEmpresa.Left
        .fraComplementos.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set frmLiquidacaoDespesaCETIP = Nothing
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

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 780 - .fraComplementos.Height
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

On Error GoTo ErrorHandler
    
    Item.Selected = True
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

Private Sub lvwOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim objListItem                             As ListItem
    
On Error GoTo ErrorHandler
    
    Item.Selected = True
    Call flCalcularTotais
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_ItemCheck", Me.Caption

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
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia
            
        Case "discordancia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarDiscordancia
            
        Case "pagamento"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento
            
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
            frmMural.Display = strValidaProcessamento
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            GoTo ExitSub
        End If
        
        If intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia Then
            If MsgBox("Confirma a concordância do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Then
            If MsgBox("Confirma a discordância do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento Then
            If MsgBox("Confirma o pagamento do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoContingencia Then
            If MsgBox("Confirma o pagamento em contingência do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        ElseIf intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar Then
            If MsgBox("Confirma a regularização do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                GoTo ExitSub
            End If
        End If
        
        strResultadoOperacao = flProcessar
        
        If strResultadoOperacao <> vbNullString Then
            Call flMostrarResultado(strResultadoOperacao)
            Call flPesquisar
        End If
    
        Call flMarcarRejeitadosPorGradeHorario
    End If
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
