VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiquidacaoBruta 
   Caption         =   "CBLC - Liquidação Bruta"
   ClientHeight    =   8625
   ClientLeft      =   975
   ClientTop       =   855
   ClientWidth     =   12975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   12975
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOptions 
      Caption         =   "Tipo Transferência"
      Height          =   570
      Index           =   1
      Left            =   4500
      TabIndex        =   11
      Top             =   60
      Width           =   3585
      Begin VB.OptionButton optTipoTransf 
         Caption         =   "Todos"
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   3
         Top             =   270
         Width           =   765
      End
      Begin VB.OptionButton optTipoTransf 
         Caption         =   "Book Transfer"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optTipoTransf 
         Caption         =   "Mercado"
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame fraNaturezaMovto 
      Caption         =   "Natureza Movimento"
      Height          =   570
      Left            =   8190
      TabIndex        =   10
      Top             =   60
      Width           =   3735
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "Todos"
         Height          =   195
         Index           =   2
         Left            =   2850
         TabIndex        =   6
         Top             =   270
         Width           =   765
      End
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "Recebimento"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton optNaturezaMovimento 
         Caption         =   "Pagamento"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
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
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   7650
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
            Picture         =   "frmLiquidacaoBruta.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacaoBruta.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   3255
      Left            =   60
      TabIndex        =   7
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
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   8295
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
         NumButtons      =   8
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
            Caption         =   "Pagamento "
            Key             =   "pagamento"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Pg. Conting."
            Key             =   "pagamentocontingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Regularizar  "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   8
      Top             =   4140
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6853
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
      TabIndex        =   9
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmLiquidacaoBruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Liquidação Bruta de Operações CBLC

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_PARTICIP_CAMARA       As Integer = 0
Private Const COL_MSG_NUMERO_COMANDO        As Integer = 1
Private Const COL_MSG_DEB_CRED              As Integer = 2
Private Const COL_MSG_VALOR                 As Integer = 3
Private Const COL_MSG_TIPO_TRANSF           As Integer = 4
Private Const COL_MSG_STATUS                As Integer = 5

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const KEY_MSG_NU_CTRL_IF            As Integer = 1
Private Const KEY_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const KEY_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const KEY_MSG_CO_MESG_SPB           As Integer = 4
Private Const KEY_MSG_IN_OPER_DEBT_CRED     As Integer = 5

'Constantes de posicionamento de campos na propriedade Tag do item do ListView de Mensagens
Private Const TAG_MSG_NU_SEQU_OPER_ATIV     As Integer = 1

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_PARTICIP_NEG           As Integer = 0
Private Const COL_OP_VEICULO_LEGAL          As Integer = 1
Private Const COL_OP_NUMERO_COMANDO         As Integer = 2
Private Const COL_OP_DEB_CRED               As Integer = 3
Private Const COL_OP_VALOR                  As Integer = 4
Private Const COL_OP_IDENTIF_ATIVO          As Integer = 5
Private Const COL_OP_STATUS                 As Integer = 6
Private Const COL_OP_TIPO_TRANSF            As Integer = 7

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const KEY_OP_NU_SEQU_OPER_ATIV      As Integer = 1

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmLiquidacaoBruta"
'------------------------------------------------------------------------------------------
'Fim declaração constantes

Private Enum enumTipoTransferencia
    BookTransfer = 0
    Mercado = 1
    Todos = 2
End Enum

Private Enum enumNaturezaMovimento
    Pagamento = 0
    Recebimento = 1
    Todos = 2
End Enum

Private Enum enumTipoPesquisa
    Operacao = 0
    MENSAGEM = 1
End Enum

Private lngPerfil                           As Long
Private intAcaoProcessamento                As enumAcaoConciliacao
Private blnDummyH                           As Boolean
Private blnAcaoContingencia                 As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

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
        
        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaMensagens")
        End If
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("CO_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text
                             
            With lvwMensagem.ListItems.Add(, strListItemKey)
                
                .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                
                If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                    .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                End If
                
                .SubItems(COL_MSG_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_MSG_DEB_CRED) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_MSG_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                .SubItems(COL_MSG_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                
                If objDomNode.selectSingleNode("CO_MESG_SPB").Text <> "LTR0007" Then
                    .SubItems(COL_MSG_TIPO_TRANSF) = "Mercado"
                Else
                    .SubItems(COL_MSG_TIPO_TRANSF) = "Book Transfer"
                End If
                
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

'Carregar lista de operações
Private Sub flCarregarListaOperacoes(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim strRetLeitura                           As String
Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
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
    
    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaOperacoes")
        End If
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")
            
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                    
            With lvwOperacao.ListItems.Add(, strListItemKey)
                
                .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                
                If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                    .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                End If
                
                .SubItems(COL_OP_VEICULO_LEGAL) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_OP_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                .SubItems(COL_OP_IDENTIF_ATIVO) = objDomNode.selectSingleNode("NU_ATIV_MERC").Text
                .SubItems(COL_OP_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                
                If Val(objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    .SubItems(COL_OP_DEB_CRED) = "Débito"
                Else
                    .SubItems(COL_OP_DEB_CRED) = "Crédito"
                End If
                
                If Val(objDomNode.selectSingleNode("TP_CNPT").Text) = enumTipoContraparte.Externo Then
                    .SubItems(COL_OP_TIPO_TRANSF) = "Mercado"
                Else
                    .SubItems(COL_OP_TIPO_TRANSF) = "Book Transfer"
                End If
                
            End With
        Next
    End If
    
    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaOperacoes", 0)

End Sub

'Habilita ou desabilita botões de acordo com a funcionalidade
Private Sub flConfigurarBotoesPorFuncionalidade(ByVal blnEnabled As Boolean, _
                                                ByVal objListItem As ListItem, _
                                                ByVal intTipoPesquisa As enumTipoPesquisa)
    
Dim objLvwPrinc                             As ListView
Dim objLvwSec                               As ListView
Dim objListItemAux                          As ListItem
    
On Error GoTo ErrorHandler
    
    If intTipoPesquisa = MENSAGEM Then
        Set objLvwPrinc = Me.lvwMensagem
        Set objLvwSec = Me.lvwOperacao
    Else
        Set objLvwPrinc = Me.lvwOperacao
        Set objLvwSec = Me.lvwMensagem
    End If
    
    With tlbComandos
        .Buttons("concordancia").Enabled = blnEnabled
        .Buttons("pagamento").Enabled = blnEnabled
        .Buttons("pagamentocontingencia").Enabled = Not blnEnabled
        .Buttons("regularizacao").Enabled = blnEnabled
        .Refresh
    End With
    
    If Not blnEnabled Then
        For Each objListItemAux In objLvwPrinc.ListItems
            If objListItemAux.Key <> objListItem.Key Then
                objListItemAux.Checked = False
                objListItemAux.Selected = False
            End If
        Next
        
        If Not blnAcaoContingencia Then
            For Each objListItemAux In objLvwSec.ListItems
                objListItemAux.Checked = False
                objListItemAux.Selected = False
            Next
        End If
    
        blnAcaoContingencia = True
        
    Else
        For Each objListItemAux In objLvwPrinc.ListItems
            If objListItemAux.Tag = vbNullString Then
                objListItemAux.Checked = False
                objListItemAux.Selected = False
            End If
        Next
        
        For Each objListItemAux In objLvwSec.ListItems
            If objListItemAux.Tag = vbNullString Then
                objListItemAux.Checked = False
                objListItemAux.Selected = False
            End If
        Next
    
        blnAcaoContingencia = False
        
    End If

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorFuncionalidade", 0

End Sub

'Altera a exibição dos botões de acordo com o perfil do usuário
Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)
    
On Error GoTo ErrorHandler
    
    With tlbComandos
        .Buttons("concordancia").Visible = True
        .Buttons("pagamento").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("pagamentocontingencia").Visible = .Buttons("pagamento").Visible
        .Buttons("regularizacao").Visible = .Buttons("pagamento").Visible
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

'Formata as colunas da lista de mensagens
Private Sub flInicializarLvwMensagem()

On Error GoTo ErrorHandler
    
    With Me.lvwMensagem.ColumnHeaders
        .Clear
        
        .Add , , "Participante Negociação", 2300
        .Add , , "Número Comando", 1500
        .Add , , "D/C", 800
        .Add , , "Valor", 1600, lvwColumnRight
        .Add , , "Tipo Transferência", 1700
        .Add , , "Status", 2000
            
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
        .Add , , "Participante Negociação", 2300
        .Add , , "Veículo Legal", 3100
        .Add , , "Número Comando", 1500
        .Add , , "D/C", 800
        .Add , , "Valor", 1600, lvwColumnRight
        .Add , , "Identificador Ativo", 1700
        .Add , , "Status", 2000
        .Add , , "Tipo Transferência", 1700
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

'Exibe de forma diferenciada os itens que tenham sido rejeitados por motivo de grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim objListItem                             As MSComctlLib.ListItem
'Dim intCont                                 As Integer
'
'    On Error GoTo ErrorHandler
'
'    If Not xmlRetornoErro Is Nothing Then
'        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='3095' or CodigoErro='3023']")
'            For Each objListItem In lvwMensagem.ListItems
'                With objListItem
'                    If Split(.Key, "|")(KEY_MSG_CO_LOCA_LIQU) = objDomNode.selectSingleNode("CO_LOCA_LIQU").Text And _
'                       Split(.Key, "|")(KEY_MSG_CO_ISPB_CNPT) = objDomNode.selectSingleNode("CO_ISPB_CNPT").Text And _
'                       Split(.Key, "|")(KEY_MSG_CO_CNPJ_CNPT) = objDomNode.selectSingleNode("CO_CNPJ_CNPT").Text And _
'                       Split(.Key, "|")(KEY_MSG_TP_IF_CRED_DEBT) = objDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text And _
'                       Split(.Key, "|")(KEY_MSG_CO_AGEN_COTR) = objDomNode.selectSingleNode("CO_AGEN_COTR").Text And _
'                       Split(.Key, "|")(KEY_MSG_NU_CC_COTR) = objDomNode.selectSingleNode("NU_CC_COTR").Text Then
'
'                        For intCont = 1 To .ListSubItems.Count
'                            .ListSubItems(intCont).ForeColor = vbRed
'                        Next
'
'                        .Text = "Horário Excedido"
'                        .ToolTipText = "Horário limite p/envio da mensagem excedido"
'                        .ForeColor = vbRed
'
'                        Exit For
'
'                    End If
'                End With
'            Next
'        Next
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorGradeHorario", 0

End Sub

'Monta o XML com os dados de filtro para seleção de operações
Private Function flMontarXMLFiltroPesquisa(ByVal intTipoPesquisa As enumTipoPesquisa) As String
    
Dim xmlFiltros                              As MSXML2.DOMDocument40
    
On Error GoTo ErrorHandler
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CLBCAcoes)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")
    
    If intTipoPesquisa = Operacao Then
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        
        If PerfilAcesso = BackOffice Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
        Else
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
        End If
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
        Call fgAppendNode(xmlFiltros, "Grupo_LayoutEntrada", "LayoutEntrada", "122")
            
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bruta)
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoContraparte", "")
        
        If optTipoTransf(enumTipoTransferencia.BookTransfer).value Then
            Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Interno)
        
        ElseIf optTipoTransf(enumTipoTransferencia.Mercado).value Then
            Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo)
        
        ElseIf optTipoTransf(enumTipoTransferencia.Todos).value Then
            Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Interno)
            Call fgAppendNode(xmlFiltros, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo)
        
        End If
        
    ElseIf intTipoPesquisa = MENSAGEM Then
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        
        If PerfilAcesso = BackOffice Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        Else
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackoffice)
        End If
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        
        If optTipoTransf(enumTipoTransferencia.BookTransfer).value Then
            Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0007")
        
        ElseIf optTipoTransf(enumTipoTransferencia.Mercado).value Then
            If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")
            
            ElseIf optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0005R2")
            
            ElseIf optNaturezaMovimento(enumNaturezaMovimento.Todos).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0005R2")
            
            End If
        
        ElseIf optTipoTransf(enumTipoTransferencia.Todos).value Then
            If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0007")
        
            ElseIf optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0005R2")
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0007")
        
            ElseIf optNaturezaMovimento(enumNaturezaMovimento.Todos).value Then
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0005R2")
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0007")
            
            End If
        
        End If
    End If
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_NaturezaMovimento", "")
    
    If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
        Call fgAppendNode(xmlFiltros, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Debito)
    
    ElseIf optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value Then
        Call fgAppendNode(xmlFiltros, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Credito)
    
    ElseIf optNaturezaMovimento(enumNaturezaMovimento.Todos).value Then
        Call fgAppendNode(xmlFiltros, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Debito)
        Call fgAppendNode(xmlFiltros, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Credito)
    
    End If
    
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
Dim intIgnoraGradeHorario                   As Integer

On Error GoTo ErrorHandler

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    For Each objListItemMesg In Me.lvwMensagem.ListItems
        With objListItemMesg
            If .Checked Then

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
                                                "CO_MESG_SPB", _
                                                Split(.Key, "|")(KEY_MSG_CO_MESG_SPB))
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "IN_OPER_DEBT_CRED", _
                                                Split(.Key, "|")(KEY_MSG_IN_OPER_DEBT_CRED))
                
                If .Tag <> vbNullString Then
                    Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                    "NU_SEQU_OPER_ATIV", _
                                                    Split(.Tag, "|")(TAG_MSG_NU_SEQU_OPER_ATIV))
                Else
                    For Each objListItemOper In Me.lvwOperacao.ListItems
                        With objListItemOper
                            If .Checked Then
                                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                                "NU_SEQU_OPER_ATIV", _
                                                                Split(.Key, "|")(KEY_OP_NU_SEQU_OPER_ATIV))
                                Exit For
                            End If
                        End With
                    Next
                
                End If
                    
'                intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)
                intIgnoraGradeHorario = 1
                
                Call fgAppendNode(xmlItemEnvio, "Grupo_Envio", _
                                                "IgnoraGradeHorario", _
                                                intIgnoraGradeHorario)

                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvio.xml)
                
                Set xmlItemEnvio = Nothing
                
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
    
    If Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    fgCursor True
    
    strDocFiltros = flMontarXMLFiltroPesquisa(Operacao)
    Call flCarregarListaOperacoes(strDocFiltros)
    
    strDocFiltros = flMontarXMLFiltroPesquisa(MENSAGEM)
    Call flCarregarListaMensagens(strDocFiltros)
    
    blnAcaoContingencia = False
    
    With tlbComandos
        .Buttons("concordancia").Enabled = True
        .Buttons("pagamento").Enabled = True
        .Buttons("pagamentocontingencia").Enabled = False
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
        strXMLRetorno = objOperacaoMensagem.ProcessarLoteLiquidacaoBrutaCBLC(intAcaoProcessamento, _
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

'Selecionar itens correspondentes de operações e mensagens já conciliados anteriormente
Private Sub flSelecionarItemLvwCorrespondente(ByVal intTipoPesquisa As enumTipoPesquisa, ByVal objListItem As ListItem)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim xmlLeitura                              As MSXML2.DOMDocument40

Dim lngNU_SEQU_OPER_ATIV                    As Long
Dim strNU_CTRL_IF                           As String
Dim strDH_REGT_MESG_SPB                     As String

Dim strItemKeyAuxOper                       As String
Dim strItemKeyAuxMesg                       As String

Dim blnErro                                 As Boolean
Dim objListView                             As ListView
Dim intSubItens                             As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
    
On Error GoTo ErrorHandler
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    blnErro = False
    
    With objListItem
        If .Tag <> vbNullString Then
            If intTipoPesquisa = MENSAGEM Then
                lvwOperacao.ListItems(.Tag).Checked = .Checked
                lvwOperacao.ListItems(.Tag).EnsureVisible
            Else
                lvwMensagem.ListItems(.Tag).Checked = .Checked
                lvwMensagem.ListItems(.Tag).EnsureVisible
            End If
        
        Else
            If intTipoPesquisa = MENSAGEM Then
                strNU_CTRL_IF = Split(.Key, "|")(KEY_MSG_NU_CTRL_IF)
                strDH_REGT_MESG_SPB = Split(.Key, "|")(KEY_MSG_DH_REGT_MESG_SPB)
            Else
                lngNU_SEQU_OPER_ATIV = Split(.Key, "|")(KEY_OP_NU_SEQU_OPER_ATIV)
            End If
            
            Call xmlLeitura.loadXML(objOperacao.ObterConciliacao(lngNU_SEQU_OPER_ATIV, _
                                                                 strNU_CTRL_IF, _
                                                                 strDH_REGT_MESG_SPB, _
                                                                 vbNullString, _
                                                                 0, _
                                                                 0, _
                                                                 0, _
                                                                 vntCodErro, _
                                                                 vntMensagemErro))
            
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If
            
            If xmlLeitura.xml <> vbNullString Then
                strItemKeyAuxOper = "|" & xmlLeitura.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                
                strItemKeyAuxMesg = "|" & xmlLeitura.selectSingleNode("//NU_CTRL_IF").Text & _
                                    "|" & xmlLeitura.selectSingleNode("//DH_REGT_MESG_SPB").Text & _
                                    "|" & xmlLeitura.selectSingleNode("//NU_SEQU_CNTR_REPE").Text & _
                                    "|" & xmlLeitura.selectSingleNode("//CO_MESG_SPB").Text & _
                                    "|" & xmlLeitura.selectSingleNode("//IN_OPER_DEBT_CRED").Text
                
                If intTipoPesquisa = MENSAGEM Then
                    If fgExisteItemLvw(Me.lvwOperacao, strItemKeyAuxOper) Then
                        .Tag = strItemKeyAuxOper
                        lvwOperacao.ListItems(strItemKeyAuxOper).Tag = strItemKeyAuxMesg
                        lvwOperacao.ListItems(strItemKeyAuxOper).Checked = .Checked
                        lvwOperacao.ListItems(strItemKeyAuxOper).EnsureVisible
                    Else
                        blnErro = True
                    End If
                
                Else
                    If fgExisteItemLvw(Me.lvwMensagem, strItemKeyAuxMesg) Then
                        .Tag = strItemKeyAuxMesg
                        lvwMensagem.ListItems(strItemKeyAuxMesg).Tag = strItemKeyAuxOper
                        lvwMensagem.ListItems(strItemKeyAuxMesg).Checked = .Checked
                        lvwMensagem.ListItems(strItemKeyAuxMesg).EnsureVisible
                    Else
                        blnErro = True
                    End If
                End If
            Else
                blnErro = True
            End If
        End If
    
        Call flConfigurarBotoesPorFuncionalidade(Not blnErro, objListItem, intTipoPesquisa)
        
    End With
    
    Set objOperacao = Nothing
    Set xmlLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set objOperacao = Nothing
    Set xmlLeitura = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSelecionarItemLvwCorrespondente", 0

End Sub

'Valida a seleção dos itens na tela, para posterior processamento
Private Function flValidarItensProcessamento(ByVal intAcao As enumAcaoConciliacao) As String

Dim objListItem                             As MSComctlLib.ListItem

Dim strOperValor                            As String
Dim strOperTransf                           As String
Dim strOperDC                               As String
Dim strOperKey                              As String

Dim strMesgValor                            As String
Dim strMesgTransf                           As String
Dim strMesgDC                               As String
Dim strMesgKey                              As String

    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flValidarItensProcessamento = "Selecione um item da lista de mensagens, antes de prosseguir com a operação desejada."
        Exit Function
    End If

    If PerfilAcesso = BackOffice Then
        For Each objListItem In Me.lvwOperacao.ListItems
            If objListItem.Checked Then
                strOperValor = objListItem.SubItems(COL_OP_VALOR)
                strOperTransf = objListItem.SubItems(COL_OP_TIPO_TRANSF)
                strOperDC = objListItem.SubItems(COL_OP_DEB_CRED)
                strOperKey = objListItem.Key
                Exit For
            End If
        Next
    
        For Each objListItem In Me.lvwMensagem.ListItems
            If objListItem.Checked Then
                strMesgValor = objListItem.SubItems(COL_MSG_VALOR)
                strMesgTransf = objListItem.SubItems(COL_MSG_TIPO_TRANSF)
                strMesgDC = objListItem.SubItems(COL_MSG_DEB_CRED)
                strMesgKey = objListItem.Key
                Exit For
            End If
        Next
        
        If strOperTransf <> strMesgTransf Then
            flValidarItensProcessamento = "Tipo de Transferência da Operação é diferente da Mensagem. Operação não permitida."
            Exit Function
        End If
        
        If strOperDC <> strMesgDC Then
            flValidarItensProcessamento = "Natureza de Movimento da Operação é diferente da Mensagem. Operação não permitida."
            Exit Function
        End If
    
        If strOperValor <> strMesgValor Then
            flValidarItensProcessamento = "Valor de Operação e Mensagem diferentes. Operação não permitida."
            Exit Function
        End If
        
        lvwMensagem.ListItems(strMesgKey).Tag = strOperKey

    Else
    
    
        'KIDA - CBLC - 10/10/2008
        'If intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento Then
            
            For Each objListItem In Me.lvwOperacao.ListItems
                If objListItem.Checked Then
                    strOperDC = objListItem.SubItems(COL_OP_DEB_CRED)
                    strOperTransf = objListItem.SubItems(COL_OP_TIPO_TRANSF)
    
                    'If (strOperDC <> "Débito" Or strOperTransf <> "Mercado") And intAcao <> AdmGeralEnviarConcordancia Then
                    '    flValidarItensProcessamento = "Pagamentos, Pagamentos em Contingência e Regularizações só são permitidas para operações Mercado / Débito."
                    '    Exit Function
                    'End If
                    
                    If strOperDC <> "Débito" And intAcao <> AdmGeralEnviarConcordancia Then
                        flValidarItensProcessamento = "Pagamentos, Pagamentos em Contingência e Regularizações só são permitidas para operações de Débito."
                        Exit Function
                    End If
                    
    
                End If
            Next
            
            For Each objListItem In Me.lvwMensagem.ListItems
                If objListItem.Checked Then
                    strMesgDC = objListItem.SubItems(COL_MSG_DEB_CRED)
                    strMesgTransf = objListItem.SubItems(COL_MSG_TIPO_TRANSF)
    
                    'If (strMesgDC <> "Débito" Or strMesgTransf <> "Mercado") And intAcao <> AdmGeralEnviarConcordancia Then
                    '    flValidarItensProcessamento = "Pagamentos, Pagamentos em Contingência e Regularizações só são permitidas para mensagens Mercado / Débito."
                    '    Exit Function
                    'End If
    
                    If strMesgDC <> "Débito" And intAcao <> AdmGeralEnviarConcordancia Then
                        flValidarItensProcessamento = "Pagamentos, Pagamentos em Contingência e Regularizações só são permitidas para mensagens de Débito."
                        Exit Function
                    End If
    
                End If
            Next
        'End If
        
        If intAcao = AdmGeralPagamentoContingencia Then
            For Each objListItem In Me.lvwOperacao.ListItems
                If objListItem.Checked Then
                    strOperValor = objListItem.SubItems(COL_OP_VALOR)
                    strOperTransf = objListItem.SubItems(COL_OP_TIPO_TRANSF)
                    strOperDC = objListItem.SubItems(COL_OP_DEB_CRED)
                    strOperKey = objListItem.Key
                    Exit For
                End If
            Next
        
            If strOperKey = vbNullString Then Exit Function
            
            For Each objListItem In Me.lvwMensagem.ListItems
                If objListItem.Checked Then
                    strMesgValor = objListItem.SubItems(COL_MSG_VALOR)
                    strMesgTransf = objListItem.SubItems(COL_MSG_TIPO_TRANSF)
                    strMesgDC = objListItem.SubItems(COL_MSG_DEB_CRED)
                    strMesgKey = objListItem.Key
                    Exit For
                End If
            Next
            
            If strOperTransf <> strMesgTransf Then
                flValidarItensProcessamento = "Tipo de Transferência da Operação é diferente da Mensagem. Operação não permitida."
                Exit Function
            End If
            
            If strOperDC <> strMesgDC Then
                flValidarItensProcessamento = "Natureza de Movimento da Operação é diferente da Mensagem. Operação não permitida."
                Exit Function
            End If
        
            If strOperValor = strMesgValor Then
                flValidarItensProcessamento = "Diferença entre Valores de Operação e Mensagem é igual a Zero (0). Pagamento em Contingência não permitido."
                Exit Function
            End If
            
            lvwMensagem.ListItems(strMesgKey).Tag = strOperKey
        
        End If
    End If
    
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
            Me.Caption = "CBLC - Liquidação Bruta (Backoffice)"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "CBLC - Liquidação Bruta (Administrador de Área)"
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
    Call flInicializarLvwMensagem
    Call flInicializarLvwOperacao
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
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If

        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width
        
        .lvwOperacao.Top = .fraNaturezaMovto.Top + .fraNaturezaMovto.Height + 240
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwOperacao.Top, .Height - .lvwOperacao.Top - 720)
        .lvwOperacao.Width = .Width - 240

        .lvwMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = .Height - .lvwMensagem.Top - 720
        .lvwMensagem.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set frmLiquidacaoBruta = Nothing
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

        .lvwOperacao.Height = .imgDummyH.Top - .lvwOperacao.Top
        .lvwMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwMensagem.Height = .Height - .lvwMensagem.Top - 720
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
    
On Error GoTo ErrorHandler
    
    Item.Selected = True
    
    If PerfilAcesso = BackOffice Then
        If Item.Checked Then
            For Each objListItem In lvwMensagem.ListItems
                If objListItem.Key <> Item.Key Then
                    objListItem.Checked = False
                    objListItem.Selected = False
                End If
            Next
        End If
    Else
        Call flSelecionarItemLvwCorrespondente(enumTipoPesquisa.MENSAGEM, Item)
    End If
    
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

Private Sub lvwOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim objListItem                             As ListItem
    
On Error GoTo ErrorHandler
    
    Item.Selected = True
    
    If PerfilAcesso = BackOffice Then
        If Item.Checked Then
            For Each objListItem In lvwOperacao.ListItems
                If objListItem.Key <> Item.Key Then
                    objListItem.Checked = False
                    objListItem.Selected = False
                End If
            Next
        End If
    Else
        Call flSelecionarItemLvwCorrespondente(enumTipoPesquisa.Operacao, Item)
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_ItemCheck", Me.Caption

End Sub

Private Sub optNaturezaMovimento_Click(Index As Integer)

On Error GoTo ErrorHandler

    Call flLimparListas
    If cboEmpresa.Text <> vbNullString Then
        Call flPesquisar
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optNaturezaMovimento_Click", Me.Caption

End Sub

Private Sub optTipoTransf_Click(Index As Integer)

On Error GoTo ErrorHandler
    
    Call flLimparListas
    If cboEmpresa.Text <> vbNullString Then
        Call flPesquisar
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - optTipoTransf_Click", Me.Caption

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
            If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                intAcaoProcessamento = enumAcaoConciliacao.BOConcordar
            Else
                intAcaoProcessamento = enumAcaoConciliacao.AdmGeralEnviarConcordancia
            End If
            
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

