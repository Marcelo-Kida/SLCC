VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompromissadaGenerica 
   Caption         =   "Liquidação Compromissada Genérica"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
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
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   7650
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
         NumButtons      =   7
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
            Caption         =   "Concordar    "
            Key             =   "concordancia"
            Object.ToolTipText     =   "Concodar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Liberar         "
            Key             =   "liberacao"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Rejeitar        "
            Key             =   "retorno"
            Object.ToolTipText     =   "Retornar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair              "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   2685
      Left            =   60
      TabIndex        =   2
      Top             =   4140
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   4736
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
            Picture         =   "frmCompromissadaGenerica.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompromissadaGenerica.frx":16D8
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
Attribute VB_Name = "frmCompromissadaGenerica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pela liquidação de Operações Compromissada Genérica,
'' através da camada de controle de caso de uso MIU.
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_VEIC_LEGA             As Integer = 0
Private Const COL_MSG_DATA_MOVI             As Integer = 1
Private Const COL_MSG_NATU_OPER             As Integer = 2
Private Const COL_MSG_DATA_LIQU             As Integer = 3
Private Const COL_MSG_DATA_RETN             As Integer = 4
Private Const COL_MSG_VALR_OPER             As Integer = 5
Private Const COL_MSG_VALR_MESG             As Integer = 6
Private Const COL_MSG_DIFERENCA             As Integer = 7
Private Const COL_MSG_STATUS                As Integer = 8

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const KEY_MSG_CO_VEIC_LEGA          As Integer = 1
Private Const KEY_MSG_DT_OPER_ATIV          As Integer = 2
Private Const KEY_MSG_DT_LIQU_OPER_ATIV     As Integer = 3
Private Const KEY_MSG_DT_OPER_ATIV_RETN     As Integer = 4
Private Const KEY_MSG_CO_ULTI_SITU_PROC     As Integer = 5

'Constantes de posicionamento de campos na propriedade Tag do item do ListView de Mensagens
Private Const TAG_MSG_NU_CTRL_IF            As Integer = 1
Private Const TAG_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const TAG_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const TAG_MSG_NU_CTRL_CAMR          As Integer = 4
Private Const TAG_MSG_DH_ULTI_ATLZ          As Integer = 5

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_NUMERO_COMANDO         As Integer = 0
Private Const COL_OP_DC                     As Integer = 1
Private Const COL_OP_VALOR                  As Integer = 2
Private Const COL_OP_DATA_RETORNO           As Integer = 3
Private Const COL_OP_CODIGO                 As Integer = 4

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Operações
Private Const KEY_OP_NU_SEQU_OPER_ATIV      As Integer = 1

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmCompromissadaGenerica"

'Constantes de erros de negócio específicos
Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3095
'------------------------------------------------------------------------------------------
'Fim declaração constantes

Private Enum enumTipoPesquisa
    Operacao = 0
    MENSAGEM = 1
End Enum

Private intAcaoProcessamento                As enumAcaoConciliacao

Private lngPerfil                           As Long
Private blnDummyH                           As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Calcular a diferença dos valores das operações e mensagens SPB.
Private Sub flCalcularDiferencasListView()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

On Error GoTo ErrorHandler
    
    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VALR_OPER)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VALR_MESG)))
            
            .SubItems(COL_MSG_DIFERENCA) = fgVlrXml_To_Interface(dblValorMensagem - dblValorOperacao)
            
            If dblValorMensagem - dblValorOperacao <> 0 Then
                .ListSubItems(COL_MSG_DIFERENCA).ForeColor = vbRed
            Else
                If PerfilAcesso = enumPerfilAcesso.BackOffice Then
                    objListItem.Checked = True
                End If
            End If
        End With
    Next

Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListView", 0)

End Sub

'' Carrega as operações existentes e preencher a interface com os mesmos
Private Sub flCarregarListaDetalheOperacoes()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strItemKey                              As String

On Error GoTo ErrorHandler
    
    If lvwMensagem.SelectedItem Is Nothing Then Exit Sub
    
    strItemKey = lvwMensagem.SelectedItem.Key
    lvwOperacao.ListItems.Clear

    For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))

        With lvwOperacao.ListItems.Add(, "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            .Text = objDomNode.selectSingleNode("NU_COMD_OPER").Text
            .SubItems(COL_OP_DC) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
            .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
            .SubItems(COL_OP_CODIGO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            
            If objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text <> gstrDataVazia Then
                .SubItems(COL_OP_DATA_RETORNO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text)
            End If
            
        End With

    Next

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
Exit Sub
ErrorHandler:
   
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaDetalheOperacoes", 0)

End Sub

'' Carrega o NET das mensagens SPB e preencher a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método A8MIU.clsMensagem.ObterDetalheMensagem

Private Sub flCarregarListaNetMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem        As A8MIU.clsMensagem
#End If

Dim strRetLeitura          As String
Dim xmlRetLeitura          As MSXML2.DOMDocument40
Dim objDomNode             As MSXML2.IXMLDOMNode
Dim strListItemKey         As String
Dim strListItemTag         As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(pstrFiltro, vntCodErro, vntMensagemErro)
    
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
            strListItemKey = flMontarChaveItemListview(objDomNode, enumTipoPesquisa.MENSAGEM)
                        
            strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("NU_CTRL_CAMR").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
                             
            If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                With lvwMensagem.ListItems.Add(, strListItemKey)
                    .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_MSG_DATA_MOVI) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_OPER").Text)
                    .SubItems(COL_MSG_DATA_LIQU) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_LIQU").Text)
                    .SubItems(COL_MSG_DATA_RETN) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text)
                    .SubItems(COL_MSG_VALR_OPER) = " "
                    .SubItems(COL_MSG_VALR_MESG) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                End With
            Else
                With lvwMensagem.ListItems(strListItemKey)
                    .SubItems(COL_MSG_VALR_MESG) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
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

'' Carrega o NET das operações compromissadas e preencher a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método A8MIU.clsOperacao.ObterDetalheOperacao

Private Sub flCarregarListaNetOperacoes(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objOperacao        As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao        As A8MIU.clsOperacao
#End If

Dim strRetLeitura          As String
Dim xmlRetLeitura          As MSXML2.DOMDocument40
Dim objDomNode             As MSXML2.IXMLDOMNode
Dim strListItemKey         As String
Dim dblValorOperacao       As Double
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(pstrFiltro, vntCodErro, vntMensagemErro)
    
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
            
            strListItemKey = flMontarChaveItemListview(objDomNode, enumTipoPesquisa.Operacao)
                    
            If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                dblValorOperacao = flValorOperacoes(strListItemKey)
            
                With lvwMensagem.ListItems.Add(, strListItemKey)
                    
                    .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_MSG_DATA_MOVI) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                    .SubItems(COL_MSG_NATU_OPER) = IIf(dblValorOperacao >= 0, "Venda", "Compra")
                    .SubItems(COL_MSG_DATA_LIQU) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
                    .SubItems(COL_MSG_DATA_RETN) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text)
                    .SubItems(COL_MSG_VALR_OPER) = fgVlrXml_To_Interface(Abs(dblValorOperacao))
                    .SubItems(COL_MSG_VALR_MESG) = " "
                    .SubItems(COL_MSG_STATUS) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    
                    If PerfilAcesso = enumPerfilAcesso.AdmArea Then
                        If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConcordanciaBackofficeBMA0013 Then
                            .ListSubItems(COL_MSG_STATUS).ForeColor = vbBlue
                        End If
                    End If
            
                End With
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

'Configurar os botões da tela conforme o perfil do usuário
'Libera o botão de Concordância somente para o perfil Back Office
'Libera os botões de Liberação e Retorno somente para o perfil Administrador da Área

Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)
    
On Error GoTo ErrorHandler
    
    With tlbComandos
        .Buttons("concordancia").Visible = (PerfilAcesso = enumPerfilAcesso.BackOffice)
        .Buttons("liberacao").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("retorno").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Refresh
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorPerfil", 0

End Sub

'' Carrega as propriedades necessárias a interface frmCompromissadaGenerica, através da
'' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializarFormulario()

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

'Configurar o list view de mensagens
Private Sub flInicializarLvwMensagem()

On Error GoTo ErrorHandler
    
    With Me.lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Veículo Legal", 2900
        .Add , , "Data Operação", 1400
        .Add , , "Natureza Operação", 1800
        .Add , , "Data Liquidação", 1400
        .Add , , "Data Retorno", 1400
        .Add , , "Valor Operações", 1400, lvwColumnRight
        .Add , , "Valor Mensagem", 1400, lvwColumnRight
        .Add , , "Diferença", 1000, lvwColumnRight
        .Add , , "Status", 2100
    End With
    
Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwMensagem", 0

End Sub

'Configurar o list view de operações
Private Sub flInicializarLvwOperacao()

On Error GoTo ErrorHandler
    
    With Me.lvwOperacao.ColumnHeaders
        .Clear
        .Add , , "Número Comando", 2900
        .Add , , "D/C", 1400
        .Add , , "Valor Financeiro", 1800, lvwColumnRight
        .Add , , "Data Retorno", 1400
        .Add , , "Código Operação", 5200
    End With
    
Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwOperacao", 0

End Sub

'Limpar os list view de mensagem e de operação
Private Sub flLimparListas()
    
    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear

End Sub

'Marcar as mensagens SPB que deveriam ser enviadas mas retornaram rejeitadas por grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim intCont                                 As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            For Each objListItem In lvwMensagem.ListItems
                With objListItem
                    If Split(.Key, "|")(KEY_MSG_CO_VEIC_LEGA) = objDomNode.selectSingleNode("CO_LOCA_LIQU").Text And _
                       Split(.Key, "|")(KEY_MSG_DT_OPER_ATIV) = objDomNode.selectSingleNode("DT_OPER_ATIV").Text And _
                       Split(.Key, "|")(KEY_MSG_DT_LIQU_OPER_ATIV) = objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text And _
                       Split(.Key, "|")(KEY_MSG_DT_OPER_ATIV_RETN) = objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text And _
                       Split(.Key, "|")(KEY_MSG_CO_ULTI_SITU_PROC) = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text Then
                        
                        For intCont = 1 To .ListSubItems.Count
                            .ListSubItems(intCont).ForeColor = vbRed
                        Next
                        
                        .Text = "Horário Excedido"
                        .ToolTipText = "Horário limite p/envio da mensagem excedido"
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

'Montar a chave dos item de mensagem e de operações nos list view
Private Function flMontarChaveItemListview(ByVal objDomNode As MSXML2.IXMLDOMNode, _
                                           ByVal intTipo As enumTipoPesquisa)
                
Dim strListItemKey                          As String
Dim intStatusConciliacao                    As Integer

On Error GoTo ErrorHandler
    
    If intTipo = enumTipoPesquisa.MENSAGEM Then
        Select Case Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text)
            Case enumStatusMensagem.ConcordanciaBackoffice
                intStatusConciliacao = enumStatusOperacao.ConcordanciaBackofficeBMA0013
            Case enumStatusMensagem.DiscordanciaBackoffice
                intStatusConciliacao = enumStatusOperacao.DiscordanciaBackoffice
            Case enumStatusMensagem.AConciliar
                intStatusConciliacao = enumStatusOperacao.AConciliarBMA0013
        End Select
    
        strListItemKey = "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                         "|" & objDomNode.selectSingleNode("DT_OPER").Text & _
                         "|" & objDomNode.selectSingleNode("DT_LIQU").Text & _
                         "|" & objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text & _
                         "|" & intStatusConciliacao
    
    Else
        intStatusConciliacao = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
    
        strListItemKey = "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                         "|" & objDomNode.selectSingleNode("DT_OPER_ATIV").Text & _
                         "|" & objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text & _
                         "|" & objDomNode.selectSingleNode("DT_OPER_ATIV_RETN").Text & _
                         "|" & intStatusConciliacao
    
    End If
    
    flMontarChaveItemListview = strListItemKey
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Montar uma expressão xpath para seleção das operações
Private Function flMontarCondicaoNavegacaoXMLOperacoes(ByVal strItemKey As String)
                
Dim strCondicao                             As String
    
On Error GoTo ErrorHandler
    
    strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_MSG_CO_VEIC_LEGA) & "' " & _
                                                          " and DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DT_OPER_ATIV) & "' " & _
                                                          " and DT_LIQU_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DT_LIQU_OPER_ATIV) & "' " & _
                                                          " and DT_OPER_ATIV_RETN='" & Split(strItemKey, "|")(KEY_MSG_DT_OPER_ATIV_RETN) & "' " & _
                                                          " and CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "']"
    
    flMontarCondicaoNavegacaoXMLOperacoes = strCondicao
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoes", 0

End Function

'Montar uma expressão xpath para seleção do net das operações
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String)
                
Dim strDebito                               As String
Dim strCredito                              As String
    
On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_MSG_CO_VEIC_LEGA) & "' " & _
                                     " and ../DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DT_OPER_ATIV) & "' " & _
                                     " and ../DT_LIQU_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DT_LIQU_OPER_ATIV) & "' " & _
                                     " and ../DT_OPER_ATIV_RETN='" & Split(strItemKey, "|")(KEY_MSG_DT_OPER_ATIV_RETN) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' "
    
    strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_MSG_CO_VEIC_LEGA) & "' " & _
                                     " and ../DT_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DT_OPER_ATIV) & "' " & _
                                     " and ../DT_LIQU_OPER_ATIV='" & Split(strItemKey, "|")(KEY_MSG_DT_LIQU_OPER_ATIV) & "' " & _
                                     " and ../DT_OPER_ATIV_RETN='" & Split(strItemKey, "|")(KEY_MSG_DT_OPER_ATIV_RETN) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_MSG_CO_ULTI_SITU_PROC) & "' "
    
    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "
    
    flMontarExpressaoCalculoNetOperacoes = strDebito & strCredito
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoCalculoNetOperacoes", 0

End Function

'Montar o xml de filtro para pesquisa das operações
Private Function flMontarXMLFiltroPesquisa(ByVal intTipoPesquisa As enumTipoPesquisa) As String
    
Dim xmlFiltros                              As MSXML2.DOMDocument40
    
On Error GoTo ErrorHandler
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))
        
    If intTipoPesquisa = Operacao Then
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", "95")
        Call fgAppendNode(xmlFiltros, "Grupo_TipoOperacao", "TipoOperacao", "96")
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliarBMA0013)
        
        If PerfilAcesso = AdmArea Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeBMA0013)
        End If
    
    Else
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        
        If PerfilAcesso = AdmArea Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackoffice)
        End If
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "BMA0013")
        
    End If
    
    flMontarXMLFiltroPesquisa = xmlFiltros.xml
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Montar o xml para o processamento das informações apresentadas na interface
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
                                                   "CO_VEIC_LEGA", _
                                                   Split(.Key, "|")(KEY_MSG_CO_VEIC_LEGA))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DT_OPER_ATIV", _
                                                   Split(.Key, "|")(KEY_MSG_DT_OPER_ATIV))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DT_LIQU_OPER_ATIV", _
                                                   Split(.Key, "|")(KEY_MSG_DT_LIQU_OPER_ATIV))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DT_OPER_ATIV_RETN", _
                                                   Split(.Key, "|")(KEY_MSG_DT_OPER_ATIV_RETN))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_ULTI_SITU_PROC", _
                                                   Split(.Key, "|")(KEY_MSG_CO_ULTI_SITU_PROC))
                
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
                                                   "NU_CTRL_CAMR", _
                                                   Split(.Tag, "|")(TAG_MSG_NU_CTRL_CAMR))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DH_ULTI_ATLZ", _
                                                   Split(.Tag, "|")(TAG_MSG_DH_ULTI_ATLZ))

                intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "IgnoraGradeHorario", _
                                                   intIgnoraGradeHorario)

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

'Montar o resultado do processamento
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

'Controlar as chamadas das funcionalidades de pesquisa

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
    
    strDocFiltros = flMontarXMLFiltroPesquisa(Operacao)
    Call flCarregarListaNetOperacoes(strDocFiltros)
    
    strDocFiltros = flMontarXMLFiltroPesquisa(MENSAGEM)
    Call flCarregarListaNetMensagens(strDocFiltros)
    
    Call flCalcularDiferencasListView
    
    If lvwMensagem.ListItems.Count > 0 Then
        lvwMensagem.ListItems(1).Selected = True
        Call lvwMensagem_ItemClick(lvwMensagem.ListItems(1))
    End If

    fgCursor
    
Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPesquisar", 0)

End Sub

'' Executar o processamento efetuado através da camada controladora de casos de uso
'' MIU, método A8MIU.clsOperacaoMensagem.ProcessarCompromissadaGenerica
Private Function flProcessar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem     As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem     As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLRetorno               As String
Dim strXMLProc                  As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    strXMLProc = flMontarXMLProcessamento
    
    If strXMLProc <> vbNullString Then
        fgCursor True
        
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.ProcessarCompromissadaGenerica(intAcaoProcessamento, strXMLProc, vntCodErro, vntMensagemErro)
        
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

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString
Private Function flValidarItensProcessamento() As String

Dim objListItem                             As MSComctlLib.ListItem
Dim dblValorConsist                         As Double
Dim intStatus                               As Integer

    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flValidarItensProcessamento = "Selecione pelo menos um item da lista, antes de prosseguir com a operação desejada."
        Exit Function
    End If

    For Each objListItem In Me.lvwMensagem.ListItems
        If objListItem.Checked Then
            If PerfilAcesso = BackOffice Then
                dblValorConsist = fgVlrXml_To_Decimal(fgVlr_To_Xml(objListItem.SubItems(COL_MSG_DIFERENCA)))
                
                If dblValorConsist <> 0 Then
                    flValidarItensProcessamento = "Um ou mais itens estão com divergência de valores. Solicitação de Concordância não permitida."
                    Exit Function
                End If
            
            Else
                If intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar Then
                    If objListItem.SubItems(COL_MSG_VALR_OPER) <> vbNullString Then
                        intStatus = Val(Split(objListItem.Key, "|")(KEY_MSG_CO_ULTI_SITU_PROC))
                    Else
                        intStatus = 0
                    End If
                    
                    If intStatus <> enumStatusOperacao.ConcordanciaBackofficeBMA0013 Then
                        flValidarItensProcessamento = "Só é possível 'Liberar' itens de mensagem nos seguintes status:" & vbNewLine & vbNewLine & _
                                                      "- Concordância BO BMA0013"
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
End Function

' Retorna o NET das operações selecionadas
Private Function flValorOperacoes(ByVal strItemKey As String)
    
Dim strExpression                   As String
Dim vntValor                        As Variant
    
    vntValor = 0
    
    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlOperacoes, strExpression))
    
    flValorOperacoes = vntValor

End Function

Property Get PerfilAcesso() As enumPerfilAcesso
    
    PerfilAcesso = lngPerfil

End Property

Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
    
    lngPerfil = pPerfil
    
    Select Case pPerfil
        Case enumPerfilAcesso.BackOffice
            Me.Caption = "Conciliação - NET Compromissada Genérica (Backoffice)"
        Case enumPerfilAcesso.AdmArea
            Me.Caption = "Liberação - NET Compromissada Genérica  (Administrador de Área)"
    End Select
    
    Call flConfigurarBotoesPorPerfil(PerfilAcesso)
    Call flLimparListas
    
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

        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width

        .lvwMensagem.Top = .cboEmpresa.Top + .cboEmpresa.Height + 120
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwMensagem.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
        .lvwOperacao.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set frmCompromissadaGenerica = Nothing

End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    blnDummyH = True

End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not blnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If

    Me.imgDummyH.Top = Y + imgDummyH.Top

    On Error Resume Next

    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
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
    
    If Item.Checked Then
        If Trim$(Item.SubItems(COL_MSG_VALR_MESG)) = vbNullString Then
            frmMural.Display = "Seleção do item para conciliação não permitida. Valor de mensagem não encontrado."
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            Item.Checked = False
        ElseIf Trim$(Item.SubItems(COL_MSG_VALR_OPER)) = vbNullString Then
            frmMural.Display = "Seleção do item para conciliação não permitida. Valor de operação não encontrado."
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            Item.Checked = False
        End If
    End If
        
    Item.Selected = True
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemCheck", Me.Caption

End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call flCarregarListaDetalheOperacoes
    
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
            intAcaoProcessamento = enumAcaoConciliacao.BOConcordar
            
        Case "liberacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar
            
        Case "retorno"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaRejeitar
            
        Case gstrSair
            Unload Me
            
    End Select
    
    If intAcaoProcessamento <> 0 Then
        strValidaProcessamento = flValidarItensProcessamento
        If strValidaProcessamento <> vbNullString Then
            frmMural.Display = strValidaProcessamento
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            GoTo ExitSub
        End If
        
        If intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar Then
            If MsgBox("Confirma a liberação do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
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
