VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmIntegrarCCOnLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conta Corrente - Integração Conta Corrente"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   14100
   WindowState     =   2  'Maximized
   Begin VB.Frame fraJustificativa 
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   6075
      Begin VB.TextBox txtJustificativa 
         Height          =   285
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   3
         Top             =   180
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Justificativa"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8025
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   582
      ButtonWidth     =   2487
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar Filtro"
            Key             =   "AplicarFiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
            Key             =   "DefinirFiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Integrar C/C"
            Key             =   "integrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Operações"
            Key             =   "MostrarOperacao"
            Object.ToolTipText     =   "Mostrar Operações"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Mensagens"
            Key             =   "MostrarMensagem"
            Object.ToolTipText     =   "Mostrar Mensagens"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair            "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   9660
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
            Picture         =   "frmIntegrarCCOnLine.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegrarCCOnLine.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   12360
      Top             =   7110
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin MSComctlLib.ListView lvwContaCorrente 
      Height          =   6795
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   11986
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   28
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Empresa"
         Object.Width           =   5212
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sistema"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Operação"
         Object.Width           =   2357
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Número Comando"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Veiculo Legal"
         Object.Width           =   5980
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Situação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tipo Operação"
         Object.Width           =   5477
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Local Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Banco"
         Object.Width           =   5927
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Agência"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Número C/C"
         Object.Width           =   2194
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Valor Lançamento"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tipo Movto."
         Object.Width           =   1850
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Tipo Lançamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Sub-tipo Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Conta Contábil Débito"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Conta Contábil Crédito"
         Object.Width           =   3281
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Código Histórico Contábil"
         Object.Width           =   3544
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Descrição Histórico Contábil"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Codigo Veiculo Legal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Codigo Situação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Codigo Tipo Operação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Codigo Banco"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Codigo Local Liquidação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Condição Net Operacões"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Canal Venda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Sub-Produto"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Código Lote"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmIntegrarCCOnLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Possibilita a integração OnLine de Conta Corrente

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40
Private xmlRetLeitura                       As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmIntegrarCCOnLine"

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

'Constantes de erros de negócio específicos
Private Const COD_ERRO_NEGOCIO_GRADE        As Long = 3043
Private Const COD_ERRO_NEGOCIO_ESTORNO      As Long = 3044

'Constantes de Configuração de Colunas
Private Const COL_EMPRESA                   As Integer = 0
Private Const COL_SISTEMA                   As Integer = 1
Private Const COL_DATA_OPERACAO             As Integer = 2
Private Const COL_NUMERO_COMANDO            As Integer = 3
Private Const COL_VEICULO_LEGAL             As Integer = 4
Private Const COL_SITUACAO                  As Integer = 5
Private Const COL_TIPO_OPERACAO             As Integer = 6
Private Const COL_LOCA_LIQU                 As Integer = 7
Private Const COL_BANCO                     As Integer = 8
Private Const COL_AGENCIA                   As Integer = 9
Private Const COL_CONTA_CORRENTE            As Integer = 10
Private Const COL_VALOR                     As Integer = 11
Private Const COL_TIPO_MOVIMENTO            As Integer = 12
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const COL_TIPO_LANCAMENTO           As Integer = 13
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const COL_SUB_TIPO_ATIVO            As Integer = 14
Private Const COL_CONTA_CONTABIL_DEB        As Integer = 15
Private Const COL_CONTA_CONTABIL_CRED       As Integer = 16
Private Const COL_COD_HIST_CONTABIL         As Integer = 17
Private Const COL_DES_HIST_CONTABIL         As Integer = 18
Private Const COL_COD_VEIC_LEGA             As Integer = 19
Private Const COL_COD_SITUACAO              As Integer = 20
Private Const COL_COD_TIPO_OPER             As Integer = 21
Private Const COL_COD_BANCO                 As Integer = 22
Private Const COL_COD_LOCA_LIQU             As Integer = 23
Private Const COL_COND_NET_OPERACOES        As Integer = 24
Private Const COL_CANAL_VENDA               As Integer = 25
Private Const COL_PRODUTO                   As Integer = 26
Private Const COL_COD_LOTE                  As Integer = 27

'Constantes de posicionamento de campos na propriedade Key do item do ListView
Private Const POS_NU_SEQU_OPER_ATIV         As Integer = 1
Private Const POS_TP_LANC_ITGR              As Integer = 2
Private Const POS_DH_ULTI_ATLZ              As Integer = 3
Private Const POS_NR_SEQU_LANC              As Integer = 4

Private Const KEY_EMPRESA                   As Integer = 1
Private Const KEY_DATA_OPERACAO             As Integer = 2
Private Const KEY_TIPO_OPERACAO             As Integer = 3
Private Const KEY_VEICULO_LEGAL             As Integer = 4
Private Const KEY_LOCA_LIQU                 As Integer = 5
Private Const KEY_BANCO                     As Integer = 6
Private Const KEY_AGENCIA                   As Integer = 7
Private Const KEY_CONTA_CORRENTE            As Integer = 8
Private Const KEY_CO_ULTI_SITU_PROC         As Integer = 9
Private Const KEY_CD_LOTE                   As Integer = 10

Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strDocFiltros                       As String
Private blnPrimeiraConsulta                 As Boolean

Private lngIndexClassifList                 As Long

'Monta a lista de lançamentos passíveis de processamento
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsContaCorrente
#End If

Dim objListItem                             As MSComctlLib.ListItem

Dim strRetLeitura                           As String
Dim strListItemKey                          As String
Dim strListItemKey2                         As String
'Dim xmlRetLeitura                           As MSXML2.DOMDocument40
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSelecaoFiltro                        As String
Dim strDocFiltrosAux                        As String
Dim lngCont                                 As Long
Dim dblValorOperacao                        As Double
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim strSequenciaOperacao                    As String
Dim intDebitoCredito                        As Integer
Dim blnCorretoras                           As Boolean
Dim blnExisteNoGrid                         As Boolean

On Error GoTo ErrorHandler

    Call flLimparLista
    
    strSelecaoFiltro = enumStatusIntegracao.Disponível & ";" & _
                       enumStatusIntegracao.ErroCC & ";" & _
                       enumStatusIntegracao.ErroSaldoCC & ";" & _
                       enumStatusIntegracao.Antecipado & ";" & _
                       enumStatusIntegracao.ErroEstornoCC
        
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    If Me.tlbComandos.Buttons("AplicarFiltro").value = tbrPressed Then
        strDocFiltrosAux = strDocFiltros
    End If
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlFiltros.loadXML(strDocFiltrosAux) Then
        Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    End If
    
    fgAppendNode xmlFiltros, "Repeat_Filtros", "Grupo_Status", ""
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    '>>> -------------------------------------------------------------------------------------------
    
    fgCursor True
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
    strRetLeitura = objOperacao.ObterDetalheLancamento(xmlFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheLancamento/*")
            strSequenciaOperacao = IIf(objDomNode.selectSingleNode("OWNER").Text = "A8", _
                                       objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       "-" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            'Cesar 08/05/2007 - Conta Corrente Corretoras
            blnCorretoras = False
            blnExisteNoGrid = False
            If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                strListItemKey = "|" & objDomNode.selectSingleNode("CO_EMPR").Text & _
                                 "|" & objDomNode.selectSingleNode("DT_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("TP_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_BANC").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_AGEN").Text & _
                                 "|" & objDomNode.selectSingleNode("NU_CC").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & _
                                 "|" & objDomNode.selectSingleNode("CD_LOTE").Text
                
                dblValorOperacao = flNetOperacoes(strListItemKey)
                
                For Each objListItem In lvwContaCorrente.ListItems
                      strListItemKey2 = "|" & fgObterCodigoCombo(objListItem.Text) & _
                                        "|" & fgDt_To_Xml(objListItem.SubItems(COL_DATA_OPERACAO)) & _
                                        "|" & objListItem.SubItems(COL_COD_TIPO_OPER) & _
                                        "|" & objListItem.SubItems(COL_COD_VEIC_LEGA) & _
                                        "|" & objListItem.SubItems(COL_COD_LOCA_LIQU) & _
                                        "|" & objListItem.SubItems(COL_COD_BANCO) & _
                                        "|" & objListItem.SubItems(COL_AGENCIA) & _
                                        "|" & objListItem.SubItems(COL_CONTA_CORRENTE) & _
                                        "|" & objListItem.SubItems(COL_COD_SITUACAO) & _
                                        "|" & objListItem.SubItems(COL_COD_LOTE)
                    
                    If strListItemKey = strListItemKey2 Then
                        blnCorretoras = True
                        objListItem.SubItems(COL_COND_NET_OPERACOES) = objListItem.SubItems(COL_COND_NET_OPERACOES) & _
                                                                       "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                        blnExisteNoGrid = True
                        Exit For
                    End If
                Next
                
                If Not blnExisteNoGrid Then
                    If dblValorOperacao > 0 Then
                        intDebitoCredito = enumTipoDebitoCredito.Credito
                        
'                        If Trim(objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text) = "Crédito" Then
'                            blnCorretoras = False
'                        Else
'                            blnCorretoras = True
'                        End If
                        
                    Else
                        intDebitoCredito = enumTipoDebitoCredito.Debito
                        
'                        If objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text = "Débito" Then
'                            blnCorretoras = False
'                        Else
'                            blnCorretoras = True
'                        End If
                        
                    End If
                End If
            End If
            
            If blnCorretoras = False Then
                With lvwContaCorrente.ListItems.Add(, _
                        "|" & strSequenciaOperacao & _
                        "|" & objDomNode.selectSingleNode("TP_LANC_ITGR").Text & _
                        "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                        "|" & objDomNode.selectSingleNode("NR_SEQU_LANC").Text)
                    
                    'Empresa
                    If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                        'Obtem a descrição da Empresa via QUERY XML
                        .Text = _
                                objDomNode.selectSingleNode("CO_EMPR").Text & " - " & xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                                objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                    End If
                    
                    'Número do Comando
                    .SubItems(COL_SISTEMA) = objDomNode.selectSingleNode("SG_SIST").Text & " - " & objDomNode.selectSingleNode("NO_SIST").Text
                    
                    'Data Operação
                    If objDomNode.selectSingleNode("DT_OPER").Text <> gstrDataVazia Then
                        .SubItems(COL_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER").Text)
                    End If
                    
                    'Número do Comando
                    .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    
                    'Veiculo Legal
                    .SubItems(COL_VEICULO_LEGAL) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_COD_VEIC_LEGA) = objDomNode.selectSingleNode("CO_VEIC_LEGA").Text
                    
                    'Situação
                    .SubItems(COL_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_COD_SITUACAO) = objDomNode.selectSingleNode("CO_SITU_PROC").Text
                    
                    'Tipo de Operação
                    .SubItems(COL_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                    .SubItems(COL_COD_TIPO_OPER) = objDomNode.selectSingleNode("TP_OPER").Text
                    
                    'Local de Liquidação
                    If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                        
                        If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                               objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU") Is Nothing Then
                            
                            'Obtem a descrição do Local de Liquidação via QUERY XML
                            .SubItems(COL_LOCA_LIQU) = _
                                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                    objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU").Text
                    
                        Else
                            
                            vntCodErro = 5
                            vntMensagemErro = "Usuário não possui acesso ao Local de Liquidação " & _
                                              objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "."
                            GoTo ErrorHandler
                            
                        End If
                    
                    End If
                    .SubItems(COL_COD_LOCA_LIQU) = objDomNode.selectSingleNode("CO_LOCA_LIQU").Text
    
                    'Banco
                    .SubItems(COL_BANCO) = objDomNode.selectSingleNode("CO_BANC").Text
                    .SubItems(COL_COD_BANCO) = objDomNode.selectSingleNode("CO_BANC").Text
                                               
                    'Agência
                    .SubItems(COL_AGENCIA) = objDomNode.selectSingleNode("CO_AGEN").Text
                    
                    'Número C/C
                    .SubItems(COL_CONTA_CORRENTE) = objDomNode.selectSingleNode("NU_CC").Text
                    
                    If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                        'Valor do Lançamento
                        .SubItems(COL_VALOR) = fgVlrXml_To_Interface(fgVlr_To_Xml(Abs(dblValorOperacao)))
                        
                        'Tipo Movto.
                        .SubItems(COL_TIPO_MOVIMENTO) = IIf(intDebitoCredito = enumTipoDebitoCredito.Debito, "Débito", "Crédito")
                    Else
                        'Valor do Lançamento
                        .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_LANC_CC").Text)
                        
                        'Tipo Movto.
                        .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text
                    End If
                    
                    'Tipo de Lançamento
                    .SubItems(COL_TIPO_LANCAMENTO) = IIf(Val(objDomNode.selectSingleNode("TP_LANC_ITGR").Text) = enumTipoLancamentoIntegracao.Estorno, "Estorno", "Normal")
                    
                    'Sub-tipo Ativo
                    .SubItems(COL_SUB_TIPO_ATIVO) = objDomNode.selectSingleNode("CO_SUB_TIPO_ATIV").Text
                    
                    'Conta Contábil Débito
                    .SubItems(COL_CONTA_CONTABIL_DEB) = IIf(Val(objDomNode.selectSingleNode("CO_CNTA_DEBT").Text) = 0, _
                                                            vbNullString, _
                                                            objDomNode.selectSingleNode("CO_CNTA_DEBT").Text)
                    
                    'Conta Contábil Crédito
                    .SubItems(COL_CONTA_CONTABIL_CRED) = IIf(Val(objDomNode.selectSingleNode("CO_CNTA_CRED").Text) = 0, _
                                                             vbNullString, _
                                                             objDomNode.selectSingleNode("CO_CNTA_CRED").Text)
                    
                    'Código Histórico Contábil
                    .SubItems(COL_COD_HIST_CONTABIL) = objDomNode.selectSingleNode("CO_HIST_CNTA_CNTB").Text
                    
                    'Descriçao do Histórico contábil
                    .SubItems(COL_DES_HIST_CONTABIL) = objDomNode.selectSingleNode("DE_HIST_CNTA_CNTB").Text
                    
                    .SubItems(COL_COND_NET_OPERACOES) = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                    
                    'KIDA - SGC
                    .SubItems(COL_CANAL_VENDA) = fgDescricaoCanalVenda(objDomNode.selectSingleNode("TP_CNAL_VEND").Text)
                    
                    '.SubItems(COL_PRODUTO) = fgObterDescricaoProduto(objDomNode.selectSingleNode("CD_PROD").Text)
                    'Joabe - Correcao Erro Homologação
                    If Not objDomNode.selectSingleNode("CD_SUB_PROD") Is Nothing Then
                        .SubItems(COL_PRODUTO) = objDomNode.selectSingleNode("CD_SUB_PROD").Text
                    Else
                        .SubItems(COL_PRODUTO) = ""
                    End If
                
                    'Código do Lote
                    If Not objDomNode.selectSingleNode("CD_LOTE") Is Nothing Then
                        .SubItems(COL_COD_LOTE) = objDomNode.selectSingleNode("CD_LOTE").Text
                    Else
                        .SubItems(COL_COD_LOTE) = ""
                    End If
                End With
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lvwContaCorrente, lngIndexClassifList, True)
    
    Set xmlFiltros = Nothing
    Set xmlRetLeitura = Nothing

    fgCursor
    
    Exit Sub

ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarLista", 0)

End Sub

'Envia os itens selecionados para processamento
Private Function flGerenciar() As String

#If EnableSoap = 1 Then
    Dim objContaCorrente                    As MSSOAPLib30.SoapClient30
#Else
    Dim objContaCorrente                    As A8MIU.clsContaCorrente
#End If

Dim xmlLoteLancamentos                      As MSXML2.DOMDocument40
Dim strXMLRetorno                           As String
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long
Dim intIgnoraGradeHorario                   As Integer
Dim strJustificativa                        As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set xmlLoteLancamentos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLoteLancamentos, "", "Repeat_Filtros", "")
    
    With lvwContaCorrente.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlLoteLancamentos, "Repeat_Filtros", "Grupo_Lote", "")
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "NU_SEQU_OPER_ATIV", Split(.Item(lngCont).Key, "|")(POS_NU_SEQU_OPER_ATIV), "Repeat_Filtros")

                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "TP_LANC_ITGR", Split(.Item(lngCont).Key, "|")(POS_TP_LANC_ITGR), "Repeat_Filtros")

                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "CO_ULTI_SITU_PROC", enumStatusIntegracao.IntegradoCC, "Repeat_Filtros")
                    
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "DH_ULTI_ATLZ", Split(.Item(lngCont).Key, "|")(POS_DH_ULTI_ATLZ), "Repeat_Filtros")
                
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "NR_SEQU_LANC", Split(.Item(lngCont).Key, "|")(POS_NR_SEQU_LANC), "Repeat_Filtros")
                
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "TipoLancamentoIntegracao", Split(.Item(lngCont).Key, "|")(POS_TP_LANC_ITGR), "Repeat_Filtros")
                
                'Cesar 08/05/2007 - Conta Corrente Corretoras
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "VA_LANC_CC", fgVlr_To_Xml(.Item(lngCont).SubItems(COL_VALOR)), "Repeat_Filtros")
                
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "IN_LANC_DEBT_CRED", IIf(.Item(lngCont).SubItems(COL_TIPO_MOVIMENTO) = "Débito", _
                                                                 enumTipoDebitoCredito.Debito, _
                                                                 enumTipoDebitoCredito.Credito), "Repeat_Filtros")
                
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "NetOperacoes", .Item(lngCont).SubItems(COL_COND_NET_OPERACOES), "Repeat_Filtros")
                
                intIgnoraGradeHorario = IIf(.Item(lngCont).ForeColor = vbRed And _
                                            .Item(lngCont).Tag = COD_ERRO_NEGOCIO_GRADE, 1, 0)
                                            
                Call fgAppendNode(xmlLoteLancamentos, _
                          "Grupo_Lote", "IgnoraGradeHorario", intIgnoraGradeHorario, "Repeat_Filtros")
                          
                Call fgAppendNode(xmlLoteLancamentos, _
                            "Grupo_Lote", "TX_JUST_CANC", strJustificativa, "Repeat_Filtros")
                          
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        Set objContaCorrente = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
        strXMLRetorno = objContaCorrente.Gerenciar(xmlLoteLancamentos.xml, _
                                                   vntCodErro, _
                                                   vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objContaCorrente = Nothing
        
        'Verifica se o retorno da operação possui erros
        If strXMLRetorno <> vbNullString Then
            '...se sim, carrega o XML de Erros
            Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlRetornoErro.loadXML(strXMLRetorno)
        Else
            '...se não, apenas destrói o objeto
            Set xmlRetornoErro = Nothing
        End If
        
        flGerenciar = strXMLRetorno
    Else
        flGerenciar = vbNullString
    End If
    
    Set xmlLoteLancamentos = Nothing

    Exit Function

ErrorHandler:
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    fgRaiseError App.EXEName, Me.Name, "flGerenciar", 0
    'mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flGerenciar", Me.Caption

End Function

'Inicializa controles de tela e variáveis
Private Sub flInicializar()

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
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Limpa o conteúdo do ListView
Private Sub flLimparLista()
    Me.lvwContaCorrente.ListItems.Clear
End Sub

'Exibe de forma diferenciada quaisquer itens que tenham sido rejeitados por motivo de grade de horário
Private Sub flMarcarRejeitadosPorGradeHorario()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_GRADE & "']")
            With lvwContaCorrente.ListItems
                For lngCont = 1 To .Count
                    If Split(.Item(lngCont).Key, "|")(POS_NU_SEQU_OPER_ATIV) = objDomNode.selectSingleNode("Operacao").Text Then
                        For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                            .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                        Next
                        
                        .Item(lngCont).Text = "Horário Excedido"
                        .Item(lngCont).ToolTipText = "Horário limite p/envio da mensagem excedido"
                        .Item(lngCont).ForeColor = vbRed
                        .Item(lngCont).Tag = COD_ERRO_NEGOCIO_GRADE
                        
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

'Exibe de forma diferenciada quaisquer itens que tenham sido rejeitados por motivo de estorno
Private Sub flMarcarRejeitadosPorProcessoEstorno()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngCont                                 As Long
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    If Not xmlRetornoErro Is Nothing Then
        For Each objDomNode In xmlRetornoErro.documentElement.selectNodes("Grupo_ControleErro[CodigoErro='" & COD_ERRO_NEGOCIO_ESTORNO & "']")
            With lvwContaCorrente.ListItems
                For lngCont = 1 To .Count
                    If Split(.Item(lngCont).Key, "|")(POS_NU_SEQU_OPER_ATIV) = objDomNode.selectSingleNode("Operacao").Text Then
                        For intContAux = 1 To .Item(lngCont).ListSubItems.Count
                            .Item(lngCont).ListSubItems(intContAux).ForeColor = vbRed
                        Next
                        
                        .Item(lngCont).Text = "Operação Processo Estorno"
                        .Item(lngCont).ToolTipText = "Operação estornada, ou em processo de estorno"
                        .Item(lngCont).ForeColor = vbRed
                        .Item(lngCont).Tag = COD_ERRO_NEGOCIO_ESTORNO
                        
                        Exit For
                    End If
                Next
            End With
        Next
    End If

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMarcarRejeitadosPorProcessoEstorno", 0

End Sub

'Exibe o resultado do último processamento
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " integrados "
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

Dim objListItem                             As MSComctlLib.ListItem
Dim blnEncontrouDataFutura                  As Boolean

    On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(lvwContaCorrente, Retorno)
    
            If Retorno = enumTipoSelecao.MarcarTodas Then
                blnEncontrouDataFutura = False
                
                For Each objListItem In lvwContaCorrente.ListItems
                    If Val(fgDt_To_Xml(objListItem.SubItems(COL_DATA_OPERACAO))) > Val(fgDt_To_Xml(fgDataHoraServidor(DataAux))) Then
                        blnEncontrouDataFutura = True
                        objListItem.Checked = False
                    End If
                Next
                
                If blnEncontrouDataFutura Then
                    With frmMural
                        .Display = "Lançamentos com data futura não podem ser integrados em D0. Seleção não permitida."
                        .Show vbModal
                    End With
                End If
            End If
            
    End Select
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call tlbComandos_ButtonClick(tlbComandos.Buttons("refresh"))
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_KeyDown", Me.Caption
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    
    tlbComandos.Buttons("integrar").Caption = "Integrar C/C"
    fraJustificativa.Visible = False
    
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    
    Call flInicializar
    blnPrimeiraConsulta = True
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmIntegrarCCOnLine
    Load objFiltro
    objFiltro.fgCarregarPesquisaAnterior
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .lvwContaCorrente.Top = 0
        .lvwContaCorrente.Height = .Height - 720
        .lvwContaCorrente.Width = .Width - 120
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
End Sub

Private Sub lvwContaCorrente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwContaCorrente, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwContaCorrente_DblClick", Me.Caption

End Sub

Private Sub lvwContaCorrente_DblClick()
    
On Error GoTo ErrorHandler

    If Not lvwContaCorrente.SelectedItem Is Nothing Then
        If lvwContaCorrente.SelectedItem.ForeColor = vbRed Then Exit Sub
        
        With frmHistLancamentoCC
            .lngCodigoEmpresa = fgObterCodigoCombo(lvwContaCorrente.SelectedItem)
            .vntSequenciaOperacao = Split(lvwContaCorrente.SelectedItem.Key, "|")(POS_NU_SEQU_OPER_ATIV)
            .lngTipoLancamentoITGR = Split(lvwContaCorrente.SelectedItem.Key, "|")(POS_TP_LANC_ITGR)
            .intSequenciaLancamento = Split(lvwContaCorrente.SelectedItem.Key, "|")(POS_NR_SEQU_LANC)
            .strNetOperacoes = lvwContaCorrente.SelectedItem.SubItems(COL_COND_NET_OPERACOES)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwContaCorrente_DblClick", Me.Caption

End Sub

Private Sub lvwContaCorrente_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
    
    If Item.Checked Then
        If Val(fgDt_To_Xml(Item.SubItems(COL_DATA_OPERACAO))) > Val(fgDt_To_Xml(fgDataHoraServidor(DataAux))) Then
            With frmMural
                .Display = "Lançamentos com data futura não podem ser integrados em D0. Seleção não permitida."
                .Show vbModal
            End With
            
            Item.Checked = False
        End If
    End If
End Sub

Private Sub lvwContaCorrente_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call tlbComandos_ButtonClick(tlbComandos.Buttons("refresh"))
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmIntegrarCCOnLine - lvwContaCorrente_KeyDown", Me.Caption

End Sub

Private Sub lvwContaCorrente_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwContaCorrente_MouseDown", Me.Caption

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

On Error GoTo ErrorHandler

    strDocFiltros = xmlDocFiltros
    fgCursor True
    
    If Trim$(xmlDocFiltros) <> "" Then
        If fgMostraFiltro(strDocFiltros, blnPrimeiraConsulta) Then
            blnPrimeiraConsulta = False
            Call tlbComandos_ButtonClick(tlbComandos.Buttons("DefinirFiltro"))
        End If
        
        Me.tlbComandos.Buttons("AplicarFiltro").value = tbrPressed
        
        If InStr(1, strDocFiltros, "DataIni") = 0 Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
        Else
            Call flCarregarLista
        End If
    Else
        Call flLimparLista
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - objFiltro_AplicarFiltro", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    Call fgCursor(True)
    
    Select Case Button.Key
        Case "DefinirFiltro"
            blnPrimeiraConsulta = False
            objFiltro.Show vbModal
            
        Case "refresh"
            If InStr(1, strDocFiltros, "DataIni") = 0 Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
            Else
                Call flCarregarLista
            End If
            
        Case "integrar"
            strResultadoOperacao = flGerenciar
            
            If strResultadoOperacao <> vbNullString Then
                Call flMostrarResultado(strResultadoOperacao)
                Call flCarregarLista
            End If
            
            Call flMarcarRejeitadosPorGradeHorario
            Call flMarcarRejeitadosPorProcessoEstorno
            Button.Enabled = True: DoEvents
        Case gstrSair
            Unload Me
            
    End Select
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
'Calcula o Net da operações
Public Function flNetOperacoes(ByVal strItemKey As String)
    
Dim strExpression                   As String
Dim vntValor                        As Variant
    
    vntValor = 0
    
    strExpression = flMontarCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlRetLeitura, strExpression))
    
    flNetOperacoes = vntValor

End Function
'Monta uma expressão XPath para a somatória dos valores de operações
Public Function flMontarCalculoNetOperacoes(ByVal strItemKey As String)
                
Dim strDebito                               As String
Dim strCredito                              As String
    
On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                     " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                     " and ../DT_OPER='" & Split(strItemKey, "|")(KEY_DATA_OPERACAO) & "' " & _
                                     " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                     " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                     " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                     " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' " & _
                                     " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_CD_LOTE) & "' "
    
    strCredito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                     " and ../DT_OPER='" & Split(strItemKey, "|")(KEY_DATA_OPERACAO) & "' " & _
                                     " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                     " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                     " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                     " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' " & _
                                     " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_CD_LOTE) & "' "

    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "
    
    flMontarCalculoNetOperacoes = strDebito & strCredito
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCalculoNetOperacoes", 0

End Function


