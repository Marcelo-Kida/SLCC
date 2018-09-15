VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCancelamentoEntradaManual 
   Caption         =   "Ferramentas - Cancelamento Entrada Manual"
   ClientHeight    =   8580
   ClientLeft      =   240
   ClientTop       =   1320
   ClientWidth     =   14235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   4725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   8334
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selecionar"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Veículo Legal(Parte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contra-Parte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tipo Movto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Título"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Data Vencto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Operação/Evento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tipo Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Cta Cedente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Cta. Cessionário"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8250
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   582
      ButtonWidth     =   3201
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
            Caption         =   "Cancelar        "
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Operações"
            Key             =   "MostrarOperacao"
            Object.ToolTipText     =   "Mostrar Operações"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Mensagens"
            Key             =   "MostrarMensagem"
            Object.ToolTipText     =   "Mostrar Mensagens"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                 "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   9720
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
            Picture         =   "frmCancelamentoEntradaManual.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelamentoEntradaManual.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   4725
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   8334
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selecionar"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Código Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Agendamento"
         Object.Width           =   2540
      EndProperty
   End
   Begin A8.ctlMenu ctlMenu1 
      Left            =   13200
      Top             =   4620
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin VB.Image imgDummyH 
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   4800
      Width           =   14040
   End
End
Attribute VB_Name = "frmCancelamentoEntradaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:10
'-------------------------------------------------
'' Objeto reponsável pelo cancelamento de uma operação ou entrada manual, através
'' de interação com a camada de controle de caos de uso MIU.
''
'' Classes consideradas especificamente de destino:
''   A8MIU.clsMIU
''   A8MIU.clsOperacaoMensagem
''   A8MIU.clsOperacao
''   A8MIU.clsMensagem
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private intControleMenuPopUp                As enumTipoConfirmacao

Private fblnDummyH                          As Boolean

'Constantes de Configuração de Colunas de Operação
Private Const COL_OP_DATA_OPERACAO          As Integer = 1
Private Const COL_OP_NUM_COMANDO            As Integer = 2
Private Const COL_OP_VEICULO_LEGAL_PARTE    As Integer = 3
Private Const COL_OP_CONTRAPARTE            As Integer = 4
Private Const COL_OP_SITUACAO               As Integer = 5
Private Const COL_OP_TIPO_MOVIMENTO         As Integer = 6
Private Const COL_OP_TITULO                 As Integer = 7
Private Const COL_OP_VALOR                  As Integer = 8
Private Const COL_OP_DATA_VENCIMENTO        As Integer = 9
Private Const COL_OP_TIPO_OPER              As Integer = 10
Private Const COL_OP_DATA_LIQUIDACAO        As Integer = 11
Private Const COL_OP_TIPO_LIQUIDACAO        As Integer = 12
Private Const COL_OP_EMPRESA                As Integer = 13
Private Const COL_OP_CONTA_CEDENTE          As Integer = 14
Private Const COL_OP_CONTA_CESSIONARIO      As Integer = 15

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_MSG_DATA_MENSAGEM         As Integer = 1
Private Const COL_MSG_NUMERO_COMANDO        As Integer = 2
Private Const COL_MSG_CODIGO_MENSAGEM       As Integer = 3
Private Const COL_MSG_SITUACAO              As Integer = 4
Private Const COL_MSG_EMPRESA               As Integer = 5
Private Const COL_MSG_AGENDAMENTO           As Integer = 6

Private Const strFuncionalidade             As String = "frmCancelamentoEntradaManual"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'' Cancela um lote de operações e/ou mensagem através de interação com a camada de
'' controle de caso de uso MIU, método A8MIU.clsOperacaoMensagem.
'' CancelarEntradaManual
Private Function flCancelar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem                 As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem                 As A8MIU.clsOperacaoMensagem
#End If

Dim xmlDomLoteOperacaoMensagem              As MSXML2.DOMDocument40
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Const POS_STATUS                            As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1

Const POS_NUMERO_CONTROLE_IF                As Integer = 0
Const POS_DATA_REGISTRO_MENSAGEM_SPB        As Integer = 1

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ---------------------------------------------------------------------------
    Set xmlDomLoteOperacaoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomLoteOperacaoMensagem, "", "Repeat_Filtros", "")
    
    'Captura o filtro cumulativo OPERAÇÃO
    With lvwOperacao.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Repeat_Filtros", "Grupo_Lote", "")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.Operacao, "Repeat_Filtros")
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "Operacao", Mid(.Item(lngCont).Key, 2), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
            End If
        Next
    End With
    
    'Captura o filtro cumulativo MENSAGEM
    With lvwMensagem.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, "Repeat_Filtros", "Grupo_Lote", "")
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.MENSAGEM, "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "NumeroControleIF", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NUMERO_CONTROLE_IF), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "DTRegistroMensagemSPB", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DATA_REGISTRO_MENSAGEM_SPB), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "CodigoMensagem", .Item(lngCont).SubItems(COL_MSG_CODIGO_MENSAGEM), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacaoMensagem, _
                          "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
                
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        flCancelar = objOperacaoMensagem.CancelarEntradaManual(xmlDomLoteOperacaoMensagem.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objOperacaoMensagem = Nothing
    Else
        flCancelar = vbNullString
    End If
    
    Set xmlDomLoteOperacaoMensagem = Nothing

Exit Function
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - flCancelar", Me.Caption

End Function

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, _
             enumTipoSelecao.DesmarcarTodas
            
            Call fgMarcarDesmarcarTodas(IIf(intControleMenuPopUp = enumTipoConfirmacao.Operacao, _
                                                lvwOperacao, lvwMensagem), _
                                        Retorno)
    
    End Select
    
    Exit Sub
    
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConfirmacaoOperacao - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("refresh"))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    Call flInicializar
    flCarregarListaOperacao
    flCarregarListaMensagem
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - Form_Load", Me.Caption
    
End Sub

'' Obtém e preenche o listview com as operações possíveis de cancelamento através
'' de interação com a camada de controle de caso de uso MIU, método A8MIU.
'' clsOperacao.ObterDetalheOperacao
Private Sub flCarregarListaOperacao()

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strRetLeitura                           As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim strSelecaoFiltro                        As String
Dim strSelecaoAcao                          As String
Dim lngCont                                 As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim blnIncluirRegistro                      As Boolean

On Error GoTo ErrorHandler

    Call flLimparLista(lvwOperacao)
    
    strSelecaoFiltro = enumStatusOperacao.Concordancia & ";" & _
                       enumStatusOperacao.ManualEmSer & ";" & _
                       enumStatusOperacao.Registrada
                       
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlDomFiltros, "", "Repeat_Filtros", ""
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", ""
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_EntradaManual", ""
    fgAppendNode xmlDomFiltros, "Grupo_EntradaManual", "EntradaManual", enumIndicadorSimNao.Sim
    '>>> -------------------------------------------------------------------------------------------
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaOperacao")
        End If
    
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                    
                blnIncluirRegistro = True
                
                If Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Registrada Then
                    Select Case Val(objDomNode.selectSingleNode("TP_OPER").Text)
                        Case enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP, _
                             enumTipoOperacaoLQS.NETEntradaManualMultilateralBMA, _
                             enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC, _
                             enumTipoOperacaoLQS.NETEntradaManualMultilateralBMD, _
                             enumTipoOperacaoLQS.NETEntradaManualMultilateralCBLC, _
                             enumTipoOperacaoLQS.NETEntradaManualMultilateralCETIP
                            
                            blnIncluirRegistro = True
                            
                        Case Else
                    
                            blnIncluirRegistro = False
                            
                    End Select
                End If
               
                If blnIncluirRegistro Then
                    
                    With lvwOperacao.ListItems.Add(, "k" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                    
                    'Guarda na propriedade TAG a situação da operação | comando da ação
                    .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                           objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                           objDomNode.selectSingleNode("NU_COMD_ACAO_EXEC").Text
            
                    If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                    End If
                    
                    .SubItems(COL_OP_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    .SubItems(COL_OP_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                    .SubItems(COL_OP_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_OP_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                    .SubItems(COL_OP_TITULO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                    .SubItems(COL_OP_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                    .SubItems(COL_OP_TIPO_OPER) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                    .SubItems(COL_OP_NUM_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                    .SubItems(COL_OP_TIPO_LIQUIDACAO) = objDomNode.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text
                                    
                    If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                    End If
                    
                    If objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                        .SubItems(COL_OP_DATA_LIQUIDACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
                    End If
                    
                    If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                       Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                        'Obtem a descrição da Empresa via QUERY XML
                        .SubItems(COL_OP_EMPRESA) = objDomNode.selectSingleNode("CO_EMPR").Text & " - " & _
                            xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                                objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                    End If
                    
                    .SubItems(COL_OP_CONTA_CEDENTE) = objDomNode.selectSingleNode("CO_CNTA_CEDT").Text
                    .SubItems(COL_OP_CONTA_CESSIONARIO) = objDomNode.selectSingleNode("CO_CNTA_CESS").Text
                End With
            End If
            
            
        Next
    End If
    
    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaOperacao", 0
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .tlbFiltro.Left = 0
        .tlbFiltro.Top = .ScaleHeight - .tlbFiltro.Height
        
        .imgDummyH.Left = 0
        .imgDummyH.Width = .ScaleWidth
        
        .lvwOperacao.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .imgDummyH.Height, _
                                                      .tlbFiltro.Top - 300)
        .lvwOperacao.Width = .Width - 100
        
        .lvwMensagem.Top = IIf(.imgDummyH.Visible, .imgDummyH.Top + .imgDummyH.Height, 0)
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 300, _
                                                      .tlbFiltro.Top - 300)
        .lvwMensagem.Width = .Width - 100
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCancelamentoEntradaManual = Nothing
End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = True
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not fblnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyH.Top = y + imgDummyH.Top

    On Error Resume Next
    
    With Me
        If .imgDummyH.Top < 950 Then
            .imgDummyH.Top = 950
        End If
        If .imgDummyH.Top > (.Height - 2100) And (.Height - 2100) > 0 Then
            .imgDummyH.Top = .Height - 2100
        End If
        
        .lvwOperacao.Height = .imgDummyH.Top - .imgDummyH.Height
        .lvwMensagem.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwMensagem.Height = .tlbFiltro.Top - .imgDummyH.Top - .imgDummyH.Height - 300
    End With
    
    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyH = False
End Sub

Private Sub lvwMensagem_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lvwMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwMensagem, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwMensagem_ColumnClick"
End Sub

Private Sub lvwMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lvwOperacao.SelectedItem = Nothing
End Sub

Private Sub lvwMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.MENSAGEM
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

    Exit Sub
    
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - lvwMensagem_MouseDown", Me.Caption

End Sub

Private Sub lvwOperacao_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lvwOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwOperacao, ColumnHeader.Index)
    lngIndexClassifListOper = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwOperacao_ColumnClick"

End Sub

Private Sub lvwOperacao_DblClick()
    
On Error GoTo ErrorHandler

    If Not lvwOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .CodigoEmpresa = fgObterCodigoCombo(lvwOperacao.SelectedItem.ListSubItems(COL_OP_EMPRESA))
            .SequenciaOperacao = Mid(lvwOperacao.SelectedItem.Key, 2)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwOperacao_DblClick"
    
End Sub

Private Sub lvwOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lvwOperacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lvwMensagem.SelectedItem = Nothing
End Sub

Private Sub lvwOperacao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        intControleMenuPopUp = enumTipoConfirmacao.Operacao
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

    Exit Sub
    
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmLiberacaoOperacao - lvwOperacao_MouseDown", Me.Caption

End Sub

'' Atualiza as informações, cancela um lote de operações/mensagens, configura a
'' visualização dos componentes do objeto e fecha o mesmo.
Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strJanelas                              As String
Dim strSelecaoFiltro                        As String
Dim strResultadoCancelamento                As String

Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1
Const POS_NUM_CONTROLE_IF                   As Integer = 0
Const POS_DATA_REGISTRO_MENSAGEM            As Integer = 1

On Error GoTo ErrorHandler
    
    fgCursor True
    
    If tlbFiltro.Buttons("MostrarOperacao").value = tbrPressed Then
        strJanelas = strJanelas & "1"
    End If
    
    If tlbFiltro.Buttons("MostrarMensagem").value = tbrPressed Then
        strJanelas = strJanelas & "2"
    End If
    
    Call flArranjarJanelasExibicao(strJanelas)
    
    Select Case Button.Key
        
        Case "refresh"
            flCarregarListaOperacao
            flCarregarListaMensagem
                
        Case "Cancelar"
            strResultadoCancelamento = flCancelar
            If strResultadoCancelamento <> vbNullString Then
                Call flMostrarResultado(strResultadoCancelamento)
                flCarregarListaOperacao
                flCarregarListaMensagem
            End If
        
        Case gstrSair
            Unload Me
            
    End Select
    
    fgCursor
    
    Exit Sub

ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - tlbFiltro_ButtonClick", Me.Caption

End Sub

'' Limpa os listview de mensagem e de operação.
Private Sub flLimparLista(ByVal lstListView As ListView)
    lstListView.ListItems.Clear
End Sub

'' Exibe o resultado do cancelamento através do objeto frmResultOperacaoLote
Private Sub flMostrarResultado(ByVal pstrResultadoLiberacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " cancelados "
        .Resultado = pstrResultadoLiberacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'' Carrega as propriedades pertinentes ao objeto
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmCancelamentoEntradaManual", "flInicializar")
    End If
    
    Set objMIU = Nothing
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - flInicializar", Me.Caption

End Sub

'' Obtém e preenche o listview com as mensagens possíveis de cancelamento através
'' de interação com a camada de controle de caso de uso MIU, método A8MIU.
'' clsMensagem.ObterDetalheMensagem.
Private Sub flCarregarListaMensagem()

#If EnableSoap = 1 Then
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem        As A8MIU.clsMensagem
#End If

Dim objDomNode             As MSXML2.IXMLDOMNode
Dim xmlDomFiltros          As MSXML2.DOMDocument40
Dim strRetLeitura          As String
Dim xmlDomLeitura          As MSXML2.DOMDocument40
Dim strSelecaoFiltro       As String
Dim lngCont                As Long
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Call flLimparLista(lvwMensagem)
    
    strSelecaoFiltro = enumStatusMensagem.Concordancia & ";" & _
                       enumStatusMensagem.ManualEmSer
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlDomFiltros, "", "Repeat_Filtros", ""
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", ""
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_EntradaManual", ""
    fgAppendNode xmlDomFiltros, "Grupo_EntradaManual", "EntradaManual", enumIndicadorSimNao.Sim
    
    '>>> -------------------------------------------------------------------------------------------
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaMensagem")
        End If
        
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            With lvwMensagem.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                          objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                    
                'Guarda na propriedade TAG a situação da mensagem | data da úlmtima atualização
                .Tag = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & "|" & _
                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & "|" & _
                       objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    
                If objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_DATA_MENSAGEM) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                End If
                
                .SubItems(COL_MSG_CODIGO_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                .SubItems(COL_MSG_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_MSG_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                
                If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_MSG_AGENDAMENTO) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text)
                End If
                
                If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                   Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                    'Obtem a descrição da Empresa via QUERY XML
                    .SubItems(COL_MSG_EMPRESA) = _
                        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                            objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                End If
                
            End With
        Next
    End If
    
    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    mdiLQS.uctlogErros.MostrarErros Err, "frmCancelamentoEntradaManual - flCarregarListaMensagem", Me.Caption

End Sub

'' Configura a disposição das janelas de acordo com a opção selecionada
Private Sub flArranjarJanelasExibicao(ByVal pstrJanelas As String)
    
On Error GoTo ErrorHandler

    Select Case pstrJanelas
           Case ""
                imgDummyH.Visible = False
                lvwOperacao.Visible = False
                lvwMensagem.Visible = False
            
           Case "1"
                imgDummyH.Visible = False
                lvwOperacao.Visible = True
                lvwMensagem.Visible = False
            
           Case "2"
                imgDummyH.Visible = False
                lvwOperacao.Visible = False
                lvwMensagem.Visible = True
                
           Case "12"
                imgDummyH.Visible = True
                lvwOperacao.Visible = True
                lvwMensagem.Visible = True
                
    End Select
    
    Call Form_Resize

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flArranjarJanelasExibicao", 0
    
End Sub

