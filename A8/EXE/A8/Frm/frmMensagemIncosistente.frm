VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensagemIncosistente 
   Caption         =   "Ferramenta - Mensagem Inconsistente"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   14250
   WindowState     =   2  'Maximized
   Begin A8.ctlMenu ctlMenu1 
      Left            =   3660
      Top             =   0
      _ExtentX        =   2090
      _ExtentY        =   714
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8250
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   582
      ButtonWidth     =   2302
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      HotImageList    =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reprocessar"
            Key             =   "reprocessar"
            ImageKey        =   "reprocessar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "refresh"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r      "
            Key             =   "Sair"
            ImageKey        =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   4200
      Top             =   7920
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
            Picture         =   "frmMensagemIncosistente.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":1286
            Key             =   "reprocessar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemIncosistente.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMensagem 
      Height          =   8715
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   15372
      View            =   3
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
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Num. Controle IF"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Veículo Legal(Parte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo Movto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Entrada / Saída"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Valor da Mensagem"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Local Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Horario Envio/Receb."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Tipo Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Preço Unitário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Taxa de Negociação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Tipo de Informação"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMensagemIncosistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela listagem das mensagem inconsistentes

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private blnOrigemBotaoRefresh               As Boolean

Private Const strFuncionalidade             As String = "frmConsultaMensagem"

'Constantes de Configuração de Colunas
Private Const COL_DATA_MENSAGEM             As Integer = 2
Private Const COL_NUM_CTRL_IF               As Integer = 1
Private Const COL_NUMERO_COMANDO            As Integer = 3
Private Const COL_VEICULO_LEGAL_PARTE       As Integer = 4
Private Const COL_CONTRAPARTE               As Integer = 5
Private Const COL_SITUACAO                  As Integer = 6
Private Const COL_TIPO_MOVIMENTO            As Integer = 7
Private Const COL_ENTRADA_SAIDA             As Integer = 8
Private Const COL_VALOR                     As Integer = 9
Private Const COL_VALOR_MSG                 As Integer = 10
Private Const COL_DATA_LIQUIDACAO           As Integer = 11
Private Const COL_EMPRESA                   As Integer = 12
Private Const COL_LOCAL_LIQUIDACAO          As Integer = 13
Private Const COL_HORARIO_ENVIO_MSG         As Integer = 14

Private Const COL_TIPO_LIQUIDACAO           As Integer = 15
Private Const COL_PRECO_UNITARIO            As Integer = 16
Private Const COL_QUANTIDADE                As Integer = 17
Private Const COL_TAXA_NEGOCIACAO           As Integer = 18
Private Const COL_TIPO_INFORMACAO           As Integer = 19

Private Const POS_NU_CTRL_IF                As Integer = 0
Private Const POS_DH_REGT_MESG_SPB          As Integer = 1
Private Const POS_NU_SEQU_CNTR_REPE         As Integer = 2
Private Const POS_DH_ULTI_ATLZ              As Integer = 3

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifList                 As Long

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(lstMensagem, Retorno)
    End Select
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - ctlMenu1_ClickMenu", Me.Caption
    
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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMensagem - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    Call flInicializar
    Call flCarregarLista
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMensagem - Form_Load", Me.Caption

End Sub

Private Sub Form_Resize()

On Error Resume Next

    With Me
        
        lstMensagem.Left = 30
        lstMensagem.Width = Me.ScaleWidth - 80
        lstMensagem.Height = Me.ScaleHeight - 500
        
    End With
    
End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmConsultaMensagem = Nothing
    
End Sub

Private Sub lstMensagem_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lstMensagem, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMensagem - lstMensagem_ColumnClick", Me.Caption

End Sub

Private Sub lstMensagem_DblClick()

Dim strChave                                As String

Const POS_NUMERO_CTRL_IF                    As Integer = 0
Const POS_DATA_REGISTRO_MESG                As Integer = 1

On Error GoTo ErrorHandler

    If Not lstMensagem.SelectedItem Is Nothing Then
        strChave = Mid$(lstMensagem.SelectedItem.Key, 2)
        With frmDetalheOperacao
            .SequenciaOperacao = lstMensagem.SelectedItem.Tag
            .NumeroControleIF = Split(strChave, "|")(POS_NUMERO_CTRL_IF)
            .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(strChave, "|")(POS_DATA_REGISTRO_MESG))
            .Show vbModal
        End With
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMensagem - lstMensagem_DblClick", Me.Caption
    
End Sub

Private Sub lstMensagem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lstMensagem_MouseDown", Me.Caption

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Select Case Button.Key
        Case "refresh"
             flCarregarLista
        Case "reprocessar"
            
            strResultadoOperacao = flGerenciar
            
            If strResultadoOperacao <> vbNullString Then
                Call flMostrarResultado(strResultadoOperacao)
                Call flCarregarLista
            End If
            
        Case gstrSair
            Unload Me
    End Select
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    blnOrigemBotaoRefresh = False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaMensagem - tlbButtons_ButtonClick", Me.Caption
    
End Sub

'Carrega os dados da lista de mensagens
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRetLeitura                           As String
Dim strDataServidor                         As String
Dim lngCont                                 As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Call flLimparLista
    
    '>>> Formata XML Filtro padrão ----------------------------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusMensagem.MensagemInconsistente)
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_SegregaBackOffice", "SegregaBackOffice", "False")
    
    strDataServidor = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(strDataServidor))
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtXML_To_Oracle("99991231"))

    '>>> -----------------------------------------------------------------------------------------------------------

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(xmlDomFiltros.xml, _
                                                     vntCodErro, _
                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
            
            With lstMensagem.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                          objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & "|" & _
                          objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & "|" & _
                          objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text, _
                          objDomNode.selectSingleNode("CO_MESG_SPB").Text)
                
                .Tag = objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
                .SubItems(COL_NUM_CTRL_IF) = objDomNode.selectSingleNode("NU_CTRL_IF").Text
                .SubItems(COL_DATA_MENSAGEM) = fgDtXML_To_Date(Mid$(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text, 1, 8))
                .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                .SubItems(COL_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_ENTRADA_SAIDA) = objDomNode.selectSingleNode("IN_ENTR_SAID_RECU_FINC").Text
                
                .SubItems(COL_TIPO_LIQUIDACAO) = objDomNode.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text
                .SubItems(COL_PRECO_UNITARIO) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("PU_ATIV_MERC").Text)
                .SubItems(COL_QUANTIDADE) = objDomNode.selectSingleNode("QT_ATIV_MERC").Text
                .SubItems(COL_TAXA_NEGOCIACAO) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("PE_TAXA_NEGO").Text)
                .SubItems(COL_TIPO_INFORMACAO) = objDomNode.selectSingleNode("TP_INFO_LDL").Text
               
                If objDomNode.selectSingleNode("VA_OPER_ATIV").Text <> vbNullString Then
                    .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                End If

                If objDomNode.selectSingleNode("VA_FINC").Text <> vbNullString Then
                    .SubItems(COL_VALOR_MSG) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                End If

                If objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_LIQUIDACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
                End If
                
                If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                   Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                    'Obtem a descrição da Empresa via QUERY XML
                    .SubItems(COL_EMPRESA) = _
                        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                            objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                End If
                
                If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text <> vbNullString And _
                   Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                    
                    If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU") Is Nothing Then
                        
                        'Obtem a descrição do Local de Liquidação via QUERY XML
                        .SubItems(COL_LOCAL_LIQUIDACAO) = _
                            xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU").Text
                
                    Else
                        
                        vntCodErro = 5
                        vntMensagemErro = "Usuário não possui acesso ao Local de Liquidação " & _
                                          objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "."
                        GoTo ErrorHandler
                        
                    End If
                
                End If
                
'                If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                If objDomNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                   .SubItems(COL_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text), "HH:MM")
                End If
            End With
        Next
    End If
    
    Call fgClassificarListview(Me.lstMensagem, lngIndexClassifList, True)
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing

    Exit Sub

ErrorHandler:
    Set objMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarLista", 0)

End Sub

'Limpa o conteúdo da lista
Private Sub flLimparLista()
    Me.lstMensagem.ListItems.Clear
End Sub

'Inicializa os controles de tela e variáveis
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
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConfirmacaoOperacao", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

End Sub

'Procede com o reprocessamento das mensagens selecionadas
Private Function flGerenciar() As String

#If EnableSoap = 1 Then
    Dim objContaCorrente                    As MSSOAPLib30.SoapClient30
#Else
    Dim objContaCorrente                    As A8MIU.clsMensagem
#End If

Dim xmlLoteReprocessamento                  As MSXML2.DOMDocument40
Dim xmlRetornoErro                          As MSXML2.DOMDocument40
Dim strXMLRetorno                           As String
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set xmlLoteReprocessamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLoteReprocessamento, "", "Repeat_Filtros", "")
        
    'Captura o filtro cumulativo MENSAGEM
    With lstMensagem.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlLoteReprocessamento, "Repeat_Filtros", "Grupo_Lote", "")
                
                Call fgAppendNode(xmlLoteReprocessamento, _
                          "Grupo_Lote", "NU_CTRL_IF", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NU_CTRL_IF), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlLoteReprocessamento, _
                          "Grupo_Lote", "DH_REGT_MESG_SPB", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DH_REGT_MESG_SPB), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlLoteReprocessamento, _
                          "Grupo_Lote", "NU_SEQU_CNTR_REPE", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_NU_SEQU_CNTR_REPE), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlLoteReprocessamento, _
                          "Grupo_Lote", "DH_ULTI_ATLZ", Split(Mid(.Item(lngCont).Key, 2), "|")(POS_DH_ULTI_ATLZ), "Repeat_Filtros")
                        
                Call fgAppendNode(xmlLoteReprocessamento, _
                          "Grupo_Lote", "CO_MESG", .Item(lngCont).Text, "Repeat_Filtros")
                        
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        Set objContaCorrente = fgCriarObjetoMIU("A8MIU.clsMensagem")
        strXMLRetorno = objContaCorrente.ReprocessarMensagemInconcistente(xmlLoteReprocessamento.xml, _
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
    
    Set xmlLoteReprocessamento = Nothing

Exit Function
ErrorHandler:
    Set xmlLoteReprocessamento = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flGerenciar", Me.Caption

End Function

'Exibe o resultado do último reprocessamento
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " Mensagens reprocessadas "
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

Exit Sub

ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

