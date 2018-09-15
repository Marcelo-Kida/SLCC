VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComplementacaoOperacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ferramentas - Complementação Operação Compromissada"
   ClientHeight    =   7395
   ClientLeft      =   780
   ClientTop       =   2265
   ClientWidth     =   13215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   13215
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13235
      _ExtentX        =   23336
      _ExtentY        =   12356
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo Operação"
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
         Text            =   "Tipo Compromisso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tipo Compromisso Retorno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Veículo Legal (Parte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Contra-Parte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Tipo Movto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Data Op. Ret."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Título"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Dt.Vencto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Local de Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Agendamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Cta. Cedente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Cta. Cessionário"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
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
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":5922
            Key             =   "Confirmar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":5F56
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":6270
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementacaoOperacao.frx":658A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   7065
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   582
      ButtonWidth     =   2540
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "Atualizar"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Complementar"
            Key             =   "Complementar"
            Object.ToolTipText     =   "Complementar Operação"
            ImageKey        =   "AlterarAgendamento"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                 "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmComplementacaoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:12
'-------------------------------------------------
'' Objeto reponsável pela complementação de uma operação, através de interação com
'' a camada de controle de caso de uso MIU
''
'' Classes especificamente consideradas de destino:
''   A8MIU.clsMIU
''   A8MIU.clsOperacao
''
Option Explicit

Public strHoraAgendamento                   As String

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlDominioCompromisso               As MSXML2.DOMDocument40
Private strOperacao                         As String

'Constantes de Configuração de Colunas
Private Const COL_TIPO_OPERACAO             As Integer = 0
Private Const COL_DATA_OPERACAO             As Integer = 1
Private Const COL_NUMERO_COMANDO            As Integer = 2
Private Const COL_TIPO_COMPROMISSO          As Integer = 3
Private Const COL_TIPO_COMPROMISSO_RETN     As Integer = 4
Private Const COL_TP_ACAO_OPER_ATIV_EXEC    As Integer = 5
Private Const COL_VEICULO_LEGAL_PARTE       As Integer = 6
Private Const COL_CONTRAPARTE               As Integer = 7
Private Const COL_TIPO_MOVIMENTO            As Integer = 8
Private Const COL_DATA_OP_RET               As Integer = 9
Private Const COL_TITULO                    As Integer = 10
Private Const COL_VALOR                     As Integer = 11
Private Const COL_DATA_VENCIMENTO           As Integer = 12
Private Const COL_DATA_LIQUIDACAO           As Integer = 13
Private Const COL_LOCAL_LIQUIDACAO          As Integer = 14
Private Const COL_EMPRESA                   As Integer = 15
Private Const COL_HORARIO_ENVIO_MSG         As Integer = 16     '<-- Agendamento
Private Const COL_CONTA_CEDENTE             As Integer = 17
Private Const COL_CONTA_CESSIONARIO         As Integer = 18

Private Const strFuncionalidade             As String = "frmComplementacaoOperacao"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private lngIndexClassifList                 As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons(gstrAtualizar))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmComplementacaoOperacao - Form_KeyDown", Me.Caption
    
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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmComplementacaoOperacao - Form_Load", Me.Caption
    
End Sub

'' Obtem as propriedades requeridas pelo objeto através da camada de controle de
'' caso de uso MIU, método A8MIU.clsMiu.ObterMapaNavegacao.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
    Dim objOperacao        As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
    Dim objOperacao        As A8MIU.clsOperacao
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
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmComplementacaoOperacao", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
    Set xmlDominioCompromisso = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    If Not xmlDominioCompromisso.loadXML(objOperacao.ObterDominiosCompromissoOperacao(vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlDominioCompromisso, App.EXEName, "frmComplementacaoOperacao", "flInicializar")
    End If
    
    Set objOperacao = Nothing
    
    Exit Sub

ErrorHandler:

    Set objOperacao = Nothing
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Sub

'' Obtém as operações que podem ser complementadas e preenche o listview com as
'' mesmas, através da camada de controle de caso de uso MIU, método A8MIU.
'' clsOperacao.ObterDetalheOperacao.
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
#End If

Dim objDomNode              As MSXML2.IXMLDOMNode
Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim xmlDomLeitura           As MSXML2.DOMDocument40
Dim strRetLeitura           As String
Dim strSelecaoFiltro        As String
Dim lngCont                 As Long
Dim datDMenos2              As Date
Dim datDZero                As Date
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Call flLimparLista
    
    strSelecaoFiltro = enumStatusOperacao.AComplementar
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    
    'Filtro Datas
    datDMenos2 = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux), 2, enumPaginacao.Anterior)
    datDZero = fgDataHoraServidor(DataAux)

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(datDMenos2)))
    Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(fgDate_To_DtXML(datDZero)))
    
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
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarLista")
        End If
    
        For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheOperacao/*")
            With lstOperacao.ListItems.Add(, _
                    "k" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                    
                'Guarda na propriedade TAG o tipo da operação e
                'a data da última atualização
                .Tag = objDomNode.selectSingleNode("TP_OPER").Text & "|" & _
                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
    
                If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                End If
                
                If objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text <> vbNullString And fgVlrXml_To_Decimal(objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) <> 0 Then
                    .SubItems(COL_TP_ACAO_OPER_ATIV_EXEC) = objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text & " - " & fgDescricaoTipoAcao(CLng("0" & objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text))
                End If
                .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_TITULO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                
                If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
                End If
                
                .Text = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                
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
                
                If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text), "HH:MM")
                End If
                
                .SubItems(COL_CONTA_CEDENTE) = objDomNode.selectSingleNode("CO_CNTA_CEDT").Text
                .SubItems(COL_CONTA_CESSIONARIO) = objDomNode.selectSingleNode("CO_CNTA_CESS").Text
                
                If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                End If

                If objDomNode.selectSingleNode("TP_CPRO_OPER_ATIV").Text <> vbNullString And _
                    Val(objDomNode.selectSingleNode("TP_CPRO_OPER_ATIV").Text) <> 0 Then

                    .SubItems(COL_TIPO_COMPROMISSO) = objDomNode.selectSingleNode("TP_CPRO_OPER_ATIV").Text & " - " & _
                                                      xmlDominioCompromisso.documentElement. _
                                                         selectSingleNode("Repeat_DominioAtributo[@NO_ATRIBUTO='TP_CPRO_OPER_ATIV']" & _
                                                                          "/Grupo_DominioAtributo[CO_DOMI=" & objDomNode.selectSingleNode("TP_CPRO_OPER_ATIV").Text & "]/DE_DOMI").Text
                End If
                
                
                If objDomNode.selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text <> vbNullString And _
                    Val(objDomNode.selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text) <> 0 Then

                    .SubItems(COL_TIPO_COMPROMISSO_RETN) = objDomNode.selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text & " - " & _
                                                           xmlDominioCompromisso.documentElement. _
                                                              selectSingleNode("Repeat_DominioAtributo[@NO_ATRIBUTO='TP_CPRO_RETN_OPER_ATIV']" & _
                                                                               "/Grupo_DominioAtributo[CO_DOMI=" & objDomNode.selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text & "]/DE_DOMI").Text
                                                                               
                End If
                
            End With
        Next
    End If
    
    Call fgClassificarListview(Me.lstOperacao, lngIndexClassifList, True)
    
    Set xmlDomLeitura = Nothing

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
   
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With lstOperacao
        .Top = 0
        .Left = 0
        .Width = Me.Width - 100
        .Height = (Me.Height - tlbFiltro.Height) - 1000
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmComplementacaoOperacao = Nothing
End Sub

Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ColumnClick"

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strSelecaoFiltro                        As String
Dim strResultadoConfirmacao                 As String

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
            
        Case "Complementar"
            flComplementar
            
        Case gstrAtualizar
            flCarregarLista
        
        Case gstrSair
            Unload Me
            
    End Select
    fgCursor
    Exit Sub

ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmComplementacaoOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

'' Carrega e exibe o objeto frmTipoCompromisso para complementação da operação.
Private Sub flComplementar()

Dim lstItem                                 As MSComctlLib.ListItem
Dim strTipoCompromisso                      As String
Dim strXmlTipoCompromisso                   As String
Dim strAcaoAnterior                         As String

Const POS_TP_OPER                           As Integer = 0
Const POS_DH_ULTI_ATLZ                      As Integer = 1

On Error GoTo ErrorHandler

    If Not lstOperacao.SelectedItem Is Nothing Then
        Set lstItem = lstOperacao.SelectedItem
        
        If lstItem.SubItems(COL_TIPO_COMPROMISSO) <> vbNullString Then
            strTipoCompromisso = fgObterCodigoCombo(lstItem.SubItems(COL_TIPO_COMPROMISSO))
        ElseIf lstItem.SubItems(COL_TIPO_COMPROMISSO_RETN) <> vbNullString Then
            strTipoCompromisso = fgObterCodigoCombo(lstItem.SubItems(COL_TIPO_COMPROMISSO_RETN))
        End If
        
        If Split(lstItem.Tag, "|")(POS_TP_OPER) = enumTipoOperacaoLQS.CompromissadaIda Then
            strXmlTipoCompromisso = xmlDominioCompromisso.selectSingleNode("Grupo_Dados/Repeat_DominioAtributo[@NO_ATRIBUTO='TP_CPRO_OPER_ATIV']").xml
            strAcaoAnterior = lstItem.SubItems(COL_TIPO_COMPROMISSO)
        Else
            strXmlTipoCompromisso = xmlDominioCompromisso.selectSingleNode("Grupo_Dados/Repeat_DominioAtributo[@NO_ATRIBUTO='TP_CPRO_RETN_OPER_ATIV']").xml
            strAcaoAnterior = lstItem.SubItems(COL_TIPO_COMPROMISSO_RETN)
        End If
        
        With frmAlteracaoTipoCompromisso
            .NumeroSequencia = Mid$(lstItem.Key, 2)
            .Comando = lstItem.SubItems(COL_NUMERO_COMANDO)
            .DataUltimaAtualizacao = Split(lstItem.Tag, "|")(POS_DH_ULTI_ATLZ)
            .TipoCompromisso = strTipoCompromisso
            .XmlTipoCompromisso = strXmlTipoCompromisso
            .AcaoAnterior = strAcaoAnterior
            .TipoOperacao = Split(lstItem.Tag, "|")(POS_TP_OPER)
            .StatusOperacao = enumStatusOperacao.AComplementar
            .Show vbModal
        End With
        
        flCarregarLista
        
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flComplementar", 0

End Sub

'' Limpa o listview de operações.
Private Sub flLimparLista()
    lstOperacao.ListItems.Clear
End Sub
