VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlteracaoStatusOperacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contingência - Baixar/Liquidar Operação"
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
   Begin VB.Timer tmrInfo 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   12900
      Top             =   6240
   End
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   6645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   11721
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
      NumItems        =   17
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
         Text            =   "Veículo Legal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contra-Parte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Situação"
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
         Text            =   "Tipo Operação"
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
         Text            =   "Agendamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Cta. Cedente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
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
            Picture         =   "frmAlteracaoStatusOperacao.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":5922
            Key             =   "Confirmar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":5F56
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":6270
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoStatusOperacao.frx":658A
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
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "Atualizar"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixar        "
            Key             =   "Baixar"
            Object.ToolTipText     =   "Baixar Operação"
            ImageKey        =   "AlterarAgendamento"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Liquidar       "
            Key             =   "Liquidar"
            Object.ToolTipText     =   "Liquidar Operação"
            ImageKey        =   "AlterarAgendamento"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                 "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "lblInfo"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   6660
      Visible         =   0   'False
      Width           =   13095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAlteracaoStatusOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 10:52:16
'-------------------------------------------------
'' Objeto responsável por atualizar o status da operação em caso de contingência
'' de sistema. Permite que a operação seja baixada ou liquidada através de
'' interação com a camada controladora de caso de uso A8MIU
''
'' São consideradas especificamente classes de destino:
''   A8MIU.clsOperacao
''
''

Option Explicit

Public strHoraAgendamento                   As String

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas
Private Const COL_DATA_OPERACAO             As Integer = 1
Private Const COL_NUMERO_COMANDO            As Integer = 2
Private Const COL_VEICULO_LEGAL_PARTE       As Integer = 3
Private Const COL_CONTRAPARTE               As Integer = 4
Private Const COL_SITU_PROC                 As Integer = 5
Private Const COL_TIPO_MOVIMENTO            As Integer = 6
Private Const COL_TITULO                    As Integer = 7
Private Const COL_VALOR                     As Integer = 8
Private Const COL_DATA_VENCIMENTO           As Integer = 9
Private Const COL_TIPO_OPERACAO             As Integer = 10
Private Const COL_DATA_LIQUIDACAO           As Integer = 11
Private Const COL_TIPO_LIQUIDACAO           As Integer = 12
Private Const COL_EMPRESA                   As Integer = 13
Private Const COL_HORARIO_ENVIO_MSG         As Integer = 14     '<-- Agendamento
Private Const COL_CONTA_CEDENTE             As Integer = 15
Private Const COL_CONTA_CESSIONARIO         As Integer = 16

Private Const strFuncionalidade             As String = "frmAlteracaoStatusOperacao"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const INFO_BAIXAR                   As String = "Utilizar Baixar quando uma operação recebida" & " do legado já foi enviada para o Banco Central através" & " da Entrada Manual do SLCC. " & vbCrLf & "A situação da operação será alterada para BAIXADO VIA CONTINGENCIA e " & "será retornado para o sistema legado de origem a situação de " & "liquidada para a operação "
                                                        
Private Const INFO_LIQUIDAR                 As String = "Utilizar Liquidar quando uma operação recebida do legado já foi enviada " & " para o Banco Central através da Contingência do sistema PK ou Internet." & vbCrLf & "A situação da operação será alterada para LIQUIDADO VIA CONTINGENCIA e será " & "disponibilizado o Lançamento em Conta Corrente se houver"

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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoStatusOperacao - Form_KeyDown", Me.Caption

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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoStatusOperacao - Form_Load", Me.Caption
    
End Sub

'' Obtém as propriedades da operação através da camada de controle de caso de uso
'' MIU, utilizando os seguintes métodos:   A8MIU.clsMiu.ObterMapaNavegacao
Private Sub flInicializar()

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

        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmAlteracaoStatusOperacao", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", lngCodigoErroNegocio

End Sub


'' Obtém as operações com status EmSer e Manual EmSer através de chamada ao método:
'' A8MIU.clsOperacao.ObterDetalheOperacao
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objOperacao        As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao        As A8MIU.clsOperacao
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

    Call flLimparLista
    
    strSelecaoFiltro = enumStatusOperacao.EmSer & ";" & _
                       enumStatusOperacao.ManualEmSer
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(strSelecaoFiltro, ";")) To UBound(Split(strSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", _
                                         "Status", Split(strSelecaoFiltro, ";")(lngCont))
    Next
    
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Grupo_SituacaoContingencia", ""
    fgAppendNode xmlDomFiltros, "Grupo_SituacaoContingencia", "SituacaoContingencia", enumIndicadorSimNao.sim
    
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
                
                .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                .SubItems(COL_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                .SubItems(COL_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
                .SubItems(COL_TITULO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
                
                If objDomNode.selectSingleNode("DT_VENC_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_VENCIMENTO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_VENC_ATIV").Text)
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
                
                If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                    .SubItems(COL_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text), "HH:MM")
                End If
                
                .SubItems(COL_CONTA_CEDENTE) = objDomNode.selectSingleNode("CO_CNTA_CEDT").Text
                .SubItems(COL_CONTA_CESSIONARIO) = objDomNode.selectSingleNode("CO_CNTA_CESS").Text
                
                If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
                    .SubItems(COL_DATA_OPERACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
                End If
                
                .SubItems(COL_SITU_PROC) = objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
                .SubItems(COL_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                .SubItems(COL_TIPO_LIQUIDACAO) = objDomNode.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text
                
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
        lblInfo.Top = .Height
        lblInfo.Width = Me.Width
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAlteracaoStatusOperacao = Nothing
End Sub

Private Sub lstOperacao_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

'' Evoca as funções flLiquidarContingencia e flBaixarContigencia, recarrega as
'' operações passívei de exclusão através de flCarregarLista, e fecha o objeto de
'' acordo com o botão selecionado
Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstOperacao_ColumnClick"
End Sub

Private Sub lstOperacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub tmrInfo_Timer()
    lblInfo.Visible = False
    tmrInfo.Enabled = False
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoLote                        As String

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
            
        Case "Baixar"
            
            strResultadoLote = flBaixarViaContingencia
            If strResultadoLote <> vbNullString Then
                Call flMostrarResultado(strResultadoLote, " baixados ")
            End If
            flCarregarLista
            
        Case "Liquidar"
            
            strResultadoLote = flLiquidarViaContingencia
            If strResultadoLote <> vbNullString Then
                Call flMostrarResultado(strResultadoLote, " liquidados ")
            End If
            flCarregarLista

        Case gstrAtualizar
            
            flCarregarLista
        
        Case gstrSair
            
            Unload Me
            
    End Select
    fgCursor
    Exit Sub

ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoStatusOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

'' Invoca o objeto frmResultOperacaoLote para exibir o resultado de um lote que
'' foi baixado/liquidado
Private Sub flMostrarResultado(ByVal pstrResultadoLote As String, _
                               ByVal pstrDescricaoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = pstrDescricaoOperacao
        .Resultado = pstrResultadoLote
        .Show vbModal
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'' Retorna um objeto contendo lotes das operações que serão baixadas/liquidadas.
Private Function flGerarLote() As MSXML2.DOMDocument40

Dim xmlDomLoteOperacao                      As MSXML2.DOMDocument40
Dim lngCont                                 As Long
Dim lngItensChecked                         As Long

Const POS_STATUS                            As Integer = 0
Const POS_DATA_ULTIMA_ATUALIZACAO           As Integer = 1

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão ---------------------------------------------------------------------------
    Set xmlDomLoteOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomLoteOperacao, "", "Repeat_Filtros", "")
    
    'Captura o filtro cumulativo OPERAÇÃO
    With lstOperacao.ListItems
        For lngCont = 1 To .Count
            If .Item(lngCont).Checked Then
                lngItensChecked = lngItensChecked + 1
                
                Call fgAppendNode(xmlDomLoteOperacao, "Repeat_Filtros", "Grupo_Lote", "")
                Call fgAppendNode(xmlDomLoteOperacao, _
                          "Grupo_Lote", "TipoConfirmacao", enumTipoConfirmacao.Operacao, "Repeat_Filtros")
                
                Call fgAppendNode(xmlDomLoteOperacao, _
                          "Grupo_Lote", "Operacao", Mid(.Item(lngCont).Key, 2), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacao, _
                          "Grupo_Lote", "Status", Split(.Item(lngCont).Tag, "|")(POS_STATUS), "Repeat_Filtros")
                          
                Call fgAppendNode(xmlDomLoteOperacao, _
                          "Grupo_Lote", "DHUltimaAtualizacao", Split(.Item(lngCont).Tag, "|")(POS_DATA_ULTIMA_ATUALIZACAO), "Repeat_Filtros")
                
            End If
        Next
    End With
    
    If lngItensChecked > 0 Then
        Set flGerarLote = xmlDomLoteOperacao
    Else
        Set flGerarLote = Nothing
    End If
    
    Set xmlDomLoteOperacao = Nothing

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flGerarLote", 0
End Function

'' Liquida o lote gerado por flGerarLote e retorna uma String contendo o resultado
'' do processamento, através de interação com a camada controladora de casos de
'' uso MIU , método:  A8MIU.clsOperacao.LiquidarViaContingencia
Private Function flLiquidarViaContingencia() As String

#If EnableSoap = 1 Then
    Dim objOperacao        As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao        As A8MIU.clsOperacao
#End If

Dim xmlDomLoteOperacao     As MSXML2.DOMDocument40
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set xmlDomLoteOperacao = flGerarLote
    If Not xmlDomLoteOperacao Is Nothing Then
        Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
        flLiquidarViaContingencia = objOperacao.LiquidarViaContingencia(xmlDomLoteOperacao.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objOperacao = Nothing
    End If
    
    Set xmlDomLoteOperacao = Nothing

Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlDomLoteOperacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiquidar", lngCodigoErroNegocio)

End Function

'' Baixa o lote gerado por flGerarLote e retorna uma String contendo o resultado
'' do processamento, através de interação com a camada controladora de casos de
'' uso MIU , método:  A8MIU.clsOperacao.BaixarViaContingencia
Private Function flBaixarViaContingencia() As String

#If EnableSoap = 1 Then
    Dim objOperacao        As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao        As A8MIU.clsOperacao
#End If

Dim xmlDomLoteOperacao     As MSXML2.DOMDocument40
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set xmlDomLoteOperacao = flGerarLote
    If Not xmlDomLoteOperacao Is Nothing Then
        Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
        flBaixarViaContingencia = objOperacao.BaixarViaContingencia(xmlDomLoteOperacao.xml, _
                                                                    vntCodErro, _
                                                                    vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objOperacao = Nothing
    End If
    
    Set xmlDomLoteOperacao = Nothing

Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlDomLoteOperacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiquidar", lngCodigoErroNegocio)

End Function

'' Limpa o conteúdo exibido no objeto
Private Sub flLimparLista()
    lstOperacao.ListItems.Clear
End Sub

Private Sub tlbFiltro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

    If X > tlbFiltro.Buttons(3).Left And X < tlbFiltro.Buttons(3).Left + tlbFiltro.Buttons(3).Width Then
        lblInfo.Caption = INFO_BAIXAR
        lblInfo.Visible = True
        tmrInfo.Enabled = False
        tmrInfo.Interval = 5000
        tmrInfo.Enabled = True
    ElseIf X > tlbFiltro.Buttons(4).Left And X < tlbFiltro.Buttons(4).Left + tlbFiltro.Buttons(4).Width Then
        lblInfo.Caption = INFO_LIQUIDAR
        lblInfo.Visible = True
        tmrInfo.Enabled = False
        tmrInfo.Interval = 5000
        tmrInfo.Enabled = True
    Else
        tmrInfo.Enabled = False
        lblInfo.Visible = False
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbFiltro_MouseMove"
End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

