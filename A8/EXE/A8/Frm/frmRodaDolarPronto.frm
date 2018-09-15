VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRodaDolarPronto 
   Caption         =   "Registro de Operações Roda de Dólar Pronto - BMC"
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
   Begin VB.Timer tmrRefresh 
      Interval        =   60000
      Left            =   5040
      Top             =   7080
   End
   Begin VB.TextBox txtTimer 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   4200
      TabIndex        =   4
      Text            =   "10"
      Top             =   7200
      Width           =   420
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4590
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
      ButtonWidth     =   2487
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela  "
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Liberar           "
            Key             =   "liberacao"
            Object.ToolTipText     =   "Liberar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Discordar       "
            Key             =   "discordancia"
            Object.ToolTipText     =   "Retornar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reg. Conting."
            Key             =   "contingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Regularizar    "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
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
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   7500
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
            Picture         =   "frmRodaDolarPronto.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodaDolarPronto.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.UpDown udTimer 
      Height          =   315
      Left            =   4621
      TabIndex        =   5
      Top             =   7200
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtTimer"
      BuddyDispid     =   196610
      OrigLeft        =   4860
      OrigTop         =   4470
      OrigRight       =   5100
      OrigBottom      =   4815
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "Intervalo para Refresh automático da tela (em minutos) :"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Width           =   3945
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
Attribute VB_Name = "frmRodaDolarPronto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pela conciliação de Registro de Operações Roda de Dólar - BMC,
'' através da camada de controle de caso de uso MIU.
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40

Private Const COL_NO_VEIC_LEGA              As Integer = 0
Private Const COL_VA_OPER_ATIV              As Integer = 1
Private Const COL_VA_FINC                   As Integer = 2
Private Const COL_VA_DIFE                   As Integer = 3
Private Const COL_DE_SITU_PROC              As Integer = 4
Private Const COL_TP_ACAO_MESG_SPB_EXEC     As Integer = 5
Private Const COL_CO_SISB_COTR              As Integer = 6
Private Const COL_IN_OPER_DEBT_CRED         As Integer = 7
Private Const COL_CO_PRAC                   As Integer = 8
Private Const COL_CO_MOED_ESTR              As Integer = 9
Private Const COL_VA_MOED_ESTR              As Integer = 10
Private Const COL_PE_TAXA_NEGO              As Integer = 11
Private Const COL_DT_LIQU_OPER_ATIV         As Integer = 12
Private Const COL_DH_RECB_ENVI_MESG_SPB     As Integer = 13

Private Const KEY_NU_CTRL_IF                As Integer = 1
Private Const KEY_DH_REGT_MESG_SPB          As Integer = 2
Private Const KEY_NU_SEQU_CNTR_REPE         As Integer = 3
Private Const KEY_CO_ULTI_SITU_PROC         As Integer = 4
Private Const KEY_TP_ACAO_MESG_SPB_EXEC     As Integer = 5

Private Const strFuncionalidade             As String = "frmCompromissadaGenerica"

Private intAcaoProcessamento                As enumAcaoConciliacao

Private lngPerfil                           As Long
Private blnDummyH                           As Boolean

'Controla o timer de refresh da tela
Private intContMinutos                      As Integer

Private lngIndexClassifList                 As Long

'Calcular a diferença dos valores das operações e mensagens SPB.

Private Sub flCalcularDiferencasListView()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

On Error GoTo ErrorHandler
    
    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VA_OPER_ATIV)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VA_FINC)))
            
            .SubItems(COL_VA_DIFE) = fgVlrXml_To_Interface(dblValorMensagem - dblValorOperacao)
            
            If dblValorMensagem - dblValorOperacao <> 0 Then
                .ListSubItems(COL_VA_DIFE).ForeColor = vbRed
            End If
        End With
    Next

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListView", 0)

End Sub

'' Carregar mensagens SPB e preencher a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método A8MIU.clsMensagem.ObterDetalheMensagem

Private Sub flCarregarListaNetMensagens(ByVal pstrFiltro As String)

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
    strRetLeitura = objMensagem.ObterDetalheMensagem(pstrFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetLeitura.loadXML(strRetLeitura)
        
        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            If objDomNode.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text <> enumTipoAcao.EnviadaBMC0012 Then
                strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                                 "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                                 "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & _
                                 "|" & objDomNode.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text
                                 
                With lvwMensagem.ListItems.Add(, strListItemKey)
                    
                    .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                    
                    .SubItems(COL_VA_FINC) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    .SubItems(COL_VA_DIFE) = " "
                    .SubItems(COL_DE_SITU_PROC) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_TP_ACAO_MESG_SPB_EXEC) = fgDescricaoTipoAcao(Val(objDomNode.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text))
                    .SubItems(COL_CO_SISB_COTR) = objDomNode.selectSingleNode("CO_SISB_COTR").Text
                    .SubItems(COL_CO_PRAC) = objDomNode.selectSingleNode("CO_PRAC").Text
                    .SubItems(COL_CO_MOED_ESTR) = objDomNode.selectSingleNode("CO_MOED_ESTR").Text
                    .SubItems(COL_VA_MOED_ESTR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_MOED_ESTR").Text)
                    .SubItems(COL_PE_TAXA_NEGO) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("PE_TAXA_NEGO").Text)
                    .SubItems(COL_DT_LIQU_OPER_ATIV) = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR").Text)
                    .SubItems(COL_DH_RECB_ENVI_MESG_SPB) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text)
                
                    If objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.Conciliada Or _
                       objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.ConciliadaAutomatica Then
                        .SubItems(COL_VA_OPER_ATIV) = .SubItems(COL_VA_FINC)
                    Else
                        .SubItems(COL_VA_OPER_ATIV) = " "
                    End If
                
                    If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                        .SubItems(COL_IN_OPER_DEBT_CRED) = "Compra"
                    Else
                        .SubItems(COL_IN_OPER_DEBT_CRED) = "Venda"
                    End If
                End With
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifList, True)
    
    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetMensagens", 0)

End Sub

'' Carrega as propriedades necessárias a interface, através da
'' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

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

'Configurar o list view de mensagens
Private Sub flInicializarLvwMensagem()

On Error GoTo ErrorHandler
    
    With Me.lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Veículo Legal", 4050
        .Add , , "Valor Sistema Origem", 1950, lvwColumnRight
        .Add , , "Valor Mensagem", 1815, lvwColumnRight
        .Add , , "Diferença", 1500, lvwColumnRight
        .Add , , "Situação", 2730
        .Add , , "Ação", 2685
        .Add , , "Código SISBACEN Corretora", 2250
        .Add , , "Tipo Operação BMC", 1695
        .Add , , "Código Praça IF", 1350
        .Add , , "Código Moeda", 1245
        .Add , , "Valor Moeda Estrangeira", 1920
        .Add , , "Taxa Câmbio", 1140
        .Add , , "Data Liquidação", 1400
        .Add , , "Data Hora BMC", 2040
    End With
    
Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwMensagem", 0

End Sub

'Montar o xml de filtro para pesquisa

Private Function flMontarXMLFiltroPesquisa() As String
    
Dim xmlFiltros                              As MSXML2.DOMDocument40
    
On Error GoTo ErrorHandler
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.BMC)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.Conciliada)
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConciliadaAutomatica)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "BMC0011")
        
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
                                                   "NU_CTRL_IF", _
                                                   Split(.Key, "|")(KEY_NU_CTRL_IF))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DH_REGT_MESG_SPB", _
                                                   Split(.Key, "|")(KEY_DH_REGT_MESG_SPB))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "NU_SEQU_CNTR_REPE", _
                                                   Split(.Key, "|")(KEY_NU_SEQU_CNTR_REPE))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_ULTI_SITU_PROC", _
                                                   Split(.Key, "|")(KEY_CO_ULTI_SITU_PROC))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "TP_ACAO_MESG_SPB_EXEC", _
                                                   Split(.Key, "|")(KEY_TP_ACAO_MESG_SPB_EXEC))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_MESG_SPB", _
                                                   "BMC0011")

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
    
    lvwMensagem.ListItems.Clear

    If Me.cboEmpresa.ListIndex = -1 Or Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If
    
    fgCursor True
    
    strDocFiltros = flMontarXMLFiltroPesquisa
    Call flCarregarListaNetMensagens(strDocFiltros)
    Call flCalcularDiferencasListView
    
    fgCursor
    
Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPesquisar", 0)

End Sub

'' Executar o processamento efetuado através da camada controladora de casos de uso
'' MIU, método A8MIU.clsOperacaoMensagem.ProcessarCompromissadaGenerica
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
        strXMLRetorno = objOperacaoMensagem.ProcessarLoteLiberacaoRodaDolarPronto(intAcaoProcessamento, _
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

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString

Private Function flValidarItensProcessamento() As String

Dim objListItem                             As MSComctlLib.ListItem
Dim intStatus                               As Integer
Dim intAcao                                 As Integer

    If fgItemsCheckedListView(Me.lvwMensagem) = 0 Then
        flValidarItensProcessamento = "Selecione pelo menos um item da lista, antes de prosseguir com a operação desejada."
        Exit Function
    End If

    For Each objListItem In lvwMensagem.ListItems
        If objListItem.Checked Then
            
            intStatus = Split(objListItem.Key, "|")(KEY_CO_ULTI_SITU_PROC)
            intAcao = Split(objListItem.Key, "|")(KEY_TP_ACAO_MESG_SPB_EXEC)
            
            Select Case intAcaoProcessamento
                Case enumAcaoConciliacao.AdmAreaLiberar
                    If intStatus <> enumStatusMensagem.Conciliada And _
                       intStatus <> enumStatusMensagem.ConciliadaAutomatica Then
                        flValidarItensProcessamento = "Liberação só é permitida para mensagens com situação Conciliada ou Conciliada Automática."
                        Exit Function
                    End If
                    
                    If intAcao = enumTipoAcao.RegistroContingencia Then
                        flValidarItensProcessamento = "Registro em Contingência já foi enviado para esta mensagem."
                        Exit Function
                    End If
                    
                Case enumAcaoConciliacao.AdmAreaRejeitar
                    If intStatus <> enumStatusMensagem.AConciliar Then
                        flValidarItensProcessamento = "Discordância só é permitida para mensagens com situação A Conciliar."
                        Exit Function
                    End If
                    
                    If intAcao = enumTipoAcao.RegistroContingencia Then
                        flValidarItensProcessamento = "Registro em Contingência já foi enviado para esta mensagem."
                        Exit Function
                    ElseIf intAcao = enumTipoAcao.DiscordanciaAdmBO Then
                        flValidarItensProcessamento = "Discordância Adm. BO já foi enviada para esta mensagem."
                        Exit Function
                    End If
                    
                Case enumAcaoConciliacao.AdmAreaLiberarContingencia
                    If intStatus <> enumStatusMensagem.AConciliar Then
                        flValidarItensProcessamento = "Registro em Contingência só é permitido para mensagens com situação A Conciliar."
                        Exit Function
                    End If
                    
                    If intAcao = enumTipoAcao.RegistroContingencia Then
                        flValidarItensProcessamento = "Registro em Contingência já foi enviado para esta mensagem."
                        Exit Function
                    ElseIf intAcao = enumTipoAcao.DiscordanciaAdmBO Then
                        flValidarItensProcessamento = "Discordância Adm. BO já foi enviada para esta mensagem."
                        Exit Function
                    End If
                    
                Case enumAcaoConciliacao.AdmAreaRegularizar
                    If intStatus <> enumStatusMensagem.Conciliada And _
                       intStatus <> enumStatusMensagem.ConciliadaAutomatica Then
                        flValidarItensProcessamento = "Regularização só é permitida para mensagens com situação Conciliada ou Conciliada Automática."
                        Exit Function
                    End If
                    
                    If intAcao <> enumTipoAcao.RegistroContingencia Then
                        flValidarItensProcessamento = "Registro em Contingência não foi enviado para esta mensagem. Regularização não permitida."
                        Exit Function
                    End If
                    
            End Select
        
        End If
    Next
    
End Function

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
    fgCursor
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .lvwMensagem.Top = .cboEmpresa.Top + .cboEmpresa.Height + 120
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = .Height - .lvwMensagem.Top - 1150
        .lvwMensagem.Width = .Width - 240
        lblTimer.Top = .lvwMensagem.Top + .lvwMensagem.Height + 100
        lblTimer.Left = .lvwMensagem.Left
        txtTimer.Top = lblTimer.Top - 50
        udTimer.Top = lblTimer.Top - 50
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlRetornoErro = Nothing
    Set frmRodaDolarPronto = Nothing

End Sub

Private Sub lvwMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwMensagem, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ColumnClick", Me.Caption

End Sub

Private Sub lvwMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    If Item.Checked Then
        If Trim$(Item.SubItems(COL_VA_FINC)) = vbNullString Then
            frmMural.Display = "Seleção do item não permitida. Valor de mensagem não encontrado."
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

Private Sub lvwMensagem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_MouseDown", Me.Caption

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
            
        Case "liberacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar
            
        Case "discordancia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaRejeitar
            
        Case "contingencia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberarContingencia
            
        Case "regularizacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaRegularizar
            
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
        
        Select Case intAcaoProcessamento
            Case enumAcaoConciliacao.AdmAreaLiberar
                If MsgBox("Confirma a Liberação do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    GoTo ExitSub
                End If
            Case enumAcaoConciliacao.AdmAreaRejeitar
                If MsgBox("Confirma a Discordância do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    GoTo ExitSub
                End If
            Case enumAcaoConciliacao.AdmAreaLiberarContingencia
                If MsgBox("Confirma o Registro em Contingência do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    GoTo ExitSub
                End If
            Case enumAcaoConciliacao.AdmAreaRegularizar
                If MsgBox("Confirma a Regularização do(s) item(s) selecionado(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    GoTo ExitSub
                End If
        End Select
        
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
Private Sub tmrRefresh_Timer()

On Error GoTo ErrorHandler

    If Not IsNumeric(txtTimer.Text) Then Exit Sub
    
    If CLng(txtTimer.Text) = 0 Then Exit Sub
    
    If fgVerificaJanelaVerificacao() Then Exit Sub
    
    fgCursor True

    intContMinutos = intContMinutos + 1
    
    If intContMinutos >= txtTimer.Text Then

        Call flPesquisar

        intContMinutos = 0
    End If

    fgCursor False

Exit Sub
ErrorHandler:
    
    fgCursor False
    
    fgRaiseError App.EXEName, TypeName(Me), "tmrRefresh_Timer", 0

End Sub



