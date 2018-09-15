VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroWorkflow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parametrização de Workflow"
   ClientHeight    =   7950
   ClientLeft      =   2610
   ClientTop       =   1485
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10065
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   6855
   End
   Begin VB.Frame fraDetalhe 
      Caption         =   "Detalhe"
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
      ForeColor       =   &H00800000&
      Height          =   4005
      Left            =   120
      TabIndex        =   1
      Top             =   3540
      Width           =   9825
      Begin VB.CheckBox chkIncondicionalSaldo 
         Caption         =   "Incondicional a Saldo"
         Height          =   255
         Left            =   7200
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame fraExcecao 
         Caption         =   "Exceção quando Função Automática"
         Enabled         =   0   'False
         Height          =   2715
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   9600
         Begin VB.ComboBox cboGrupoUsuario 
            Height          =   315
            Left            =   6510
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   540
            Width           =   1935
         End
         Begin VB.ComboBox cboLocalLiqu 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   3075
         End
         Begin VB.ComboBox cboSistema 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   540
            Width           =   3165
         End
         Begin MSComctlLib.Toolbar tlbAplicar 
            Height          =   330
            Left            =   8475
            TabIndex        =   16
            Top             =   525
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            ButtonWidth     =   1958
            ButtonHeight    =   582
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgIcons"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Aplicar"
                  Key             =   "Aplicar"
                  ImageIndex      =   18
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lstExcecao 
            Height          =   1605
            Left            =   105
            TabIndex        =   17
            Top             =   915
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   2831
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Excluir"
               Object.Width           =   1129
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Sistema"
               Object.Width           =   5028
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Local Liquidação"
               Object.Width           =   5028
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Grupo Usuário"
               Object.Width           =   5028
            EndProperty
         End
         Begin VB.Label lblWorkflow 
            AutoSize        =   -1  'True
            Caption         =   "Grupo Usuário"
            Height          =   195
            Left            =   6510
            TabIndex        =   15
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label lblWokflow 
            AutoSize        =   -1  'True
            Caption         =   "Local Liquidação"
            Height          =   195
            Index           =   6
            Left            =   3360
            TabIndex        =   13
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblWokflow 
            AutoSize        =   -1  'True
            Caption         =   "Sistema"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.CheckBox chkFuncaoAutomatica 
         Caption         =   "Indicador Função Automática"
         Height          =   255
         Left            =   7200
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblFuncaoSistema 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4050
         TabIndex        =   19
         Top             =   570
         Width           =   3075
      End
      Begin VB.Label lblTipoOperacao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   570
         Width           =   3825
      End
      Begin VB.Label lblWokflow 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Operação"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblWokflow 
         AutoSize        =   -1  'True
         Caption         =   "Função Sistema"
         Height          =   195
         Index           =   4
         Left            =   4080
         TabIndex        =   2
         Top             =   330
         Width           =   1140
      End
   End
   Begin MSComctlLib.ListView lstRegra 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo Operação"
         Object.Width           =   6085
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Função Sistema"
         Object.Width           =   3493
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Indicador Função Automática"
         Object.Width           =   4128
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Incondicional Saldo"
         Object.Width           =   2999
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   120
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":0112
            Key             =   "AtualizarExcecao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":5922
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":5F56
            Key             =   "Nova"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroRegra.frx":6270
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   7020
      TabIndex        =   4
      Top             =   7590
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   582
      ButtonWidth     =   1958
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblWokflow 
      AutoSize        =   -1  'True
      Caption         =   "Parâmetro Função Sistema"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   1905
   End
   Begin VB.Label lblWokflow 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCadastroWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:08
'-------------------------------------------------
'' Objeto reponsável pelo cadastramento do Workflow através da camada de controle
'' de caso de uso MIU.
''
'' Classes consideradas especificamente de destino
''    A8MIU.clsMIU
''    A8MIU.clsCadastroWorkflow
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private xmlLerTodos                         As MSXML2.DOMDocument40
Private objLstItem                          As ListItem

Private Const strFuncionalidade             As String = "frmCadastroWorkflow"

Private strCodTipoOperacao                  As String
Private strCodFuncaoSistema                 As String
Private strKey                              As String

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString

Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If cboEmpresa.ListIndex < 0 Then
        flValidarCampos = "Selecione a Empresa do Cadastro de Workflow."
        cboEmpresa.SetFocus
        Exit Function
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'' Carrega as propriedades necessárias a interface frmCadastroRegra, através da
'' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    Set objMIU = Nothing
    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    
    'Carregar Locais de Liquidação
    xmlLerTodos.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_LocalLiquidacao").xml
    xmlLerTodos.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
    xmlLerTodos.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("TP_SEGR").Text = "S"
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlLerTodos.loadXML objMIU.Executar(xmlLerTodos.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    Call fgCarregarCombos(Me.cboLocalLiqu, xmlLerTodos, "LocalLiquidacao", "CO_LOCA_LIQU", "SG_LOCA_LIQU")
    
    'Carregar Grupos de Usuário
    xmlLerTodos.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario").xml
    xmlLerTodos.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
    xmlLerTodos.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("TP_SEGR").Text = "S"
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlLerTodos.loadXML objMIU.Executar(xmlLerTodos.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    Call fgCarregarCombos(Me.cboGrupoUsuario, xmlLerTodos, "GrupoUsuario", "CO_GRUP_USUA", "NO_GRUP_USUA")
    
    Call fgCursor(False)
    Set objMIU = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objMIU = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'' Carrega os cadastros já existentes e preenche a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método  A8MIU.clsCadastroWorkflow.LerTodos

Private Sub flCarregarParametrizacao()

#If EnableSoap = 1 Then
    Dim objCadastroWorkflow     As MSSOAPLib30.SoapClient30
#Else
    Dim objCadastroWorkflow     As A8MIU.clsCadastroWorkflow
#End If

Dim xmlLerTodosParam            As MSXML2.DOMDocument40

Dim strLerTodos                 As String
Dim xmlDomNode                  As IXMLDOMNode
Dim objListItem                 As ListItem
Dim strLstKey                   As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    Set xmlLerTodosParam = CreateObject("MSXML2.DOMDocument.4.0")
    Set objCadastroWorkflow = fgCriarObjetoMIU("A8MIU.clsCadastroWorkflow")
    
    Me.lstRegra.ListItems.Clear
    
    strLerTodos = objCadastroWorkflow.LerTodos(CInt(fgObterCodigoCombo(cboEmpresa.Text)), vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strLerTodos <> vbNullString Then
        
        If Not xmlLerTodosParam.loadXML(strLerTodos) Then
            Call fgErroLoadXML(xmlLerTodosParam, App.EXEName, Me.Name, "flCarregarParametrizacao")
        End If

        For Each xmlDomNode In xmlLerTodosParam.selectNodes("//Repeat_ParamWorkflow/*")
            
            strKey = "K" & fgCompletaString(xmlDomNode.selectSingleNode("TP_OPER").Text, "0", 3, True) & fgCompletaString(xmlDomNode.selectSingleNode("CO_FCAO_SIST").Text, "0", 3, True)
            
            Set objListItem = lstRegra.ListItems.Add(, strKey, xmlDomNode.selectSingleNode("NO_TIPO_OPER").Text)
            
            objListItem.Tag = xmlDomNode.xml
            
            Select Case CInt(xmlDomNode.selectSingleNode("CO_FCAO_SIST").Text)
                Case enumFuncaoSistema.Conciliar
                    objListItem.SubItems(1) = "Conciliação"
                Case enumFuncaoSistema.Pagar
                    objListItem.SubItems(1) = "Conciliação LBTR"
                Case enumFuncaoSistema.Confirmar
                    objListItem.SubItems(1) = "Confirmação"
                Case enumFuncaoSistema.Liberar
                    objListItem.SubItems(1) = "Liberação"
                Case enumFuncaoSistema.LiberarPagamento
                    objListItem.SubItems(1) = "Liberação LBTR"
                Case enumFuncaoSistema.ConfirmarReativacao
                    objListItem.SubItems(1) = "Confirmação Reativação"
                Case enumFuncaoSistema.LiberarReativacao
                    objListItem.SubItems(1) = "Liberação Reativação"
                Case enumFuncaoSistema.IntegracaoCC
                    objListItem.SubItems(1) = "Integração Conta Corrente"
            
                    Select Case xmlDomNode.selectSingleNode("TP_COND_SALD").Text
                        Case enumIndicadorSimNao.Sim
                            objListItem.SubItems(3) = "Sim"
                        Case enumIndicadorSimNao.Nao
                            objListItem.SubItems(3) = "Não"
                    End Select
                Case enumFuncaoSistema.LiberarCAM0054
                    objListItem.SubItems(1) = "Liberação CAM0054"
            
            End Select
                
            Select Case xmlDomNode.selectSingleNode("IN_FCAO_SIST_AUTM").Text
            Case enumIndicadorSimNao.Sim
                objListItem.SubItems(2) = "Sim"
            Case enumIndicadorSimNao.Nao
                objListItem.SubItems(2) = "Não"
            End Select
            
        Next
    End If

    Call flLimparDetalhe
    
    Set objCadastroWorkflow = Nothing
    Set xmlLerTodosParam = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objCadastroWorkflow = Nothing
    Set xmlLerTodosParam = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarParametrizacao", 0

End Sub

' Preenche o combo de acordo com o documento XML

Private Sub flCarregarComboSistema(ByVal plngEmpresa As Long)

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    cboSistema.Clear
    
    For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_Sistema/Grupo_Sistema[CO_EMPR = '" & plngEmpresa & "']")
        cboSistema.AddItem objDomNode.selectSingleNode("SG_SIST").Text & " - " & _
                            objDomNode.selectSingleNode("NO_SIST").Text
    Next
    
    cboSistema.ListIndex = -1
    Set objDomNode = Nothing
    Call fgCursor(False)
    

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flCarregarComboSistema", 0

End Sub

'Limpar as informações de Exceção para uma nova inclusão

Private Sub flLimparExcessoes()

Dim objListItem                             As ListItem

On Error GoTo ErrorHandler

    With Me
        For Each objListItem In .lstExcecao.ListItems
            If objListItem.Checked = False Then
               objListItem.Checked = True
               objListItem.Tag = enumTipoOperacao.Excluir
            End If
        Next
        .cboGrupoUsuario.ListIndex = -1
        .cboLocalLiqu.ListIndex = -1
        .cboSistema.ListIndex = -1
        .fraExcecao.Enabled = False
    End With
    
    Set objListItem = Nothing

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparExcessoes", 0

End Sub

'Limpar os detalhes para uma nova inclusão
Private Sub flLimparDetalhe()

On Error GoTo ErrorHandler

    With Me
        .cboGrupoUsuario.ListIndex = -1
        .cboLocalLiqu.ListIndex = -1
        .cboSistema.ListIndex = -1
        .chkFuncaoAutomatica.value = vbUnchecked
        .chkIncondicionalSaldo.value = vbUnchecked
        .lblFuncaoSistema.Caption = vbNullString
        .lblTipoOperacao.Caption = vbNullString
        .lstExcecao.ListItems.Clear
        
        .fraDetalhe.Enabled = False
        .fraExcecao.Enabled = False
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparDetalhe", 0

End Sub

'' Carrega os detalhes da parametrização já existentes e preenche a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método  A8MIU.clsCadastroWorkflow.Ler
Private Sub flObterDetalheParametrizacao(pobjItemSel As ListItem)

#If EnableSoap = 1 Then
    Dim objCadastroWorkflow     As MSSOAPLib30.SoapClient30
#Else
    Dim objCadastroWorkflow     As A8MIU.clsCadastroWorkflow
#End If

Dim xmlChaveLer                 As MSXML2.DOMDocument40
Dim strLer                      As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    If pobjItemSel Is Nothing Then Exit Sub

    Set objCadastroWorkflow = fgCriarObjetoMIU("A8MIU.clsCadastroWorkflow")
    Set xmlChaveLer = CreateObject("MSXML2.DOMDocument.4.0")
    
    strKey = pobjItemSel.Key
    
    Call xmlChaveLer.loadXML(pobjItemSel.Tag)
    
    strCodTipoOperacao = xmlChaveLer.documentElement.selectSingleNode("TP_OPER").Text
    strCodFuncaoSistema = xmlChaveLer.documentElement.selectSingleNode("CO_FCAO_SIST").Text
    
    strLer = objCadastroWorkflow.Ler(CInt(strCodTipoOperacao), _
                                     CInt(strCodFuncaoSistema), _
                                     CInt(fgObterCodigoCombo(cboEmpresa.Text)), _
                                     vntCodErro, _
                                     vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If strLer <> vbNullString Then
        If Not xmlLer.loadXML(strLer) Then
            Call fgErroLoadXML(xmlLer, App.EXEName, Me.Name, "flCarregarParametrizacao")
        End If

        lblTipoOperacao.Caption = xmlLer.documentElement.selectSingleNode("NO_TIPO_OPER").Text
        
        Select Case CInt(xmlLer.documentElement.selectSingleNode("CO_FCAO_SIST").Text)
        Case enumFuncaoSistema.Conciliar
            lblFuncaoSistema.Caption = "Conciliação"
        Case enumFuncaoSistema.Pagar
            lblFuncaoSistema.Caption = "Conciliação LBTR"
        Case enumFuncaoSistema.Confirmar
            lblFuncaoSistema.Caption = "Confirmação"
        Case enumFuncaoSistema.Liberar
            lblFuncaoSistema.Caption = "Liberação"
        Case enumFuncaoSistema.LiberarPagamento
            lblFuncaoSistema.Caption = "Liberação LBTR"
        Case enumFuncaoSistema.ConfirmarReativacao
            lblFuncaoSistema.Caption = "Confirmação Reativação"
        Case enumFuncaoSistema.LiberarReativacao
            lblFuncaoSistema.Caption = "Liberação Reativação"
        Case enumFuncaoSistema.IntegracaoCC
            lblFuncaoSistema.Caption = "Integração Conta Corrente"
        End Select
                
        Select Case CInt(xmlLer.documentElement.selectSingleNode("IN_FCAO_SIST_AUTM").Text)
        Case enumIndicadorSimNao.Sim
            chkFuncaoAutomatica.value = vbChecked
        Case enumIndicadorSimNao.Nao
            chkFuncaoAutomatica.value = vbUnchecked
        End Select
        
        'Habilita o check de Incondional a Saldo somente quando o tipo de operação for integração conta corrente
        Select Case CInt(xmlLer.documentElement.selectSingleNode("CO_FCAO_SIST").Text)
        Case enumFuncaoSistema.IntegracaoCC
            If CInt(xmlLer.documentElement.selectSingleNode("TP_COND_SALD").Text) = enumIndicadorSimNao.Sim Then
                chkIncondicionalSaldo.value = vbChecked
            Else
                chkIncondicionalSaldo.value = vbUnchecked
            End If
            
            chkIncondicionalSaldo.Enabled = True
        Case Else
            chkIncondicionalSaldo.Enabled = False
            chkIncondicionalSaldo.value = vbUnchecked
        End Select
        
        fraDetalhe.Enabled = True
    Else
        fraDetalhe.Enabled = False
    End If

    Set objCadastroWorkflow = Nothing
    Set xmlChaveLer = Nothing
    
Exit Sub
ErrorHandler:
    
    Set objCadastroWorkflow = Nothing
    Set xmlChaveLer = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flObterDetalheParametrizacao", 0

End Sub

'' Carrega as Exceções da parametrização já existentes e preenche a interface com os mesmos,
'' através da classe controladora de caso de uso MIU, método  A8MIU.clsCadastroWorkflow.Ler

Private Sub flObterExcecoesParametrizacao()

#If EnableSoap = 1 Then
    Dim objCadastroWorkflow     As MSSOAPLib30.SoapClient30
#Else
    Dim objCadastroWorkflow     As A8MIU.clsCadastroWorkflow
#End If

Dim strExcecoes                 As String

Dim xmlDomNode                  As IXMLDOMNode
Dim objListItem                 As ListItem
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    Set objCadastroWorkflow = fgCriarObjetoMIU("A8MIU.clsCadastroWorkflow")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    strExcecoes = objCadastroWorkflow.ObterExcecoesParametrizacao(CInt(strCodTipoOperacao), _
                                                                  CInt(strCodFuncaoSistema), _
                                                                  CInt(fgObterCodigoCombo(cboEmpresa.Text)), _
                                                                  vntCodErro, _
                                                                  vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    lstExcecao.ListItems.Clear
    
    If strExcecoes <> vbNullString Then
        If Not xmlLerTodos.loadXML(strExcecoes) Then
            Call fgErroLoadXML(xmlLerTodos, App.EXEName, Me.Name, "flObterExcecoesParametrizacao")
        End If

        For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_ParamWorkflow/*")
            'O 'Key' está separado por ";" pois existe sistemas que começam com a letra "K", o que
            'inviabilizava a separação dos código por meio da função Split()
            
            Set objListItem = lstExcecao.ListItems.Add(, ";" & strCodTipoOperacao & _
                                                         ";" & strCodFuncaoSistema & _
                                                         ";" & fgObterCodigoCombo(cboEmpresa.Text) & _
                                                         ";" & xmlDomNode.selectSingleNode("SG_SIST").Text & _
                                                         ";" & xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text & _
                                                         ";" & xmlDomNode.selectSingleNode("CO_GRUP_USUA").Text & _
                                                         ";" & xmlDomNode.selectSingleNode("SG_LOCA_LIQU").Text)
            
            objListItem.SubItems(1) = xmlDomNode.selectSingleNode("SG_SIST").Text
            objListItem.SubItems(2) = xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text & " - " & xmlDomNode.selectSingleNode("SG_LOCA_LIQU").Text
            objListItem.SubItems(3) = xmlDomNode.selectSingleNode("CO_GRUP_USUA").Text & " - " & xmlDomNode.selectSingleNode("NO_GRUP_USUA").Text
            
            objListItem.Tag = enumTipoOperacao.None
        Next
    End If

    Set objCadastroWorkflow = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objCadastroWorkflow = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flObterExcecoesParametrizacao", 0

End Sub

'' Salva as alterações efetuadas através da camada controladora de casos de uso
'' MIU, método A8MIU.clsMIU.Executar

Private Function flSalvarTodos() As Boolean

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim xmlSalvar               As MSXML2.DOMDocument40
Dim xmlDomNode              As MSXML2.IXMLDOMNode

Dim objListItem             As ListItem
Dim blnPossuiRegraExcecao   As Boolean
Dim arrChaves()             As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If lblTipoOperacao.Caption = "" Then
       Exit Function
    End If
    
    With xmlLer.documentElement
        .selectSingleNode("/ParamWorkflow/@Operacao").Text = gstrOperAlterar
        .selectSingleNode("IN_FCAO_SIST_AUTM").Text = IIf(chkFuncaoAutomatica.value = vbChecked, _
                                                          enumIndicadorSimNao.Sim, _
                                                          enumIndicadorSimNao.Nao)
        '.selectSingleNode("CO_FCAO_SIST ").Text = lstRegra.SelectedItem.Key
   
        .selectSingleNode("TP_COND_SALD").Text = IIf(chkIncondicionalSaldo.value = vbChecked, _
                                                          enumIndicadorSimNao.Sim, _
                                                          enumIndicadorSimNao.Nao)
    End With
    
    'Verifica se existe Regras de Exceção
    If lstExcecao.ListItems.Count > 0 Then
        blnPossuiRegraExcecao = True
        
        Set xmlSalvar = CreateObject("MSXML2.DOMDocument.4.0")
    
        Call fgAppendNode(xmlSalvar, "", "Repeat_ParamExcecoes", "")
        Call fgAppendAttribute(xmlSalvar, "Repeat_ParamExcecoes", "Operacao", "Incluir")
    
        For Each objListItem In lstExcecao.ListItems
            If objListItem.Tag = enumTipoOperacao.Incluir Then
                Set xmlDomNode = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParamExcecoes")
    
                xmlDomNode.selectSingleNode("@Operacao").Text = "Incluir"
    
                arrChaves = Split(objListItem.Key, ";")
                xmlDomNode.selectSingleNode("TP_OPER").Text = arrChaves(1)
                xmlDomNode.selectSingleNode("CO_FCAO_SIST").Text = arrChaves(2)
                xmlDomNode.selectSingleNode("CO_EMPR").Text = arrChaves(3)
                xmlDomNode.selectSingleNode("SG_SIST").Text = arrChaves(4)
                xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text = arrChaves(5)
                xmlDomNode.selectSingleNode("CO_GRUP_USUA").Text = arrChaves(6)
    
                Call fgAppendXML(xmlSalvar, "Repeat_ParamExcecoes", xmlDomNode.xml)
            ElseIf objListItem.Tag = enumTipoOperacao.Excluir Then
                arrChaves = Split(objListItem.Key, ";")
                For Each xmlDomNode In xmlLerTodos.documentElement.selectNodes("//Repeat_ParamWorkflow/*")
                    If xmlDomNode.selectSingleNode("TP_OPER").Text = arrChaves(1) And _
                       xmlDomNode.selectSingleNode("CO_FCAO_SIST").Text = arrChaves(2) And _
                       xmlDomNode.selectSingleNode("CO_EMPR").Text = arrChaves(3) And _
                       xmlDomNode.selectSingleNode("SG_SIST").Text = arrChaves(4) And _
                       xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text = arrChaves(5) And _
                       xmlDomNode.selectSingleNode("CO_GRUP_USUA").Text = arrChaves(6) Then
                        
                        xmlDomNode.selectSingleNode("@Operacao").Text = gstrOperExcluir
                        Call fgAppendXML(xmlSalvar, "Repeat_ParamExcecoes", xmlDomNode.xml)
                        Exit For
                    End If
                Next
            End If
        Next
    End If
    
    If blnPossuiRegraExcecao Then
        Call fgAppendXML(xmlLer, "ParamWorkflow", xmlSalvar.xml)
    End If
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

    Set xmlSalvar = Nothing
    flSalvarTodos = True
    
    Exit Function

ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlSalvar = Nothing
    Call fgCursor(False)
    flSalvarTodos = False
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvarTodos", 0

End Function

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString

Private Function flValidarCamposExcecao() As String

On Error GoTo ErrorHandler

    If cboSistema.Text = vbNullString Then
        flValidarCamposExcecao = "Selecione o Sistema."
        cboSistema.SetFocus
        Exit Function
    End If
    
    If cboLocalLiqu.Text = vbNullString Then
        flValidarCamposExcecao = "Selecione o Local de Liquidação."
        cboLocalLiqu.SetFocus
        Exit Function
    End If
    
    If cboGrupoUsuario.Text = vbNullString Then
        flValidarCamposExcecao = "Selecione o Grupo de Usuário."
        cboGrupoUsuario.SetFocus
        Exit Function
    End If
    
    flValidarCamposExcecao = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCamposExcecao", 0

End Function

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    fgCursor True
    
    If cboEmpresa.Text <> vbNullString Then
        Call flCarregarComboSistema(fgObterCodigoCombo(cboEmpresa.Text))
        Call flCarregarParametrizacao
    End If
    
    fgCursor
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
End Sub

Private Sub chkFuncaoAutomatica_Click()

On Error GoTo ErrorHandler
   
    fgCursor True
    
    If chkFuncaoAutomatica.value = vbChecked Then
        Me.fraExcecao.Enabled = True
        Call flObterExcecoesParametrizacao
    Else
        Me.fraExcecao.Enabled = False
        flLimparExcessoes
    End If
    
    fgCursor
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    Call fgCursor(True)
    Set Me.Icon = mdiLQS.Icon
    Call fgCenterMe(Me)
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call flInicializar
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCadastroWorkflow = Nothing
End Sub

Private Sub lstExcecao_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstExcecao_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    If Item.Checked Then
        If Item.Tag = enumTipoOperacao.None Then
            Item.Tag = enumTipoOperacao.Excluir
        Else
            lstExcecao.ListItems.Remove Item.Index
        End If
    Else
        Item.Tag = enumTipoOperacao.None
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstExcecao_ItemCheck"

End Sub

Private Sub lstRegra_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstRegra_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Set objLstItem = Item
    fgCursor True
    Call flObterDetalheParametrizacao(Item)
    Call flObterExcecoesParametrizacao
    fgCursor False
    
    Exit Sub
    
ErrorHandler:
    
    fgCursor False
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
    Call flLimparDetalhe
    Call flObterDetalheParametrizacao(lstRegra.SelectedItem)
    Call flObterExcecoesParametrizacao
    
End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim objListItem                             As ListItem
Dim strValidarCampos                        As String
Dim strItemKey                              As String

On Error GoTo ErrorHandler

    If lstRegra.SelectedItem Is Nothing Then Exit Sub

    strValidarCampos = flValidarCamposExcecao
    If strValidarCampos <> vbNullString Then
        Call fgCursor(False)
        frmMural.Caption = Me.Caption
        frmMural.Display = strValidarCampos
        frmMural.Show vbModal
        Exit Sub
    End If
    
    'O 'Key' está separado por ";" pois existe sistemas que começam com a letra "K", o que
    'inviabilizava a separação dos código por meio da função Split()
    
    strItemKey = ";" & strCodTipoOperacao & _
                 ";" & strCodFuncaoSistema & _
                 ";" & fgObterCodigoCombo(cboEmpresa.Text) & _
                 ";" & fgObterCodigoCombo(cboSistema.Text) & _
                 ";" & fgObterCodigoCombo(cboLocalLiqu.Text) & _
                 ";" & fgObterCodigoCombo(cboGrupoUsuario.Text) & _
                 ";" & fgObterDescricaoCombo(cboLocalLiqu.Text)
                                                 
    On Error GoTo ErrorHandler
    
    Set objListItem = lstExcecao.ListItems.Add(, strItemKey)
    objListItem.SubItems(1) = cboSistema.Text
    objListItem.SubItems(2) = cboLocalLiqu.Text
    objListItem.SubItems(3) = cboGrupoUsuario.Text
    
    objListItem.Tag = enumTipoOperacao.Incluir
    
    Exit Sub

ErrorHandler:
    
    If Err.Number = 35602 Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Item já incluído anteriormente nesta lista."
        frmMural.Show vbModal
        Exit Sub
    End If

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbAplicar_ButtonClick"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True
    
    Select Case Button.Key
    Case "Limpar"
        Call flLimparDetalhe
    Case gstrSalvar
    
        If flValidarCampos <> "" Then
            frmMural.Caption = Me.Caption
            frmMural.Display = flValidarCampos
            frmMural.Show vbModal
            Exit Sub
        End If
        
        fgLockWindow Me.hwnd
        If flSalvarTodos Then
            Call flCarregarParametrizacao
            Call flObterDetalheParametrizacao(objLstItem)
            Call flObterExcecoesParametrizacao
            Call fgCursor(False)
            fgLockWindow 0
            MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        End If
        fgLockWindow 0
        
    Case gstrSair
        fgCursor False
        Unload Me
    End Select
    
    fgCursor False
    
    Exit Sub

ErrorHandler:
    fgLockWindow 0
    fgCursor False
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
    Call flLimparDetalhe
    
    If lstRegra.ListItems.Count > 0 Then
        Call flCarregarParametrizacao
        Call flObterDetalheParametrizacao(lstRegra.SelectedItem)
        Call flObterExcecoesParametrizacao
    End If
    
End Sub
