VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadastroContaCOTR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Conta Contábil Corretoras"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipoConta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   60
      TabIndex        =   8
      Top             =   1710
      Width           =   7275
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   4935
      End
      Begin VB.ComboBox cboLocaLiquidacao 
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtCCCorretora 
         Height          =   315
         Left            =   2070
         MaxLength       =   13
         TabIndex        =   3
         Top             =   1410
         Width           =   1455
      End
      Begin VB.TextBox txtAgenCorretora 
         Height          =   315
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtUniOrg 
         Height          =   315
         Left            =   2070
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtLocaLiqu 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtEmpresa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2070
         TabIndex        =   6
         Top             =   630
         Width           =   4935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   1380
         TabIndex        =   13
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Local de Liquidação"
         Height          =   195
         Left            =   570
         TabIndex        =   12
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conta Contábil Corretora"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1470
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Agência"
         Height          =   195
         Left            =   870
         TabIndex        =   10
         Top             =   1095
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código UNIORG"
         Height          =   195
         Left            =   780
         TabIndex        =   9
         Top             =   1860
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroContaCOTR.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3450
      TabIndex        =   7
      Top             =   4020
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwConta 
      Height          =   1605
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCadastroContaCOTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pelo cadastramento de conta corrente de Corretoras que o sistema SLCC deverá utilizar,
'' através da camada de controle de caso de uso MIU.
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private strOperacao                         As String

Private strKeyItemSelected                  As String

Public lngBackOffice                        As Long

Private Const COL_LOCAL_LIQUIDACAO          As Integer = 0
Private Const COL_EMPRESA                   As Integer = 1
Private Const COL_AGENCIA                   As Integer = 2
Private Const COL_CONTA_CORRENTE            As Integer = 3
Private Const COL_UNI_ORG                   As Integer = 4

Private Const strFuncionalidade             As String = "frmConsultaMensagem"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private xmlSubTipoAtivo                     As MSXML2.DOMDocument40

Private Sub flCarregarComboLocalLiquidacao(ByVal pstrTipoOperacao As String)

Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrTipoOperacao = vbNullString Then Exit Sub

    cboLocaLiquidacao.Clear

    For Each objNode In xmlMapaNavegacao.selectNodes("//Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & pstrTipoOperacao & "']")
        If objNode.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
            'Call fgCarregarCombos(Me.cboSubTipoAtivo, xmlSubTipoAtivo, "DominioAtributo", "CO_DOMI", "DE_DOMI")
        End If
    Next

    cboLocaLiquidacao.AddItem "< Padrão >", 0
    cboLocaLiquidacao.ListIndex = 0
    cboLocaLiquidacao.Enabled = cboLocaLiquidacao.ListCount > 1

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flCarregarComboSubTipoAtivo"

End Sub

Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lvwConta.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub

    blnEncontrou = False
    For Each objListItem In lvwConta.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lvwConta_ItemClick objListItem
           lvwConta.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing

    If Not blnEncontrou Then
       flLimpaCampos
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

Private Sub flFormataListView()

    lvwConta.ColumnHeaders.Add 1, , "Local Liquidação", 2000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 2, , "Empresa", 1720, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 3, , "Código Agência", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 4, , "Conta Contábil", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 5, , "UNIORG", 1200, lvwColumnLeft

End Sub

'' Salva as alterações efetuadas através da camada controladora de casos de uso
'' MIU, método A8MIU.clsMIU.Executar

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim strRetorno             As String
Dim strPropriedades        As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()

    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If

    Call fgCursor(True)

    Call flInterfaceToXml

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    If strOperacao = "Incluir" Then
        With xmlLer.documentElement
            strKeyItemSelected = "|" & .selectSingleNode("//CO_LOCA_LIQU").Text & _
                                 "|" & .selectSingleNode("//CO_EMPR").Text & _
                                 "|" & .selectSingleNode("//CO_AGEN_COTR").Text & _
                                 "|" & .selectSingleNode("//NU_CC_COTR").Text & _
                                 "|" & .selectSingleNode("//CO_UNI_ORG").Text
                                 
        End With
    End If

    Call objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        strOperacao = gstrOperAlterar
        flXmlToInterface
    Else
        flLimpaCampos
    End If
    Set objMIU = Nothing

    Call flCarregaListView
    Call fgCursor(False)

    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
    Exit Sub

ErrorHandler:

    Call fgCursor(False)

    Set objMIU = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmGrupoUsuario", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Validar os campos obrigatórios para execução da funcionalidade especificada.

Private Function flValidarCampos() As String

On Error GoTo ErrorHandler
    
    If strOperacao = gstrOperIncluir Then
        If cboEmpresa.ListIndex = -1 Then
            flValidarCampos = "Selecione a Empresa."
            cboEmpresa.SetFocus
            Exit Function
        End If
        
        If cboLocaLiquidacao.ListIndex = -1 Then
            flValidarCampos = "Selecione o Local de Liquidação."
            cboLocaLiquidacao.SetFocus
            Exit Function
        End If
    End If
    
    With txtCCCorretora
        If Val(.Text) = 0 Then
            flValidarCampos = "Preencha com uma Conta Corrente."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With
    
    With txtAgenCorretora
        If Val(.Text) = 0 Then
            flValidarCampos = "Preencha com uma Agencia."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With

    With txtUniOrg
        If Val(.Text) = 0 Then
            flValidarCampos = "Preencha com uma Uni Org."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With

    flValidarCampos = ""

    Exit Function

ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Limpar todos os campos para uma nova inclusão.

Private Sub flLimpaCampos()

On Error GoTo ErrorHandler

    strOperacao = "Incluir"

    cboEmpresa.Visible = True
    txtEmpresa.Visible = False
    txtEmpresa.Text = ""
    
    cboLocaLiquidacao.ListIndex = -1
    cboLocaLiquidacao.Visible = True
    cboLocaLiquidacao.Enabled = True
    txtLocaLiqu.Visible = False
    txtLocaLiqu.Text = ""
    
    txtAgenCorretora.Visible = True
    txtCCCorretora.Visible = True
    txtUniOrg.Visible = True
    txtAgenCorretora.Enabled = True
    txtCCCorretora.Enabled = True
    txtUniOrg.Enabled = True
    txtAgenCorretora.Text = ""
    txtCCCorretora.Text = ""
    txtUniOrg.Text = ""
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False

    Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, "frmGrupoUsuario", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Carregar a Interface com os dados obtidos através da leitura executando o método A8MIU.clsMIU.Executar

Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim strChaveRegistro       As String
Dim intIndCombo            As Integer
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    If lvwConta.SelectedItem Is Nothing Then
        flLimpaCampos
        Exit Sub
    End If

    strChaveRegistro = lvwConta.SelectedItem.Key

    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//CO_LOCA_LIQU").Text = Split(strChaveRegistro, "|")(1) & " - " & Split(strChaveRegistro, "|")(2)
        .selectSingleNode("//CO_EMPR").Text = Split(strChaveRegistro, "|")(3) & " - " & Split(strChaveRegistro, "|")(4)
        .selectSingleNode("//CO_AGEN_COTR").Text = Split(strChaveRegistro, "|")(5)
        .selectSingleNode("//NU_CC_COTR").Text = Split(strChaveRegistro, "|")(6)
        .selectSingleNode("//CO_UNI_ORG").Text = Split(strChaveRegistro, "|")(7)
    End With

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
'    Call xmlLer.loadXML(objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

    With xmlLer.documentElement
        
        txtLocaLiqu.Visible = True
        cboLocaLiquidacao.Visible = False
        txtLocaLiqu.Text = .selectSingleNode("//CO_LOCA_LIQU").Text '& " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & .selectSingleNode("CO_LOCA_LIQU").Text & "']/DE_LOCA_LIQU").Text
        
        txtEmpresa.Visible = True
        cboEmpresa.Visible = False
        txtEmpresa.Text = .selectSingleNode("//CO_EMPR").Text '& " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_Empresa[CO_EMPR=" & .selectSingleNode("CO_EMPR").Text & "]/NO_REDU_EMPR").Text
        
        cboLocaLiquidacao.Enabled = False
        
        txtAgenCorretora.Text = .selectSingleNode("//CO_AGEN_COTR").Text
        txtAgenCorretora.Enabled = False
        txtCCCorretora.Text = .selectSingleNode("//NU_CC_COTR").Text
        txtCCCorretora.Enabled = False
        txtUniOrg.Text = .selectSingleNode("//CO_UNI_ORG").Text
        
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    End With

    Exit Sub

ErrorHandler:

    Set objMIU = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmGrupoUsuario", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Carregar o xml de envio para a camada MIU com os dados obtidos da Interface.

Private Function flInterfaceToXml() As String

Dim lngTipoLancamentoDebtCred               As Long
Dim strAgenCorretora                        As String
Dim strCCCorretora                          As String
Dim strUniOrg                               As String

On Error GoTo ErrorHandler

    With xmlLer.documentElement

         .selectSingleNode("//@Operacao").Text = strOperacao
         
         If strOperacao <> gstrOperExcluir Then
         
            If strOperacao = gstrOperIncluir Then
               .selectSingleNode("//CO_LOCA_LIQU").Text = fgObterCodigoCombo(cboLocaLiquidacao.Text)
               .selectSingleNode("//CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa.Text)
               .selectSingleNode("//CO_AGEN_COTR").Text = txtAgenCorretora.Text
               .selectSingleNode("//NU_CC_COTR").Text = txtCCCorretora.Text
               .selectSingleNode("//CO_UNI_ORG").Text = txtUniOrg.Text
            End If
         End If

         If strOperacao = gstrOperExcluir Or strOperacao = gstrOperAlterar Then
            .selectSingleNode("//CO_LOCA_LIQU").Text = fgObterCodigoCombo(txtLocaLiqu.Text)
            .selectSingleNode("//CO_EMPR").Text = fgObterCodigoCombo(txtEmpresa.Text)
            .selectSingleNode("//CO_AGEN_COTR").Text = txtAgenCorretora.Text
            .selectSingleNode("//NU_CC_COTR").Text = txtCCCorretora.Text
            .selectSingleNode("//CO_UNI_ORG_OLD").Text = .selectSingleNode("//CO_UNI_ORG").Text
            .selectSingleNode("//CO_UNI_ORG").Text = txtUniOrg.Text
         End If

    End With

    Exit Function

ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0

End Function

'' Carrega as propriedades necessárias a interface frmCadastroConta, através da
'' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
    Dim objMensagem        As A8MIU.clsMensagem
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant
Dim strCarregar            As String

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")

    strCarregar = "A8LQS.clsContaCorrenteCOTR"

    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If

    Call fgCarregarCombos(cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    Call fgCarregarCombos(cboLocaLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU")

    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       Call fgAppendNode(xmlLer, "", "Repeat_Filtros", "")
       Call fgAppendNode(xmlLer, "Repeat_Filtros", "Grupo_Propriedades", "")
       Call fgAppendAttribute(xmlLer, "Grupo_Propriedades", "Operacao", "LerTodos")
       Call fgAppendAttribute(xmlLer, "Grupo_Propriedades", "Objeto", strCarregar)
       Call xmlLer.loadXML(objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
    End If

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    Set objMensagem = Nothing
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'' Carrega as contas corrente existentes e preenche o listview com os mesmos,
'' através da classe controladora de caso de uso MIU, método  A8MIU.clsMIU.Executar

Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim xmlLerTodos             As MSXML2.DOMDocument40
Dim xmlCarregar             As MSXML2.DOMDocument40
Dim xmlDomNode              As MSXML2.IXMLDOMNode
Dim objListItem             As MSComctlLib.ListItem
Dim strCarregar             As String

Dim strTempChave            As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    fgCursor True

    strCarregar = "A8LQS.clsContaCorrenteCOTR"

    lvwConta.ListItems.Clear
    lvwConta.HideSelection = False

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlCarregar = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlCarregar, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlCarregar, "Repeat_Filtros", "Grupo_Propriedades", "")
    Call fgAppendAttribute(xmlCarregar, "Grupo_Propriedades", "Operacao", "LerTodos")
    Call fgAppendAttribute(xmlCarregar, "Grupo_Propriedades", "Objeto", strCarregar)
    'xmlCarregar.selectSingleNode("Repeat_Filtros/Grupo_Propriedades/strCarregar").Text = "CO_LOCA_LIQU"
    If xmlCarregar.loadXML(objMIU.Executar(xmlCarregar.xml, vntCodErro, vntMensagemErro)) Then

        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        Set objMIU = Nothing

        For Each xmlDomNode In xmlCarregar.selectNodes("//Repeat_ParametroCCCorretora/*")
            With xmlDomNode
    
                strTempChave = "|" & .selectSingleNode("CO_LOCA_LIQU").Text & _
                               "|" & .selectSingleNode("DE_LOCA_LIQU").Text & _
                               "|" & .selectSingleNode("CO_EMPR").Text & _
                               "|" & .selectSingleNode("NO_REDU_EMPR").Text & _
                               "|" & .selectSingleNode("CO_AGEN_COTR").Text & _
                               "|" & .selectSingleNode("NU_CC_COTR").Text & _
                               "|" & .selectSingleNode("CO_UNI_ORG").Text
    
                Set objListItem = lvwConta.ListItems.Add(, strTempChave)
                objListItem.Text = .selectSingleNode("CO_LOCA_LIQU").Text & " - " & .selectSingleNode("DE_LOCA_LIQU").Text
                objListItem.SubItems(COL_EMPRESA) = .selectSingleNode("CO_EMPR").Text & " - " & .selectSingleNode("NO_REDU_EMPR").Text
                objListItem.SubItems(COL_AGENCIA) = .selectSingleNode("CO_AGEN_COTR").Text
                objListItem.SubItems(COL_CONTA_CORRENTE) = .selectSingleNode("NU_CC_COTR").Text
                objListItem.SubItems(COL_UNI_ORG) = .selectSingleNode("CO_UNI_ORG").Text
                
            End With
        Next

    End If
    
    Set xmlCarregar = Nothing
    fgCursor

    Exit Sub

ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    fgCursor

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"

End Sub

Private Sub Form_Load()

Dim intIndexCombo                           As Integer

    On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    
    Call fgCursor(True)
    DoEvents

    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao

    Call flLimpaCampos

    Call flFormataListView
    Call flInicializar
    Call flCarregaListView
    
    For intIndexCombo = 0 To cboEmpresa.ListCount - 1
        If Val(fgObterCodigoCombo(cboEmpresa.List(intIndexCombo))) = enumCodigoEmpresa.Meridional Then
            cboEmpresa.ListIndex = intIndexCombo
            Exit For
        End If
    Next
    
    cboEmpresa.Enabled = False
    txtEmpresa.Enabled = False

    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    Set xmlLer = Nothing

End Sub

Private Sub lvwConta_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lvwConta.Sorted = True
    lvwConta.SortKey = ColumnHeader.Index - 1

    If lvwConta.SortOrder = lvwAscending Then
        lvwConta.SortOrder = lvwDescending
    Else
        lvwConta.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwConta_ColumnClick"

End Sub

Private Sub lvwConta_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface

    strKeyItemSelected = Item.Key

    Call fgCursor(False)

    Exit Sub

ErrorHandler:

    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - lvwConta_ItemClick", Me.Caption
    flRecarregar

End Sub

Private Sub flRecarregar()

On Error GoTo ErrorHandler

    fgCursor True

    flLimpaCampos
    Call flCarregaListView

    fgCursor

Exit Sub
ErrorHandler:

    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flRecarregar"
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case gstrSalvar
            Call flSalvar
            If strOperacao = gstrOperAlterar Then
               flPosicionaItemListView
            End If
    
            If Me.cboLocaLiquidacao.Visible Then
                cboLocaLiquidacao.SetFocus
            End If
        Case gstrOperExcluir
            If MsgBox("Confirma a exclusão do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
               strOperacao = gstrOperExcluir
               Call flSalvar
            End If
        Case gstrSair
            fgCursor False
            Unload Me
            Exit Sub
    End Select

    fgCursor False

    Exit Sub

ErrorHandler:

    fgCursor False

    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - tlbCadastro_ButtonClick", Me.Caption

    Call flCarregaListView

    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub

Private Sub txtAgenCorretora_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtAgenCorretora_KeyPress"

End Sub

Private Sub txtCCCorretora_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtCCCorretora_KeyPress"

End Sub

Private Sub txtUniOrg_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtUniOrg_KeyPress"

End Sub
