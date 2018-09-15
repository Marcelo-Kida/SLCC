VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroConta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Histórico Conta Corrente"
   ClientHeight    =   8460
   ClientLeft      =   2055
   ClientTop       =   1395
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCadastro 
      Height          =   7920
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11505
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
         Height          =   4665
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   11280
         Begin VB.ComboBox cboVenda 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2340
            Width           =   3735
         End
         Begin VB.TextBox txtProduto 
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   5
            Top             =   2940
            Width           =   3735
         End
         Begin VB.ComboBox cboFinalidadeTED 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3540
            Width           =   3735
         End
         Begin VB.ComboBox cboSubTipoAtivo 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1710
            Width           =   3735
         End
         Begin VB.TextBox txtEmpresa 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   420
            Width           =   3735
         End
         Begin VB.TextBox txtSistema 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   4020
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   420
            Width           =   3744
         End
         Begin VB.TextBox txtHistorico 
            Height          =   315
            Left            =   120
            MaxLength       =   4
            TabIndex        =   10
            Top             =   4200
            Width           =   3735
         End
         Begin VB.Frame fraTipoDebitoCredito 
            Caption         =   "Tipo de Lançamento Conta Corrente"
            Height          =   1072
            Left            =   4020
            TabIndex        =   7
            Top             =   1590
            Width           =   3776
            Begin VB.OptionButton optEstornoCredito 
               Caption         =   "Estorno de Crédito"
               Height          =   255
               Left            =   2000
               TabIndex        =   21
               Top             =   672
               Width           =   1712
            End
            Begin VB.OptionButton optEstornoDebito 
               Caption         =   "Estorno de Débito"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   704
               Width           =   1712
            End
            Begin VB.OptionButton optCredito 
               Caption         =   "Crédito"
               Height          =   195
               Left            =   2000
               TabIndex        =   9
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton optDebito 
               Caption         =   "Débito"
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Top             =   360
               Width           =   795
            End
         End
         Begin VB.ComboBox cboTipoOperacao 
            Height          =   336
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1050
            Width           =   7696
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   3735
         End
         Begin VB.TextBox txtTipoOperacao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   1050
            Width           =   7095
         End
         Begin VB.Label lblVenda 
            AutoSize        =   -1  'True
            Caption         =   "Canal de Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   2100
            Width           =   1140
         End
         Begin VB.Label lblProduto 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Produto"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   2700
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Finalidade TED"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   3300
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sub-tipo Ativo"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1470
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Sistema de Conta Corrente"
            Height          =   195
            Left            =   4020
            TabIndex        =   19
            Top             =   180
            Width           =   1890
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Código do Histórico de Conta Corrente"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   3960
            Width           =   2715
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Operação"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   810
            Width           =   4515
         End
         Begin VB.Label Label2 
            Caption         =   "Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   180
            Width           =   675
         End
      End
      Begin MSComctlLib.ListView lvwConta 
         Height          =   2925
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5159
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
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   90
      Top             =   7830
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
            Picture         =   "frmCadastroConta.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroConta.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroConta.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroConta.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroConta.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroConta.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroConta.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   7440
      TabIndex        =   15
      Top             =   8040
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
End
Attribute VB_Name = "frmCadastroConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pelo cadastramento de conta corrente que o sistema SLCC deverá utilizar,
'' através da camada de controle de caso de uso MIU.
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private strOperacao                         As String

Private strKeyItemSelected                  As String

Public lngBackOffice                        As Long

Private Const COL_EMPRESA                   As Integer = 0
Private Const COL_TIPOOPERACAO              As Integer = 1
Private Const COL_TIPO_MOVTO                As Integer = 2
Private Const COL_HISTORICO                 As Integer = 3
Private Const COL_SUB_TIPO_ATIVO            As Integer = 4
Private Const COL_FINALIDADE_TED            As Integer = 5
Private Const COL_PRODUTO                   As Integer = 6
Private Const COL_CANAL_VENDA               As Integer = 7

Private Const strFuncionalidade             As String = "frmCadastroConta"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private xmlSubTipoAtivo                     As MSXML2.DOMDocument40
Private xmlFinalidadeTED                    As MSXML2.DOMDocument40
Private xmlProduto                          As MSXML2.DOMDocument40

Private Sub flCarregarComboFinalidadeTED(ByVal pstrTipoOperacao As String)
Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrTipoOperacao = vbNullString Then Exit Sub
    
    cboFinalidadeTED.Clear
    
    For Each objNode In xmlMapaNavegacao.selectNodes("//Repeat_TipoOperacao/Grupo_TipoOperacao[TP_OPER='" & pstrTipoOperacao & "']")
        
        If (IIf(objNode.selectSingleNode("TP_MESG_RECB_INTE").Text = "", 0, objNode.selectSingleNode("TP_MESG_RECB_INTE").Text)) = enumTipoMensagemLQS.EnvioTEDClientes Then
            Call fgCarregarCombos(Me.cboFinalidadeTED, xmlFinalidadeTED, "DominioAtributo", "CO_DOMI", "DE_DOMI")
        End If
    Next
    
    cboFinalidadeTED.AddItem "< Padrão >", 0
    cboFinalidadeTED.ListIndex = 0
    cboFinalidadeTED.Enabled = cboFinalidadeTED.ListCount > 1
    
    Exit Sub



ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flCarregarComboFinalidadeTED"
End Sub

Private Sub flCarregarComboSubTipoAtivo(ByVal pstrTipoOperacao As String)

Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrTipoOperacao = vbNullString Then Exit Sub
    
    cboSubTipoAtivo.Clear
    
    For Each objNode In xmlMapaNavegacao.selectNodes("//Repeat_TipoOperacao/Grupo_TipoOperacao[TP_OPER='" & pstrTipoOperacao & "']")
        If objNode.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
            Call fgCarregarCombos(Me.cboSubTipoAtivo, xmlSubTipoAtivo, "DominioAtributo", "CO_DOMI", "DE_DOMI")
        End If
    Next
    
    cboSubTipoAtivo.AddItem "< Padrão >", 0
    cboSubTipoAtivo.ListIndex = 0
    cboSubTipoAtivo.Enabled = cboSubTipoAtivo.ListCount > 1
    
    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flCarregarComboSubTipoAtivo"

End Sub

Private Sub flCarregarComboCanalVenda()

On Error GoTo ErrorHandler

    cboVenda.Clear
    
    cboVenda.AddItem "< Padrão >", 0
    cboVenda.AddItem "1 - SGC", 1
    cboVenda.AddItem "2 - SGM", 2
    cboVenda.ListIndex = 0
    cboVenda.Enabled = cboVenda.ListCount > 1
    
    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flCarregarComboCanalVenda"

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

    lvwConta.ColumnHeaders.Add 1, , "Empresa", 1720, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 2, , "Tipo Operação", 3000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 3, , "Tipo Movto.", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 4, , "Histórico", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 5, , "Sub-tipo Ativo", 1200, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 6, , "Finalidade TED", 3000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 7, , "Sub-Produto", 3000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 8, , "Canal de Venda", 3000, lvwColumnLeft

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
            strKeyItemSelected = "|" & .selectSingleNode("SG_SIST").Text & _
                                 "|" & .selectSingleNode("CO_EMPR").Text & _
                                 "|" & .selectSingleNode("TP_BKOF").Text & _
                                 "|" & .selectSingleNode("TP_OPER").Text & _
                                 "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                 "|" & IIf(.selectSingleNode("CO_SUB_TIPO_ATIV").Text = vbNullString, "0", .selectSingleNode("CO_SUB_TIPO_ATIV").Text) & _
                                 "|" & IIf(.selectSingleNode("CD_FIND_TED").Text = vbNullString, "0", .selectSingleNode("CD_FIND_TED").Text) & _
                                 "|" & IIf(.selectSingleNode("CD_SUB_PROD").Text = vbNullString, "0", .selectSingleNode("CD_SUB_PROD").Text) & _
                                 "|" & IIf(.selectSingleNode("TP_CNAL_VEND").Text = vbNullString, "0", .selectSingleNode("TP_CNAL_VEND").Text)
                                 
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

    If strOperacao <> gstrOperExcluir Then
       With xmlLer.documentElement
            strKeyItemSelected = "|" & .selectSingleNode("SG_SIST").Text & _
                                 "|" & .selectSingleNode("CO_EMPR").Text & _
                                 "|" & .selectSingleNode("TP_BKOF").Text & _
                                 "|" & .selectSingleNode("TP_OPER").Text & _
                                 "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                 "|" & IIf(.selectSingleNode("CO_SUB_TIPO_ATIV").Text = vbNullString, "0", .selectSingleNode("CO_SUB_TIPO_ATIV").Text) & _
                                 "|" & IIf(.selectSingleNode("CD_FIND_TED").Text = vbNullString, "0", .selectSingleNode("CD_FIND_TED").Text) & _
                                 "|" & IIf(.selectSingleNode("CD_SUB_PROD").Text = vbNullString, "0", .selectSingleNode("CD_SUB_PROD").Text) & _
                                 "|" & IIf(.selectSingleNode("TP_CNAL_VEND").Text = vbNullString, "0", .selectSingleNode("TP_CNAL_VEND").Text)
       End With
    End If

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
        
        If Not (optDebito.value Or optCredito.value Or optEstornoCredito.value Or optEstornoDebito.value) Then
            flValidarCampos = "Selecione o Tipo de Movimento."
            Exit Function
        End If
        
        If cboTipoOperacao.ListIndex = -1 Then
            flValidarCampos = "Selecione o Tipo de Operação."
            cboTipoOperacao.SetFocus
            Exit Function
        End If
    
    End If
    
    With txtHistorico
        If Val(.Text) = 0 Then
            flValidarCampos = "Preencha um código de histórico."
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

    cboEmpresa.ListIndex = -1
    cboEmpresa.Visible = True
    txtEmpresa.Visible = False
    
    txtHistorico.Text = vbNullString
    txtProduto.Text = vbNullString
    txtProduto.Enabled = False
    
    cboTipoOperacao.ListIndex = -1
    cboTipoOperacao.Visible = True
    txtTipoOperacao.Visible = False
    
    fraTipoDebitoCredito.Enabled = True
    optCredito.value = False
    optDebito.value = False
    optEstornoCredito.value = False
    optEstornoDebito.value = False
'    optEstornoCredito.Enabled = False
'    optEstornoDebito.Enabled = False

    cboSubTipoAtivo.Clear
    cboFinalidadeTED.Clear
    cboVenda.Clear
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "frmGrupoUsuario", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub
Private Sub flCarregarProduto(ByVal pstrTipoOper As String)
Dim blnProdutoEnable            As Boolean
    
    If pstrTipoOper = vbNullString Then Exit Sub
    blnProdutoEnable = fgIN(CLng(pstrTipoOper), enumTipoOperacaoLQS.EventosJurosSWAP, _
                                                enumTipoOperacaoLQS.RegistroContratoSWAP, _
                                                enumTipoOperacaoLQS.RegDadosComplemContratoSWAP, _
                                                enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, _
                                                enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo, _
                                                enumTipoOperacaoLQS.ExercicioOpcaoContratoSWAP, _
                                                enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais, _
                                                enumTipoOperacaoLQS.LancamentoCCCashFlow, _
                                                enumTipoOperacaoLQS.LancamentoCCSwapJuros, _
                                                enumTipoOperacaoLQS.LancamentoCCCashFlowStrikeFixo _
                                                )
         
    txtProduto.Enabled = (blnProdutoEnable And strOperacao = gstrOperIncluir)
    
    
End Sub

Private Sub flCarregarCanalVenda(ByVal pstrTipoOper As String)
Dim blnVendaEnable            As Boolean
    
    If pstrTipoOper = vbNullString Then Exit Sub
    blnVendaEnable = fgIN(CLng(pstrTipoOper), enumTipoOperacaoLQS.EventosJurosSWAP, _
                                              enumTipoOperacaoLQS.RegistroContratoSWAP, _
                                              enumTipoOperacaoLQS.RegDadosComplemContratoSWAP, _
                                              enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, _
                                              enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo, _
                                              enumTipoOperacaoLQS.ExercicioOpcaoContratoSWAP, _
                                              enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira, _
                                              enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP_CETIP21)
         
    cboVenda.Enabled = (blnVendaEnable And strOperacao = gstrOperIncluir)
    
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
        .selectSingleNode("//SG_SIST").Text = Split(strChaveRegistro, "|")(1)
        .selectSingleNode("//CO_EMPR").Text = Split(strChaveRegistro, "|")(2)
        .selectSingleNode("//TP_BKOF").Text = Split(strChaveRegistro, "|")(3)
        .selectSingleNode("//TP_OPER").Text = Split(strChaveRegistro, "|")(4)
        .selectSingleNode("//IN_LANC_DEBT_CRED").Text = Split(strChaveRegistro, "|")(5)
        .selectSingleNode("//CO_SUB_TIPO_ATIV").Text = Split(strChaveRegistro, "|")(6)
        .selectSingleNode("//CD_FIND_TED").Text = Split(strChaveRegistro, "|")(7)
        .selectSingleNode("//CD_SUB_PROD").Text = Split(strChaveRegistro, "|")(8)
        .selectSingleNode("//TP_CNAL_VEND").Text = Split(strChaveRegistro, "|")(9)
    End With

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call xmlLer.loadXML(objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

    With xmlLer.documentElement
        
        txtEmpresa.Visible = True
        cboEmpresa.Visible = False
        txtEmpresa.Text = .selectSingleNode("CO_EMPR").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_Empresa[CO_EMPR=" & .selectSingleNode("CO_EMPR").Text & "]/NO_REDU_EMPR").Text
        
        txtSistema.Text = .selectSingleNode("SG_SIST").Text
        
        txtTipoOperacao.Visible = True
        cboTipoOperacao.Visible = False
        txtTipoOperacao.Text = .selectSingleNode("TP_OPER").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_TipoOperacao[TP_OPER='" & .selectSingleNode("TP_OPER").Text & "']/NO_TIPO_OPER").Text
        
        Call flCarregarComboSubTipoAtivo(.selectSingleNode("TP_OPER").Text)
        
        For intIndCombo = 0 To cboSubTipoAtivo.ListCount - 1
            If .selectSingleNode("CO_SUB_TIPO_ATIV").Text = Left$(cboSubTipoAtivo.List(intIndCombo), Len(.selectSingleNode("CO_SUB_TIPO_ATIV").Text)) Then
                cboSubTipoAtivo.ListIndex = intIndCombo
                Exit For
            End If
        Next
        
        cboSubTipoAtivo.Enabled = False
        
        Call flCarregarComboFinalidadeTED(.selectSingleNode("TP_OPER").Text)
        
        For intIndCombo = 0 To cboFinalidadeTED.ListCount - 1
            If .selectSingleNode("CD_FIND_TED").Text = Left$(cboFinalidadeTED.List(intIndCombo), Len(.selectSingleNode("CD_FIND_TED").Text)) Then
                cboFinalidadeTED.ListIndex = intIndCombo
                Exit For
            End If
        Next
        
        cboFinalidadeTED.Enabled = False
                
        'Call flCarregarComboProduto(.selectSingleNode("TP_OPER").Text)
        
        'For intIndCombo = 0 To cboProduto.ListCount - 1
        '    If .selectSingleNode("CD_PROD").Text = Left$(cboProduto.List(intIndCombo), Len(.selectSingleNode("CD_PROD").Text)) Then
        '        cboProduto.ListIndex = intIndCombo
        '        Exit For
        '    End If
        'Next
        
        Call flCarregarProduto(.selectSingleNode("TP_OPER").Text)
        If Not .selectSingleNode("CD_SUB_PROD") Is Nothing Then
            If .selectSingleNode("CD_SUB_PROD").Text = "0" Or .selectSingleNode("CD_SUB_PROD").Text = "" Then
                txtProduto.Text = "Padrão"
            Else
                txtProduto.Text = .selectSingleNode("CD_SUB_PROD").Text
            End If
        Else
            txtProduto.Text = ""
        End If
        
        Call flCarregarComboCanalVenda
        Call flCarregarCanalVenda(.selectSingleNode("TP_OPER").Text)
        If Not .selectSingleNode("TP_CNAL_VEND") Is Nothing Then
            If Val(.selectSingleNode("TP_CNAL_VEND").Text) = 0 Then
                cboVenda.ListIndex = 0
            Else
                cboVenda.ListIndex = Val(.selectSingleNode("TP_CNAL_VEND").Text)
            End If
        Else
            cboVenda.ListIndex = -1
        End If
                 
        txtHistorico.Text = .selectSingleNode("CO_HIST_CC").Text
        
        fraTipoDebitoCredito.Enabled = False
        Select Case Val(.selectSingleNode("IN_LANC_DEBT_CRED").Text)
            Case enumTipoDebitoCredito.Credito
                optCredito.value = True
            Case enumTipoDebitoCredito.Debito
                optDebito.value = True
            Case enumTipoDebitoCreditoEstorno.EstornoCredito
                optEstornoCredito.value = True
            Case enumTipoDebitoCreditoEstorno.EstornoDebito
                optEstornoDebito.value = True
        End Select

        strUltimaAtualizacao = .selectSingleNode("DH_ULTI_ATLZ").Text
        
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

On Error GoTo ErrorHandler

    With xmlLer.documentElement

         .selectSingleNode("@Operacao").Text = strOperacao
         
         If strOperacao <> gstrOperExcluir Then
         
            If strOperacao = gstrOperIncluir Then
               .selectSingleNode("SG_SIST").Text = txtSistema.Text
               .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa.Text)
               .selectSingleNode("TP_BKOF").Text = lngBackOffice
               .selectSingleNode("TP_OPER").Text = fgObterCodigoCombo(cboTipoOperacao.Text)
               
               .selectSingleNode("CO_SUB_TIPO_ATIV").Text = "0"
               .selectSingleNode("CD_FIND_TED").Text = 0
               .selectSingleNode("CD_SUB_PROD").Text = "0"
               .selectSingleNode("TP_CNAL_VEND").Text = 0
               If cboSubTipoAtivo.Enabled Then
                    
                    .selectSingleNode("CO_SUB_TIPO_ATIV").Text = fgObterCodigoCombo(cboSubTipoAtivo.Text)
               
               ElseIf cboFinalidadeTED.Enabled Then
               
                    .selectSingleNode("CD_FIND_TED").Text = fgObterCodigoCombo(cboFinalidadeTED.Text)
                    
               End If
               
               If Trim(txtProduto.Text) <> "" Then
                    .selectSingleNode("CD_SUB_PROD").Text = txtProduto.Text
               End If
               
               If cboVenda.ListIndex <> 0 Then
                    .selectSingleNode("TP_CNAL_VEND").Text = cboVenda.ListIndex
               End If
               
               Select Case True
                    Case optCredito.value
                        lngTipoLancamentoDebtCred = enumTipoDebitoCredito.Credito
                    Case optDebito.value
                        lngTipoLancamentoDebtCred = enumTipoDebitoCredito.Debito
                    Case optEstornoCredito.value
                        lngTipoLancamentoDebtCred = enumTipoDebitoCreditoEstorno.EstornoCredito
                    Case optEstornoDebito.value
                        lngTipoLancamentoDebtCred = enumTipoDebitoCreditoEstorno.EstornoDebito
               End Select
               .selectSingleNode("IN_LANC_DEBT_CRED").Text = lngTipoLancamentoDebtCred
            End If
            
            .selectSingleNode("CO_HIST_CC").Text = Val(txtHistorico.Text)
         
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

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")

    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If

    lngBackOffice = CLng("0" & xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_TipoBackOffice/TP_BKOF").Text)
    Call fgCarregarCombos(cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    Call fgCarregarCombos(cboTipoOperacao, xmlMapaNavegacao, "TipoOperacao", "TP_OPER", "NO_TIPO_OPER")

    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_ParametroHistoricoCC").xml
    End If

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    Set xmlSubTipoAtivo = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlFinalidadeTED = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlSubTipoAtivo.loadXML(objMensagem.ObterDominioSPB("SubTpAtv", vntCodErro, vntMensagemErro))
    Call xmlFinalidadeTED.loadXML(objMensagem.ObterDominioSPB("FinlddCli", vntCodErro, vntMensagemErro))
    
    
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
Dim xmlDomNode              As MSXML2.IXMLDOMNode
Dim xmlDomNodeAux           As MSXML2.IXMLDOMNode
Dim objListItem             As MSComctlLib.ListItem

Dim strNomeTipoLiquidacao   As String

Dim strTempChave            As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
Dim blnCarregaCombo         As Boolean

On Error GoTo ErrorHandler

    fgCursor True

    lvwConta.ListItems.Clear
    lvwConta.HideSelection = False

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")

    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParametroHistoricoCC/@Operacao").Text = "LerTodos"
    Call xmlLerTodos.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParametroHistoricoCC").xml, vntCodErro, vntMensagemErro))

    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If


    Set objMIU = Nothing
    
    blnCarregaCombo = True
    
    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_ParametroHistoricoCC/*")
        
        With xmlDomNode
            
                        
            strTempChave = "|" & .selectSingleNode("SG_SIST").Text & _
                           "|" & .selectSingleNode("CO_EMPR").Text & _
                           "|" & .selectSingleNode("TP_BKOF").Text & _
                           "|" & .selectSingleNode("TP_OPER").Text & _
                           "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                           "|" & .selectSingleNode("CO_SUB_TIPO_ATIV").Text & _
                           "|" & .selectSingleNode("CD_FIND_TED").Text & _
                           "|" & .selectSingleNode("CD_SUB_PROD").Text & _
                           "|" & .selectSingleNode("TP_CNAL_VEND").Text
                            
            Set objListItem = lvwConta.ListItems.Add(, strTempChave)
            objListItem.Text = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & .selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
            objListItem.SubItems(COL_TIPO_MOVTO) = fgTraduzDebitoCredito(.selectSingleNode("IN_LANC_DEBT_CRED").Text)
            objListItem.SubItems(COL_TIPOOPERACAO) = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_TipoOperacao/Grupo_TipoOperacao[TP_OPER='" & .selectSingleNode("TP_OPER").Text & "']/NO_TIPO_OPER").Text
            objListItem.SubItems(COL_HISTORICO) = .selectSingleNode("CO_HIST_CC").Text
            
           
            Select Case Val(.selectSingleNode("TP_OPER").Text)
                
                Case enumTipoOperacaoLQS.EnvioPAG0105Clientes, _
                     enumTipoOperacaoLQS.EnvioPAG0106Clientes, _
                     enumTipoOperacaoLQS.EnvioPAG0108Clientes, _
                     enumTipoOperacaoLQS.EnvioPAG0109Clientes, _
                     enumTipoOperacaoLQS.EnvioPAG0121Clientes, _
                     enumTipoOperacaoLQS.EnvioPAG0134Clientes, _
                     enumTipoOperacaoLQS.EnvioTEDSTR0006Clientes, _
                     enumTipoOperacaoLQS.EnvioTEDSTR0007Clientes, _
                     enumTipoOperacaoLQS.EnvioTEDSTR0008Clientes, _
                     enumTipoOperacaoLQS.EnvioTEDSTR0009Clientes, _
                     enumTipoOperacaoLQS.EnvioTEDSTR0025Clientes, _
                     enumTipoOperacaoLQS.EnvioTEDSTR0034Clientes, _
                     enumTipoOperacaoLQS.EmissaoTEDPAG0106FdosFIDC, _
                     enumTipoOperacaoLQS.EmissaoTEDPAG0108FdosFIDC, _
                     enumTipoOperacaoLQS.EmissaoTEDSTR0007FdosFIDC, _
                     enumTipoOperacaoLQS.EmissaoTEDSTR0008FdosFIDC
                    
                    For Each xmlDomNodeAux In xmlFinalidadeTED.selectNodes("//Repeat_DominioAtributo/*")
                        If .selectSingleNode("CD_FIND_TED").Text = "0" Then objListItem.SubItems(COL_FINALIDADE_TED) = "Padrao": Exit For
                        If xmlDomNodeAux.selectSingleNode("CO_DOMI").Text = .selectSingleNode("CD_FIND_TED").Text Then
                            objListItem.SubItems(COL_FINALIDADE_TED) = IIf(xmlDomNodeAux.selectSingleNode("DE_DOMI").Text = "", "Padrao", xmlDomNodeAux.selectSingleNode("CO_DOMI").Text & "-" & xmlDomNodeAux.selectSingleNode("DE_DOMI").Text)
                            Exit For
                        End If
                    Next
                                                   
                    objListItem.SubItems(COL_SUB_TIPO_ATIVO) = "Padrão"
                    objListItem.SubItems(COL_PRODUTO) = "Padrão"
                
                Case enumTipoOperacaoLQS.EventosJurosSWAP, _
                     enumTipoOperacaoLQS.RegistroContratoSWAP, _
                     enumTipoOperacaoLQS.RegDadosComplemContratoSWAP, _
                     enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, _
                     enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo, _
                     enumTipoOperacaoLQS.ExercicioOpcaoContratoSWAP

                    objListItem.SubItems(COL_FINALIDADE_TED) = "Padrão"
                    
                    If Val(.selectSingleNode("TP_CNAL_VEND").Text) = 0 Then
                        objListItem.SubItems(COL_CANAL_VENDA) = "Padrão"
                    ElseIf Val(.selectSingleNode("TP_CNAL_VEND").Text) = 1 Then
                        objListItem.SubItems(COL_CANAL_VENDA) = "SGC"
                    ElseIf Val(.selectSingleNode("TP_CNAL_VEND").Text) = 2 Then
                        objListItem.SubItems(COL_CANAL_VENDA) = "SGM"
                    End If
                    
                Case Else
                    objListItem.SubItems(COL_FINALIDADE_TED) = "Padrão"
                    objListItem.SubItems(COL_PRODUTO) = "Padrão"
                    objListItem.SubItems(COL_CANAL_VENDA) = "Padrão"
                    
            End Select
            
            If Not .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                objListItem.SubItems(COL_SUB_TIPO_ATIVO) = IIf(.selectSingleNode("CO_SUB_TIPO_ATIV").Text = "0", "Padrão", .selectSingleNode("CO_SUB_TIPO_ATIV").Text)
            End If
            
            If Not .selectSingleNode("CD_SUB_PROD") Is Nothing Then
                If .selectSingleNode("CD_SUB_PROD").Text = "0" Or .selectSingleNode("CD_SUB_PROD").Text = "" Then
                    objListItem.SubItems(COL_PRODUTO) = "Padrão"
                Else
                    objListItem.SubItems(COL_PRODUTO) = .selectSingleNode("CD_SUB_PROD").Text
                End If
            Else
                objListItem.SubItems(COL_PRODUTO) = "Padrão"
            End If
            
        End With
    Next

    Set xmlLerTodos = Nothing
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

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler

    txtSistema.Text = "BG"
    
    If cboEmpresa.ListIndex = -1 Then
        txtSistema.Text = vbNullString
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboEmpresa_Click"

End Sub

Private Sub cboTipoOperacao_Click()

Dim strTipoOper                             As String

    On Error GoTo ErrorHandler
    
    strTipoOper = fgObterCodigoCombo(cboTipoOperacao.Text)
    
    Call flCarregarComboSubTipoAtivo(strTipoOper)
    Call flCarregarComboFinalidadeTED(strTipoOper)
    Call flCarregarProduto(strTipoOper)
    
    Call flCarregarComboCanalVenda
    Call flCarregarCanalVenda(strTipoOper)

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboTipoOperacao_Click"

End Sub

Private Sub cboVenda_Click()

    On Error GoTo ErrorHandler

'    optEstornoCredito.Enabled = IIf(Val(fgObterCodigoCombo(cboVenda.Text)) = enumCanalDeVenda.SGC, True, False)
'    optEstornoDebito.Enabled = optEstornoCredito.Enabled

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao

    Call flLimpaCampos

    Call fgCursor(True)
    Call flFormataListView
    Call flInicializar
    Call flCarregaListView

    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    Set xmlLer = Nothing
    Set xmlSubTipoAtivo = Nothing
    Set xmlFinalidadeTED = Nothing
    Set xmlProduto = Nothing

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

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtHistorico_KeyPress"

End Sub
