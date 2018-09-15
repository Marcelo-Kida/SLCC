VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParamReprocCC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Parametrização Reprocessamento Conta Corrente"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCadastro 
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6165
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
         Height          =   3750
         Left            =   60
         TabIndex        =   13
         Top             =   2700
         Width           =   6000
         Begin VB.TextBox numIntervalo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "0"
            Top             =   2880
            Width           =   735
         End
         Begin VB.CheckBox chkAtivo 
            Caption         =   "Ativo"
            Height          =   255
            Left            =   4320
            TabIndex        =   8
            Top             =   2880
            Width           =   975
         End
         Begin VB.CheckBox chkHrLimiItgr 
            Caption         =   "Horário Limite Integração"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   3240
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker HrLimiItgr 
            Height          =   285
            Left            =   2640
            TabIndex        =   10
            Top             =   3240
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            DateIsNull      =   -1  'True
            Format          =   92864515
            UpDown          =   -1  'True
            CurrentDate     =   40833
         End
         Begin VB.ComboBox cboCanalVenda 
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Frame fraTipoDebitoCredito 
            Caption         =   "Tipo Lançamento Conta Corrente"
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   2160
            Width           =   2775
            Begin VB.OptionButton optCredito 
               Caption         =   "Crédito"
               Height          =   195
               Left            =   1200
               TabIndex        =   6
               Top             =   300
               Width           =   855
            End
            Begin VB.OptionButton optDebito 
               Caption         =   "Débito"
               Height          =   195
               Left            =   180
               TabIndex        =   5
               Top             =   300
               Width           =   795
            End
         End
         Begin VB.ComboBox cboTipoOperacao 
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   5595
         End
         Begin VB.TextBox txtTipoOperacao 
            Enabled         =   0   'False
            Height          =   285
            Left            =   150
            TabIndex        =   14
            Top             =   1080
            Width           =   5595
         End
         Begin VB.TextBox txtEmpresa 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   5595
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   480
            Width           =   5595
         End
         Begin VB.Label Label1 
            Caption         =   "Minuto(s)"
            Height          =   225
            Left            =   3480
            TabIndex        =   20
            Top             =   2925
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Intervalo entre reprocessamentos"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label lblCanalVenda 
            Caption         =   "Canal de Venda"
            Height          =   255
            Left            =   150
            TabIndex        =   18
            Top             =   1455
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operação"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   840
            Width           =   1065
         End
      End
      Begin MSComctlLib.ListView lvwConta 
         Height          =   2505
         Left            =   45
         TabIndex        =   11
         Top             =   210
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4419
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
      Left            =   120
      Top             =   6600
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
            Picture         =   "frmParamReprocCC.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamReprocCC.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamReprocCC.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamReprocCC.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamReprocCC.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamReprocCC.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamReprocCC.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   2280
      TabIndex        =   12
      Top             =   6840
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   582
      ButtonWidth     =   1958
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
            Caption         =   "Sai&r    "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParamReprocCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cadastro de Contas por Veículo Legal

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private strOperacao                         As String

Private strKeyItemSelected                  As String

Public lngBackOffice                        As Long

Private Const COL_EMPRESA                   As Integer = 0
Private Const COL_TIPOOPERACAO              As Integer = 1
Private Const COL_CANAL_VENDA               As Integer = 2
Private Const COL_TIPO_MOVTO                As Integer = 3
'Private Const COL_QTDE_REPROC               As Integer = 4
Private Const COL_INTE_REPR                 As Integer = 4
Private Const COL_IN_ATIVO                  As Integer = 5
Private Const COL_LIMI_ITGR                 As Integer = 6

Private Const strFuncionalidade             As String = "frmParamReprocCC"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

'Seleciona o item do listview de acordo com a seleção atual
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

'Formata as colunas do listview
Private Sub flFormataListView()

    lvwConta.ColumnHeaders.Add 1, , "Empresa", 1500, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 2, , "Tipo Operação", 3000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 3, , "Canal Venda", 1200, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 4, , "Tipo Movto.", 1000, lvwColumnLeft
    'lvwConta.ColumnHeaders.Add 5, , "Qtde Repro.", 1200, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 5, , "Intervalo", 1200, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 6, , "Ativo", 800, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 7, , "Horário Limite Integração", 2400, lvwColumnLeft

End Sub

'Salva as alterações efetuadas
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strRetorno                              As String
Dim strPropriedades                         As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

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
    
    If strOperacao = "Incluir" Then
        With xmlLer.documentElement
            strKeyItemSelected = "|" & .selectSingleNode("CD_EMPR").Text & _
                                 "|" & .selectSingleNode("TP_BKOF").Text & _
                                 "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                 "|" & .selectSingleNode("TP_OPER").Text & _
                                 "|" & .selectSingleNode("TP_CNAL_VEND").Text
        End With
    End If
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        Set objMIU = Nothing
        
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
            strKeyItemSelected = "|" & .selectSingleNode("CD_EMPR").Text & _
                                 "|" & .selectSingleNode("TP_BKOF").Text & _
                                 "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                 "|" & .selectSingleNode("TP_OPER").Text & _
                                 "|" & .selectSingleNode("TP_CNAL_VEND").Text
       
       End With
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "flSalvar", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Valida os campos preenchidos
Private Function flValidarCampos() As String

On Error GoTo ErrorHandler
    
    If strOperacao = gstrOperIncluir Then
        
        If cboEmpresa.ListIndex = -1 Then
            flValidarCampos = "Selecione a Empresa."
            cboEmpresa.SetFocus
            Exit Function
        End If
        
        If Not (optDebito.value Or optCredito.value) Then
            flValidarCampos = "Selecione o Tipo de Lançamento."
            Exit Function
        End If
        
        If cboTipoOperacao.ListIndex = -1 Then
            flValidarCampos = "Selecione o Tipo de Operação."
            cboTipoOperacao.SetFocus
            Exit Function
        End If
    End If
    
    
    'FREITAS - 14/07/2010 - Este validação não é mais necessária.
'    With numQtdeReproc
'        If .Valor = 0 Then
'            flValidarCampos = "Informe a número de tentativas de reprocessamentos."
'            .SelStart = 0
'            .SelLength = Len(.Valor)
'            .SetFocus
'            Exit Function
'        End If
'    End With
    
    With numIntervalo
        If Int(.Text) = 0 Then
            flValidarCampos = "Informe o intervalo entre as tentativas de reprocessamentos."
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

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()

On Error GoTo ErrorHandler

    strOperacao = "Incluir"

    cboEmpresa.ListIndex = -1
    cboEmpresa.Visible = True
    txtEmpresa.Visible = False
    
    cboTipoOperacao.ListIndex = -1
    cboTipoOperacao.Visible = True
    txtTipoOperacao.Visible = False
    
    cboCanalVenda.ListIndex = -1
    cboCanalVenda.Enabled = True
    
    fraTipoDebitoCredito.Enabled = True
    optCredito.value = False
    optDebito.value = False

    numIntervalo.Text = 0
    'numQtdeReproc.Valor = 0
    
    chkHrLimiItgr.value = vbUnchecked
    chkHrLimiItgr.Enabled = True

    HrLimiItgr.Enabled = False
    HrLimiItgr.Hour = 0
    HrLimiItgr.Minute = 0
    
    chkAtivo.value = vbChecked
    chkAtivo.Enabled = False

    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    
    cboEmpresa.SetFocus
        
Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, "frmGrupoUsuario", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Preenche os campos em tela com o conteúdo do documento XML
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strChaveRegistro                        As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    If lvwConta.SelectedItem Is Nothing Then
        flLimpaCampos
        Exit Sub
    End If

    strChaveRegistro = lvwConta.SelectedItem.Key

    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//CD_EMPR").Text = Split(strChaveRegistro, "|")(0)
        .selectSingleNode("//TP_BKOF").Text = Split(strChaveRegistro, "|")(1)
        .selectSingleNode("//IN_LANC_DEBT_CRED").Text = Split(strChaveRegistro, "|")(2)
        .selectSingleNode("//TP_OPER").Text = Split(strChaveRegistro, "|")(3)
        .selectSingleNode("//TP_CNAL_VEND").Text = Split(strChaveRegistro, "|")(4)
        
    End With

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call xmlLer.loadXML(objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    With xmlLer.documentElement
        
        txtEmpresa.Visible = True
        cboEmpresa.Visible = False
        
        txtEmpresa.Text = .selectSingleNode("CD_EMPR").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_Empresa[CO_EMPR=" & .selectSingleNode("CD_EMPR").Text & "]/NO_REDU_EMPR").Text
                
        txtTipoOperacao.Visible = True
        cboTipoOperacao.Visible = False
        
        If .selectSingleNode("TP_OPER").Text > 0 Then
            txtTipoOperacao.Text = .selectSingleNode("TP_OPER").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_TipoOperacao[TP_OPER='" & .selectSingleNode("TP_OPER").Text & "']/NO_TIPO_OPER").Text
        Else
            txtTipoOperacao.Text = "<Padrão>"
        End If
        fraTipoDebitoCredito.Enabled = False
        
        Call fgSearchItemCombo(cboCanalVenda, 0, xmlLer.documentElement.selectSingleNode("TP_CNAL_VEND").Text)
        cboCanalVenda.Enabled = False
        
        Select Case Val(.selectSingleNode("IN_LANC_DEBT_CRED").Text)
            Case enumTipoDebitoCredito.Credito
                optCredito.value = True
            Case enumTipoDebitoCredito.Debito
                optDebito.value = True
        End Select
        
        'numQtdeReproc.Valor = .selectSingleNode("QT_REPR_CNTA_CRRT").Text
        numIntervalo.Text = .selectSingleNode("QT_HORA_INTL_REPR").Text
        strUltimaAtualizacao = .selectSingleNode("DH_ULTI_ATLZ").Text
        
        chkHrLimiItgr.Enabled = True
        
        If .selectSingleNode("HO_LIMI_ITGR").Text <> "00:00:00" Then
            HrLimiItgr.Hour = Mid(.selectSingleNode("HO_LIMI_ITGR").Text, 9, 2)
            HrLimiItgr.Minute = Mid(.selectSingleNode("HO_LIMI_ITGR").Text, 11, 2)
            chkHrLimiItgr.value = vbChecked
            HrLimiItgr.Enabled = True
        Else
            HrLimiItgr.Hour = 0
            HrLimiItgr.Minute = 0
            chkHrLimiItgr.value = vbUnchecked
            HrLimiItgr.Enabled = False
        End If

        chkAtivo.Enabled = True

        If .selectSingleNode("IN_ATIV").Text = enumIndicadorSimNao.sim Then
            chkAtivo.value = vbChecked
        Else
            chkAtivo.value = vbUnchecked
        End If
        
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
        
    End With

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmGrupoUsuario", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Preenche as tags do documento XML com o conteúdo dos controles da tela
Private Function flInterfaceToXml() As String

Dim lngTipoDebitoCredito                    As Long

On Error GoTo ErrorHandler

    With xmlLer.documentElement

         .selectSingleNode("@Operacao").Text = strOperacao
         
         If strOperacao <> gstrOperExcluir Then
         
            If strOperacao = gstrOperIncluir Then
               .selectSingleNode("CD_EMPR").Text = fgObterCodigoCombo(cboEmpresa.Text)
               .selectSingleNode("TP_BKOF").Text = gintTipoBackoffice
               
                Select Case True
                    Case optCredito.value
                        lngTipoDebitoCredito = enumTipoDebitoCredito.Credito
                    Case optDebito.value
                        lngTipoDebitoCredito = enumTipoDebitoCredito.Debito
                End Select
               
               .selectSingleNode("IN_LANC_DEBT_CRED").Text = lngTipoDebitoCredito
               .selectSingleNode("TP_OPER").Text = fgObterCodigoCombo(cboTipoOperacao.Text)
               
               .selectSingleNode("TP_CNAL_VEND").Text = fgObterCodigoCombo(cboCanalVenda.Text)
               
            End If
            
            '.selectSingleNode("QT_REPR_CNTA_CRRT").Text = numQtdeReproc.Valor
            .selectSingleNode("QT_HORA_INTL_REPR").Text = Int(numIntervalo.Text)
            .selectSingleNode("DH_ULTI_ATLZ").Text = strUltimaAtualizacao
            
            If chkHrLimiItgr.value = vbChecked Then
                .selectSingleNode("HO_LIMI_ITGR").Text = Format(HrLimiItgr.value, "DD/MM/YYYY HH:MM:SS")
            Else
                .selectSingleNode("HO_LIMI_ITGR").Text = vbNullString
            End If

            If chkAtivo.value = vbChecked Then
                .selectSingleNode("IN_ATIV").Text = enumIndicadorSimNao.sim
            Else
                .selectSingleNode("IN_ATIV").Text = enumIndicadorSimNao.Nao
            End If
         
         End If

    End With

Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0

End Function

'Inicializa os controles e variáveis
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

    Call fgCarregarCombos(cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
      
    Call flCarregarComboTipoOperacao
    
    Call flCarregarcboCanalVenda
    
    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_ParametroReprocCC").xml
    End If

    Set objMIU = Nothing
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub


'Preencher o combo de Canal de Venda
Private Sub flCarregarcboCanalVenda()

    cboCanalVenda.Clear
    cboCanalVenda.AddItem enumCanalDeVenda.Nenhum & " - " & fgDescricaoCanalVenda(enumCanalDeVenda.Nenhum)
    cboCanalVenda.AddItem enumCanalDeVenda.SGC & " - " & fgDescricaoCanalVenda(enumCanalDeVenda.SGC)
    cboCanalVenda.AddItem enumCanalDeVenda.SGM & " - " & fgDescricaoCanalVenda(enumCanalDeVenda.SGM)
    cboCanalVenda.ListIndex = -1

End Sub


'Define o número máximo de caracteres permitidos nos controles
Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0
End Sub

'Carrega o conteúdo do ListView
Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlTipoOperacao                         As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim strNomeTipoLiquidacao                   As String
Dim strDescTipoMovimento                    As String

Dim strTempChave                            As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant


On Error GoTo ErrorHandler

    fgCursor True

    lvwConta.ListItems.Clear
    lvwConta.HideSelection = False

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")

    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParametroReprocCC/@Operacao").Text = "LerTodos"
    Call xmlLerTodos.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParametroReprocCC").xml, _
                                             vntCodErro, _
                                             vntMensagemErro))

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMIU = Nothing

    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_ParametroReprocCC/*")
        With xmlDomNode
            
            
            strTempChave = .selectSingleNode("CD_EMPR").Text & _
                           "|" & .selectSingleNode("TP_BKOF").Text & _
                           "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                           "|" & .selectSingleNode("TP_OPER").Text & _
                           "|" & .selectSingleNode("TP_CNAL_VEND").Text

            Set objListItem = lvwConta.ListItems.Add(, strTempChave)
            
            objListItem.Text = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & .selectSingleNode("CD_EMPR").Text & "']/NO_REDU_EMPR").Text
            
            Select Case .selectSingleNode("IN_LANC_DEBT_CRED").Text
                Case enumTipoDebitoCredito.Credito
                    strDescTipoMovimento = "Crédito"
                Case enumTipoDebitoCredito.Debito
                    strDescTipoMovimento = "Débito"
            End Select

            objListItem.SubItems(COL_TIPO_MOVTO) = strDescTipoMovimento
            
            objListItem.SubItems(COL_CANAL_VENDA) = fgDescricaoCanalVenda(.selectSingleNode("TP_CNAL_VEND").Text)
            
            objListItem.SubItems(COL_INTE_REPR) = .selectSingleNode("QT_HORA_INTL_REPR").Text
            'objListItem.SubItems(COL_QTDE_REPROC) = .selectSingleNode("QT_REPR_CNTA_CRRT").Text
            
            If .selectSingleNode("IN_ATIV").Text = enumIndicadorSimNao.sim Then
                objListItem.SubItems(COL_IN_ATIVO) = "Sim"
            Else
                objListItem.SubItems(COL_IN_ATIVO) = "Não"
            End If
            
            If .selectSingleNode("HO_LIMI_ITGR").Text <> "00:00:00" Then
                objListItem.SubItems(COL_LIMI_ITGR) = Mid(.selectSingleNode("HO_LIMI_ITGR").Text, 9, 2) & ":" & Mid(.selectSingleNode("HO_LIMI_ITGR").Text, 11, 2)
            Else
                objListItem.SubItems(COL_LIMI_ITGR) = ""
            End If
            
            If .selectSingleNode("TP_OPER").Text > 0 Then
            
                For Each xmlTipoOperacao In xmlMapaNavegacao.selectNodes("//Repeat_TipoOperacao/Grupo_TipoOperacao")
                    If Not xmlTipoOperacao Is Nothing Then
                        If CLng(xmlTipoOperacao.selectSingleNode("TP_OPER").Text) = .selectSingleNode("TP_OPER").Text Then
                             objListItem.SubItems(COL_TIPOOPERACAO) = xmlTipoOperacao.selectSingleNode("TP_OPER").Text & _
                             " - " & xmlTipoOperacao.selectSingleNode("NO_TIPO_OPER").Text
                            Exit For
                        End If
                    Else
                        objListItem.SubItems(COL_TIPOOPERACAO) = ""
                        Exit For
                    End If
                Next

            Else
                objListItem.SubItems(COL_TIPOOPERACAO) = ""
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

Private Sub chkHrLimiItgr_Click()

    If chkHrLimiItgr.value = vbChecked Then
        HrLimiItgr.Enabled = True
    Else
        HrLimiItgr.Enabled = False
    End If

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
    Call flDefinirTamanhoMaximoCampos
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

'Atualiza o conteúdo da tela
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

Private Sub numIntervalo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler

    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
    
Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - numIntervalo_KeyPress"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Dim operacao As String

    fgCursor True

    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case gstrSalvar
            If strOperacao = gstrOperAlterar Then
                operacao = "alteração"
            Else
                operacao = "inclusão"
            End If
            If MsgBox("Confirma " & operacao & " do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                Call flSalvar
                If strOperacao = gstrOperAlterar Then
                   flPosicionaItemListView
                End If
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

Private Sub flCarregarComboTipoOperacao()

Dim objNode                                 As MSXML2.IXMLDOMNode
Dim blnCanalVenda                           As Boolean

On Error GoTo ErrorHandler

       
    cboTipoOperacao.Clear
    
    For Each objNode In xmlMapaNavegacao.selectNodes("//Repeat_TipoOperacao/Grupo_TipoOperacao")
        If Not objNode Is Nothing Then
            cboTipoOperacao.AddItem objNode.selectSingleNode("TP_OPER").Text & " - " & objNode.selectSingleNode("NO_TIPO_OPER").Text
        End If
    Next
    
    cboTipoOperacao.ListIndex = -1
    cboTipoOperacao.Enabled = cboTipoOperacao.ListCount > 1
    
    Exit Sub
ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flCarregarComboTipoOperacao"
End Sub

