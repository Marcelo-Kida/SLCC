VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroIdentificadorPartCamara 
   Caption         =   "Cadastro - Identificador Participante na Câmara Veículo Legal"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   9675
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8040
      Top             =   5640
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
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroIdentificadorPartCamara.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCadastro 
      Caption         =   "Cadastro"
      Height          =   2655
      Left            =   3000
      TabIndex        =   0
      Top             =   3000
      Width           =   6615
      Begin VB.TextBox numIdParticipante 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1800
         Width           =   4335
      End
      Begin VB.ComboBox cboLocalLiquidacao 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   4335
      End
      Begin VB.ComboBox cboTipoIdentificadorPartCamara 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtVeiculoLegal 
         BackColor       =   &H80000003&
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtSistema 
         BackColor       =   &H80000003&
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   4
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identificador Participante"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Veículo Legal"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sistema"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo Identificador"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L&ocal Liquidacao"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvwIdentificador 
      Height          =   3015
      Left            =   3000
      TabIndex        =   10
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   5820
      Width           =   9675
      _ExtentX        =   17066
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
   Begin MSComctlLib.ListView lvwVeiculoLegal 
      Height          =   5655
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2900
      _ExtentX        =   5106
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCadastroIdentificadorPartCamara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intRefresh                          As Integer
Private strOperacao                         As String
Private blnEditMode                         As Boolean

Private xmlLer                              As MSXML2.DOMDocument40
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmCadastroIdentificadorPartCamara"

Private lngIndexClassifList                 As Long

'' Formatas os títulos das colunas do grid
Private Sub flPreencherHeadersLvw()

On Error GoTo ErrorHandler

    With lvwVeiculoLegal.ColumnHeaders
        .Clear
        .Add 1, , "Veiculo Legal", lvwVeiculoLegal.Width
    End With

    With lvwIdentificador.ColumnHeaders
        .Clear
        .Add 1, , "Local Liquidacao", 1500
        .Add 2, , "Tipo Identificador", 4200
        .Add 3, , "Participante Câmara", 3200
        '.Add 4, , "Última Atualização", 2000
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPreencherHeadersLvw", 0

End Sub

Private Sub Form_Load()

Dim xmlDomSistema As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

    fgCursor True
    flCarregarIdentificadores
    flInicializar
    flPreencherHeadersLvw
    flCarregarListaVeicLegal
    
    fgCursor

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With lvwVeiculoLegal
        .Top = 0
        .Left = 60
        .Height = Me.ScaleHeight - 300
        'ocupa 40% do formulário
        .Width = Me.Width * 0.4
        .ColumnHeaders(1).Width = .Width
    End With
    
    With fraCadastro
        .Left = lvwVeiculoLegal.Left + lvwVeiculoLegal.Width + 60
        tlbCadastro.Top = Me.ScaleHeight - tlbCadastro.Height
        .Top = tlbCadastro.Top - .Height - 60
        .Width = Me.ScaleWidth - .Left
    End With
    
    With lvwIdentificador
        .Top = 0
        .Left = fraCadastro.Left
        .Width = Me.Width - 100
        .Height = fraCadastro.Top - 120
    End With
    
End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub lvwIdentificador_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    With Item
        Call fgSearchItemCombo(cboLocalLiquidacao, , Item.Text)
        cboTipoIdentificadorPartCamara.ListIndex = flRetornaIndicePorItemData(cboTipoIdentificadorPartCamara, flObterCodigo(.SubItems(1)))
        numIdParticipante.Text = .SubItems(2)
        blnEditMode = True
        cboLocalLiquidacao.Locked = True
        cboTipoIdentificadorPartCamara.Locked = True
    End With

End Sub

Private Sub lvwVeiculoLegal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwVeiculoLegal, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

    Exit Sub

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwVeiculoLegal_ColumnClick"
    
End Sub

Private Sub lvwVeiculoLegal_ItemClick(ByVal Item As MSComctlLib.ListItem)
 
On Error Resume Next
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = True
    
    With Item
        txtVeiculoLegal.Text = .Text
        txtSistema.Text = .Tag
        cboLocalLiquidacao.ListIndex = -1
        numIdParticipante.Text = 0
        cboTipoIdentificadorPartCamara.ListIndex = -1
        flCarregarDados
    End With
        
End Sub

Private Sub numIdParticipante_KeyPress(KeyAscii As Integer)
    
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case gstrSalvar
            strOperacao = IIf(blnEditMode, gstrOperAlterar, gstrOperIncluir)
            Call flSalvar
        Case gstrOperExcluir
        
            If Not blnEditMode Then
                MsgBox "Selecione um item a ser excluído.", vbOKCancel + vbInformation, Me.Caption
                fgCursor False
                Exit Sub
            End If
            
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
        
    mdiLQS.uctlogErros.MostrarErros Err, "frmTipoJustificativaConciliacao - tlbCadastro_ButtonClick", Me.Caption
       
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    End If

End Sub

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()
        
Dim ctrl As Control
    
On Error GoTo ErrorHandler

    For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is Number Then
            ctrl.Valor = 0
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
        End If
    Next
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    lvwIdentificador.ListItems.Clear
    blnEditMode = False
    cboLocalLiquidacao.Locked = False
    cboTipoIdentificadorPartCamara.Locked = False
    lvwVeiculoLegal.SetFocus
    
Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

'' É acionado através no botão 'Salvar' da barra de ferramentas.
''
'' Tem como função, encaminhar a solicitação (Atualização dos dados na tabela) à
'' camada controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
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
    If objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Set objMIU = Nothing
        Call flLimpaCampos
        MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    End If

    blnEditMode = False
    cboLocalLiquidacao.Locked = False
    cboTipoIdentificadorPartCamara.Locked = False
    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'Valida o preenchimento dos campos
Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If Trim(txtVeiculoLegal.Text) = "" Then
        flValidarCampos = "Selecione um Veículo Legal."
        lvwVeiculoLegal.SetFocus
        Exit Function
    End If
    
    If cboTipoIdentificadorPartCamara.ListIndex = -1 Then
        flValidarCampos = "Informe o Tipo de Identificador Participante Câmara."
        cboTipoIdentificadorPartCamara.SetFocus
        Exit Function
    End If
    
    If cboLocalLiquidacao.ListIndex = -1 Then
        flValidarCampos = "Informe o Código do Liquidante."
        cboLocalLiquidacao.SetFocus
        Exit Function
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Preenche o conteúdo do XML com o conteúdo dos campos apresentados em tela
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = strOperacao
        .selectSingleNode("CO_VEIC_LEGA").Text = flObterCodigo(txtVeiculoLegal.Text)
        .selectSingleNode("SG_SIST").Text = txtSistema.Text
        .selectSingleNode("CO_LOCA_LIQU").Text = fgObterCodigoCombo(cboLocalLiquidacao.Text)
        .selectSingleNode("TP_IDEF_PART_CAMR").Text = cboTipoIdentificadorPartCamara.ItemData(cboTipoIdentificadorPartCamara.ListIndex)
        .selectSingleNode("ID_PART_CAMR").Text = numIdParticipante.Text
    End With
    
    Exit Function

ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
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
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If

    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/IdentificadorCamara").xml
    End If
    
    Call fgCarregarCombos(cboLocalLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "SG_LOCA_LIQU")
    
    Set objMIU = Nothing
        
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

Private Sub flCarregarIdentificadores()
    
    With cboTipoIdentificadorPartCamara
    '    ContaEmissora = 1
        .AddItem "Conta Emissora"
        .ItemData(.NewIndex) = 1
    '    ContaPropriaBancoMandatario = 2
        .AddItem "Conta Propria Banco Mandatario"
        .ItemData(.NewIndex) = 2
    '    ContaCaucao = 3
        .AddItem "Conta Caução"
        .ItemData(.NewIndex) = 3
    '    ContaAnuenteContrato = 4
        .AddItem "Conta Anuente Contrato"
        .ItemData(.NewIndex) = 4
    '    ContaCedenteContrato = 5
        .AddItem "Conta Cedente Contrato"
        .ItemData(.NewIndex) = 5
    '    ContaAdquirenteContrato = 6
        .AddItem "Conta Adquirente Contrato"
        .ItemData(.NewIndex) = 6
    End With

End Sub

'' Carrega todos os veículos legais na lista
Private Sub flCarregarListaVeicLegal()

#If EnableSoap = 1 Then
    Dim objVeiculoLegal     As MSSOAPLib30.SoapClient30
#Else
    Dim objVeiculoLegal     As A8MIU.clsVeiculoLegal
#End If

Dim xmlRetorno              As MSXML2.DOMDocument40
Dim strRetorno              As String
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim objListItem             As ListItem
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    fgCursor True
    fgLockWindow Me.hwnd
    
    Set objVeiculoLegal = fgCriarObjetoMIU("A8MIU.clsVeiculoLegal")
    
    strRetorno = objVeiculoLegal.ObterDetalheVeiculoLegal("", vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lvwVeiculoLegal.ListItems.Clear
    
    If strRetorno = vbNullString Then
        fgCursor
        fgLockWindow
        Exit Sub
    End If
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    xmlRetorno.loadXML strRetorno

    For Each objDomNode In xmlRetorno.documentElement.childNodes
    
        Set objListItem = lvwVeiculoLegal.ListItems.Add(, "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & objDomNode.selectSingleNode("SG_SIST").Text)
        objListItem.Text = fgSelectSingleNode(objDomNode, "CO_VEIC_LEGA").Text & " - " & fgSelectSingleNode(objDomNode, "NO_VEIC_LEGA").Text
        objListItem.Tag = fgSelectSingleNode(objDomNode, "SG_SIST").Text
        
    Next objDomNode

    Call fgClassificarListview(Me.lvwVeiculoLegal, lngIndexClassifList, True)
    
    Set objVeiculoLegal = Nothing
    Set xmlRetorno = Nothing

    fgLockWindow 0
    fgCursor

Exit Sub
ErrorHandler:
    
    fgLockWindow 0
    fgCursor
    Set objVeiculoLegal = Nothing
    Set xmlRetorno = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListaVeicLegal", 0
    
End Sub

Private Function flRetornaIndicePorItemData(ByRef pCombo As ComboBox, _
                                            ByVal pData) As Integer
    
Dim i As Integer
    
    flRetornaIndicePorItemData = -1
    
    For i = 0 To pCombo.ListCount - 1
        If pCombo.ItemData(i) = pData Then
            flRetornaIndicePorItemData = i
            Exit For
        End If
    Next
    
End Function

Private Function flRetornaDescricaoPorItemData(ByRef pCombo As ComboBox, _
                                               ByVal pData) As String
    
Dim i As Integer
    
    For i = 0 To pCombo.ListCount - 1
        If pCombo.ItemData(i) = pData Then
            flRetornaDescricaoPorItemData = pCombo.List(i)
            Exit For
        End If
    Next
    
End Function

Private Sub flCarregarDados()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsIdPartCamaraVeicLegal
#End If

Dim xmlDomNode             As MSXML2.IXMLDOMNode
Dim strPropriedades        As String
Dim strLerTodos            As String
Dim xmlLerTodos            As MSXML2.DOMDocument40
Dim objListItem            As ListItem
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    lvwIdentificador.ListItems.Clear
    
    With lvwVeiculoLegal
        If .SelectedItem Is Nothing Then
            MsgBox "Selecione um veículo legal."
            lvwVeiculoLegal.SetFocus
            Exit Sub
        End If
        
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsIdPartCamaraVeicLegal")
        strLerTodos = objMIU.Ler(flObterCodigo(.SelectedItem.Text), .SelectedItem.Tag, 0, 0, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objMIU = Nothing
        
        If strLerTodos = "" Then Exit Sub
        
        Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
        If xmlLerTodos.loadXML(strLerTodos) Then
            For Each xmlDomNode In xmlLerTodos.selectNodes("Repeat_IdentificadorCamara/Grupo_IdentificadorCamara")
                Set objListItem = lvwIdentificador.ListItems.Add(, "|" & xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "|" & xmlDomNode.selectSingleNode("TP_IDEF_PART_CAMR").Text & "|" & xmlDomNode.selectSingleNode("ID_PART_CAMR").Text)
                objListItem.Text = fgSelectSingleNode(xmlDomNode, "CO_LOCA_LIQU").Text
                objListItem.SubItems(1) = fgSelectSingleNode(xmlDomNode, "TP_IDEF_PART_CAMR").Text & " - " & flRetornaDescricaoPorItemData(cboTipoIdentificadorPartCamara, fgSelectSingleNode(xmlDomNode, "TP_IDEF_PART_CAMR").Text)
                objListItem.SubItems(2) = fgSelectSingleNode(xmlDomNode, "ID_PART_CAMR").Text
                'objListItem.SubItems(3) = fgDtHrXML_To_Interface(xmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
            Next
        End If
        Set xmlLerTodos = Nothing
        
    End With
    
    Exit Sub
    
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmCadastroIdentificadorPartCamara", "flCarregarDados", 0

End Sub

Private Function flObterCodigo(ByVal pTexto As String) As String
    
Dim arrStr As Variant
    
    arrStr = Split(pTexto, "-")
    
    If Not IsEmpty(arrStr) Then
        flObterCodigo = Trim(arrStr(0))
    End If
       
End Function

