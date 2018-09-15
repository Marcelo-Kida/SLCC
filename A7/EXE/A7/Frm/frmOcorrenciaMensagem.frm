VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOcorrenciaMensagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parâmetros de Notificação de Ocorrência"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6480
   Begin VB.Frame fraDetalhe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2250
      Left            =   15
      TabIndex        =   3
      Top             =   3510
      Width           =   6405
      Begin VB.ComboBox cboEnvioEmail 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   825
      End
      Begin VB.ComboBox cboSeveridade 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox txtEnderecoEmail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1770
         Width           =   6195
      End
      Begin VB.Label lblParamNotific 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-mail ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label lblParamNotific 
         AutoSize        =   -1  'True
         Caption         =   "E-mail ( separados por ; )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2145
      End
      Begin VB.Label lblParamNotific 
         AutoSize        =   -1  'True
         Caption         =   "Grau Severidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1440
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   4695
      TabIndex        =   6
      Top             =   5820
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Excluir"
            Key             =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   30
      Top             =   5565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":0112
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":042C
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":077E
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":0890
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":0BAA
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":0EC4
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOcorrenciaMensagem.frx":11DE
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstOcorrenciaMensagem 
      Height          =   3450
      Left            =   30
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   15
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   6085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cód. Ocorrência"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome da Ocorrência"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Grau Severidade"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Envia E-mail"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmOcorrenciaMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela manutenção dos parâmetros de notificação de ocorrências.
Option Explicit

Private xmlMapaNavegacao                     As MSXML2.DOMDocument40
Private xmlLer                               As MSXML2.DOMDocument40
Private strOperacao                          As String
Private strKeyItemSelected                   As String
Private Const strFuncionalidade              As String = "frmOcorrenciaMensagem"

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

'Carregar listview com tipos de ocorrências.
Private Sub flCarregarListView()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim objNode             As MSXML2.IXMLDOMNode
Dim objListItem         As MSComctlLib.ListItem
Dim strLerTodos         As String
Dim xmlLerTodos         As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
        
    lstOcorrenciaMensagem.ListItems.Clear
    lstOcorrenciaMensagem.HideSelection = False
        
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_OcorrenciaMensagem/@Operacao").Text = "LerTodos"
    strLerTodos = objMiu.Executar(xmlMapaNavegacao.selectSingleNode("//Grupo_OcorrenciaMensagem").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlLerTodos.loadXML(strLerTodos)
    
    For Each objNode In xmlLerTodos.selectSingleNode("//Repeat_OcorrenciaMensagem").childNodes
        
        Set objListItem = lstOcorrenciaMensagem.ListItems.Add(, "K" & Trim$(objNode.selectSingleNode("CO_OCOR_MESG").Text), Trim$(objNode.selectSingleNode("CO_OCOR_MESG").Text))
        
        objListItem.SubItems(1) = Trim$(objNode.selectSingleNode("DE_OCOR_MESG").Text)
        objListItem.SubItems(2) = Trim$(objNode.selectSingleNode("TP_GRAU_SEVE").Text)
        
        If Trim$(objNode.selectSingleNode("IN_ENVI_EMAIL").Text) = "S" Then
            objListItem.SubItems(3) = "Sim"
        Else
            objListItem.SubItems(3) = "Não"
        End If
    Next
    
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flCarregarListView", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Definir tamanho máximo para campos pertinentes ao cadastramento de notificações.
Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler
                
    With xmlMapaNavegacao
        txtEnderecoEmail.MaxLength = .selectSingleNode("//Grupo_Propriedades/Grupo_OcorrenciaMensagem/TX_ENDE_EMAIL_NOTI").attributes.getNamedItem("Tamanho").Text
    End With
    
    Exit Sub

ErrorHandler:
    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flDefinirTamanhoMaximoCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInit()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strMapaNavegacao    As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    strMapaNavegacao = objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmOcorrenciaMensagem", "flInit")
    End If
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlLer.loadXML xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_OcorrenciaMensagem").xml
    
    Me.cboEnvioEmail.AddItem "Sim"
    Me.cboEnvioEmail.AddItem "Não"
    
    Me.cboSeveridade.AddItem "1 - Verde"
    Me.cboSeveridade.ItemData(Me.cboSeveridade.NewIndex) = 1
    Me.cboSeveridade.AddItem "2 - Amarelo"
    Me.cboSeveridade.ItemData(Me.cboSeveridade.NewIndex) = 2
    Me.cboSeveridade.AddItem "3 - Laranja"
    Me.cboSeveridade.ItemData(Me.cboSeveridade.NewIndex) = 3
    Me.cboSeveridade.AddItem "4 - Vermelho"
    Me.cboSeveridade.ItemData(Me.cboSeveridade.NewIndex) = 4
    
    Exit Sub

ErrorHandler:
    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flInit", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

'Mover valores do formulário para XML para envio ao objeto de negócio.
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler
        
    With xmlLer.documentElement
        .selectSingleNode("@Operacao").Text = strOperacao
        .selectSingleNode("CO_OCOR_MESG").Text = lstOcorrenciaMensagem.SelectedItem.Text
        .selectSingleNode("TP_GRAU_SEVE").Text = cboSeveridade.ItemData(cboSeveridade.ListIndex)
        .selectSingleNode("IN_ENVI_EMAIL").Text = Left$(cboEnvioEmail.Text, 1)
        .selectSingleNode("TX_ENDE_EMAIL_NOTI").Text = Trim$(txtEnderecoEmail.Text)
    End With
    
    Exit Function

ErrorHandler:
    
    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flInterfaceToXml", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Limpar campos do formulário.
Private Sub flLimparCampos()

On Error GoTo ErrorHandler
    
    lstOcorrenciaMensagem.Sorted = False
    cboEnvioEmail.ListIndex = -1
    cboSeveridade.ListIndex = -1
    txtEnderecoEmail.Text = ""
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flLimparCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Salvar as informações correntes da notificação de ocorrência.
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strRetorno          As String
Dim strKey              As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
    
    If strRetorno <> "" Then
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
        
    Call fgCursor(True)
    
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    With xmlLer.documentElement
        .selectSingleNode("@Operacao").Text = "Ler"
        .selectSingleNode("CO_OCOR_MESG").Text = lstOcorrenciaMensagem.SelectedItem.Text
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlLer.loadXML objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
        
    strKey = lstOcorrenciaMensagem.SelectedItem.Key
    strKeyItemSelected = strKey
    
    flCarregarListView
    
    lstOcorrenciaMensagem.ListItems(strKey).Selected = True
    lstOcorrenciaMensagem.ListItems(strKey).EnsureVisible
    
    Call fgCursor(False)
        
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing
    Call fgCursor(False)

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Validar os valores informados para a notificação de ocorrência.
Private Function flValidarCampos() As String
    
Dim vntEmail                                As Variant
Dim lngCont                                 As Long
    
    If cboSeveridade.ListIndex = -1 Then
        flValidarCampos = "Selecione o Grau de Severidade."
        cboSeveridade.SetFocus
        Exit Function
    End If
    
    If cboEnvioEmail.ListIndex = -1 Then
        flValidarCampos = "Selecione a opção de envio ou não, de e-mail de notificação."
        cboEnvioEmail.SetFocus
        Exit Function
    End If
    
    If cboEnvioEmail.Text = "Sim" Then
        If Trim(txtEnderecoEmail.Text) = "" Then
            flValidarCampos = "Digite o endereço de e-mail para notificação."
            txtEnderecoEmail.SetFocus
            Exit Function
        End If
    End If
        
    vntEmail = Split(txtEnderecoEmail, ";", , vbBinaryCompare)
    
    For lngCont = 0 To UBound(vntEmail)
        If Not flValidaEmail(vntEmail(lngCont)) Then
            flValidarCampos = "E-mail inválido: " & vntEmail(lngCont)
            Exit Function
        End If
    Next
        
    flValidarCampos = ""
    
End Function

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    With xmlLer.documentElement
        .selectSingleNode("@Operacao").Text = "Ler"
        .selectSingleNode("CO_OCOR_MESG").Text = lstOcorrenciaMensagem.SelectedItem.Text
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlLer.loadXML objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
                
    Call fgSearchItemCombo(Me.cboSeveridade, 0, xmlLer.documentElement.selectSingleNode("TP_GRAU_SEVE").Text)
    Call fgSearchItemCombo(Me.cboEnvioEmail, 0, xmlLer.documentElement.selectSingleNode("IN_ENVI_EMAIL").Text)
    
    txtEnderecoEmail.Text = xmlLer.documentElement.selectSingleNode("TX_ENDE_EMAIL_NOTI").Text
        
    strKeyItemSelected = "K" & xmlLer.documentElement.selectSingleNode("CO_OCOR_MESG").Text
        
    Exit Sub

ErrorHandler:
    
    Set xmlLer = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmOcorrenciaMensagem", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

Private Sub cboEnvioEmail_Click()
    
    If cboEnvioEmail.Text = "Sim" Then
        Me.txtEnderecoEmail.Enabled = True
        'If Me.Visible Then Me.txtEnderecoEmail.SetFocus
    Else
        Me.txtEnderecoEmail.Text = vbNullString
        Me.txtEnderecoEmail.Enabled = False
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    strOperacao = "Incluir"
        
    fgCenterMe Me
    Me.Icon = mdiBUS.Icon
    Me.Show
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    
    DoEvents
    
    Call flLimparCampos
    
    Call fgCursor(True)
    
    Call flInit
    Call flDefinirTamanhoMaximoCampos
    Call flCarregarListView
    
    If lstOcorrenciaMensagem.ListItems.Count > 0 Then
        lstOcorrenciaMensagem_ItemClick lstOcorrenciaMensagem.ListItems.Item(1)
    End If
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    mdiBUS.uctLogErros.MostrarErros Err, "frmOcorrenciaMensagem - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    
End Sub

Private Sub lstOcorrenciaMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Call fgClassificarListview(lstOcorrenciaMensagem, ColumnHeader.Index)
    
End Sub

Private Sub lstOcorrenciaMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
            
    Call fgCursor(True)
    Call flLimparCampos
    Call flXmlToInterface
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmOcorrenciaMensagem - lstParamNotific_ItemClick"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Salvar"
            strOperacao = "Alterar"
            Call flSalvar
        Case "Sair"
            Unload Me
            strOperacao = ""
    End Select
    
    If strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
    Exit Sub

ErrorHandler:
    Call fgCursor(False)
    mdiBUS.uctLogErros.MostrarErros Err, "frmOcorrenciaMensagem - tlbCadastro_ButtonClick"
    
    flCarregarListView
    
    If strOperacao = "Excluir" Then
        flLimparCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
End Sub

'Validar valores informados para email de destino da notificação.
Private Function flValidaEmail(ByVal pstrMail As String) As Boolean
        
Dim strTlds                            As Variant
Dim strLeft                            As String
    
    flValidaEmail = True
    
    If Mid(Trim(pstrMail), 1, 1) = "@" Then
        flValidaEmail = False
    ElseIf InStr(1, pstrMail, "@") = 0 Then
        flValidaEmail = False
    ElseIf InStr(1, pstrMail, ".") = 0 Then
        flValidaEmail = False
    ElseIf Not (InStr(1, pstrMail, " ") = 0) Then
        flValidaEmail = False
    Else
        strTlds = Split(pstrMail, ".")
        strLeft = strTlds(UBound(strTlds))
        strLeft = LCase(strLeft)
        'Domain is a TLD.
        '??|com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum
        If Not (Len(strLeft) = 2) Then
            If Len(strLeft) = 3 Then
                If Not (strLeft = "com" Or strLeft = "net" Or strLeft = "org" Or strLeft = "edu" Or strLeft = "int" Or strLeft = "mil" Or strLeft = "gov" Or strLeft = "biz" Or strLeft = "pro") Then
                    flValidaEmail = False
                End If
            ElseIf Len(strLeft) = 4 Then
                If Not (strLeft = "arpa" Or strLeft = "aero" Or strLeft = "name" Or strLeft = "coop" Or strLeft = "info") Then
                    flValidaEmail = False
                End If
            ElseIf Len(strLeft) = 6 Then
                If Not (strLeft = "museum") Then
                    flValidaEmail = False
                End If
            Else
                flValidaEmail = False
            End If
        End If
    End If

End Function

'Posicionar item no listview de ocorrências.
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

    If lstOcorrenciaMensagem.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstOcorrenciaMensagem.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstOcorrenciaMensagem_ItemClick objListItem
           lstOcorrenciaMensagem.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimparCampos
    End If

End Sub

