VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemCaixaTipoOperacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Itens de Caixa x Tipo de Operação"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8655
   Tag             =   "Associações Itens Caixa"
   Begin MSComctlLib.Toolbar tlbAssociacao 
      Height          =   330
      Left            =   2280
      TabIndex        =   10
      Top             =   5310
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   582
      ButtonWidth     =   3625
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Associar Item de Caixa"
            Key             =   "associar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Desfazer Associação   "
            Key             =   "desfazer"
            ImageKey        =   "Excluir"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Itens de Caixa "
      Height          =   4155
      Left            =   60
      TabIndex        =   5
      Top             =   1050
      Width           =   8535
      Begin MSComctlLib.TreeView treItemCaixa 
         Height          =   3825
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6747
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlIcons"
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   8535
      Begin VB.OptionButton optTipoCaixa 
         Caption         =   "Caixa Sub Reserva"
         Height          =   195
         Index           =   1
         Left            =   5430
         TabIndex        =   8
         Tag             =   "SubReserva"
         Top             =   480
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optTipoCaixa 
         Caption         =   "Caixa Futuro"
         Height          =   195
         Index           =   2
         Left            =   7260
         TabIndex        =   7
         Tag             =   "CaixaFuturo"
         Top             =   480
         Width           =   1185
      End
      Begin VB.ComboBox cboTipoOperacao 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   420
         Width           =   5160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operação"
         Height          =   195
         Left            =   165
         TabIndex        =   9
         Top             =   210
         Width           =   1290
      End
   End
   Begin VB.Frame fraMoldura 
      Caption         =   "Associações "
      Height          =   1995
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   5610
      Width           =   8535
      Begin MSComctlLib.ListView lstItemCaixaTipoOperacao 
         Height          =   1635
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2884
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo Contraparte"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo Movimento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Caixa"
            Object.Width           =   9243
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   3480
      Top             =   7470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":195E
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":1A58
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":1B52
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaTipoOperacao.frx":2094
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   6690
      TabIndex        =   1
      Top             =   7650
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmItemCaixaTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a administração da associação entre itens de caixa e
' tipos de operação.

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLerTodos                         As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmItemCaixaTipoOperacao"
Private blnSemAssociacao                    As Boolean

Private Const POS_TAG_TP_OPER               As Integer = 0
Private Const POS_TAG_TP_CAIX               As Integer = 1
Private Const POS_TAG_TP_CNPT               As Integer = 2
Private Const POS_TAG_IN_ENTR_SAID          As Integer = 3
Private Const POS_TAG_CO_ITEM_CAIX          As Integer = 4
Private Const POS_TAG_CO_ITEM_CAIX_OLD      As Integer = 5
Private Const POS_TAG_DH_ULTI_ATLZ          As Integer = 6
Private Const POS_TAG_OPER                  As Integer = 7 'I - INCLUSÃO ; A- ALTERAÇÃO ; E - EXCLUSAO

' Aplica a associação de um tipo de operação, a um ou mais itens de caixa.

Private Sub flAssociarItemCaixa()

Dim objNode                                 As MSComctlLib.Node
Dim objListItem                             As MSComctlLib.ListItem
Dim arrTag()                                As String 'Conten CO_ITEM_CAIXA;DH_ULTI_ATLZ

On Error GoTo ErrorHandler

    Set objNode = treItemCaixa.SelectedItem
        
    If objNode Is Nothing Then
        Exit Sub
    End If
    
    If objNode.Image <> "Leaf" Then
        frmMural.Display = "Somente itens elementares podem ser associados a tipos de operação"
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If
    
    For Each objListItem In lstItemCaixaTipoOperacao.ListItems
        If objListItem.Selected Then
            objListItem.SubItems(2) = objNode.Text
            
            arrTag = Split(objListItem.Tag, ";")
                        
            'If arrTag(POS_TAG_OPER) = "A" Then
            If arrTag(POS_TAG_CO_ITEM_CAIX) = "" Then
                objListItem.Tag = fgObterCodigoCombo(cboTipoOperacao.Text) & ";" & _
                                  IIf(optTipoCaixa(1), enumTipoCaixa.CaixaSubReserva, enumTipoCaixa.CaixaFuturo) & ";" & _
                                  arrTag(POS_TAG_TP_CNPT) & ";" & _
                                  arrTag(POS_TAG_IN_ENTR_SAID) & ";" & _
                                  Mid$(objNode.Key, 4, 1) & Mid$(objNode.Key, 6) & ";" & _
                                  arrTag(POS_TAG_CO_ITEM_CAIX) & ";" & _
                                  arrTag(POS_TAG_DH_ULTI_ATLZ) & ";" & _
                                  "I"
            Else
                objListItem.Tag = arrTag(POS_TAG_TP_OPER) & ";" & _
                                  arrTag(POS_TAG_TP_CAIX) & ";" & _
                                  arrTag(POS_TAG_TP_CNPT) & ";" & _
                                  arrTag(POS_TAG_IN_ENTR_SAID) & ";" & _
                                  Mid$(objNode.Key, 4, 1) & Mid$(objNode.Key, 6) & ";" & _
                                  arrTag(POS_TAG_CO_ITEM_CAIX) & ";" & _
                                  arrTag(POS_TAG_DH_ULTI_ATLZ) & ";" & _
                                  "A"
                       
            End If
            
        End If
    Next
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - flAssociarItemCaixa"

End Sub

' Carrega todas as combinações possíveis de associações com Itens de Caixa.

Private Sub flCarregarCombinacoesAssociacoes()

Dim objListItem                             As MSComctlLib.ListItem
    
On Error GoTo ErrorHandler

    With lstItemCaixaTipoOperacao
        .ListItems.Clear
        
        Set objListItem = .ListItems.Add(, "K" & enumTipoContraparte.Externo & "K" & enumTipoEntradaSaida.Entrada)
        objListItem.Text = "Externo"
        objListItem.SubItems(1) = "Entrada"
        objListItem.Tag = vbNullString & ";" & _
                          vbNullString & ";" & _
                          enumTipoContraparte.Externo & ";" & _
                          enumTipoEntradaSaida.Entrada & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          "C"
        
        
        Set objListItem = .ListItems.Add(, "K" & enumTipoContraparte.Externo & "K" & enumTipoEntradaSaida.Saida)
        objListItem.Text = "Externo"
        objListItem.SubItems(1) = "Saída"
        objListItem.Tag = vbNullString & ";" & _
                          vbNullString & ";" & _
                          enumTipoContraparte.Externo & ";" & _
                          enumTipoEntradaSaida.Saida & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          "C"
        
    
        Set objListItem = .ListItems.Add(, "K" & enumTipoContraparte.Interno & "K" & enumTipoEntradaSaida.Entrada)
        objListItem.Text = "Interno"
        objListItem.SubItems(1) = "Entrada"
        objListItem.Tag = vbNullString & ";" & _
                          vbNullString & ";" & _
                          enumTipoContraparte.Interno & ";" & _
                          enumTipoEntradaSaida.Entrada & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          "C"
        
        
        Set objListItem = .ListItems.Add(, "K" & enumTipoContraparte.Interno & "K" & enumTipoEntradaSaida.Saida)
        objListItem.Text = "Interno"
        objListItem.SubItems(1) = "Saída"
        objListItem.Tag = vbNullString & ";" & _
                          vbNullString & ";" & _
                          enumTipoContraparte.Interno & ";" & _
                          enumTipoEntradaSaida.Saida & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          "C"
        
    
        Set objListItem = .ListItems.Add(, "K" & enumTipoContraparte.Cliente1 & "K" & enumTipoEntradaSaida.Entrada)
        objListItem.Text = "Cliente1"
        objListItem.SubItems(1) = "Entrada"
        objListItem.Tag = vbNullString & ";" & _
                          vbNullString & ";" & _
                          enumTipoContraparte.Cliente1 & ";" & _
                          enumTipoEntradaSaida.Entrada & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          "C"
        
        Set objListItem = .ListItems.Add(, "K" & enumTipoContraparte.Cliente1 & "K" & enumTipoEntradaSaida.Saida)
        objListItem.Text = "Cliente1"
        objListItem.SubItems(1) = "Saída"
        objListItem.Tag = vbNullString & ";" & _
                          vbNullString & ";" & _
                          enumTipoContraparte.Cliente1 & ";" & _
                          enumTipoEntradaSaida.Saida & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          vbNullString & ";" & _
                          "C"
        
    End With

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - flCarregarCombinacoesAssociacoes"

End Sub

' Carrega a lista de itens de caixa, a partir do tipo de operação selecionado.

Private Sub flCarregarItensAssociados(ByVal plngTipoOperacao As Long)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim xmlDomNode          As MSXML2.IXMLDOMNode
Dim strxmlLerTodos      As String
Dim objListItem         As MSComctlLib.ListItem
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Call flCarregarCombinacoesAssociacoes

    Set xmlDomNode = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixaTipoOperacao")
    xmlDomNode.selectSingleNode("@Operacao").Text = "LerTodos"
    
    If optTipoCaixa(enumTipoCaixa.CaixaFuturo).Value Then
        xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaFuturo
    ElseIf optTipoCaixa(enumTipoCaixa.CaixaSubReserva).Value Then
        xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaSubReserva
    End If
    
    xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text = ""
    xmlDomNode.selectSingleNode("TP_OPER").Text = plngTipoOperacao

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    strxmlLerTodos = objMIU.Executar(xmlDomNode.xml, _
                                     vntCodErro, _
                                     vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If strxmlLerTodos <> "" Then
       Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
       If Not xmlLerTodos.loadXML(strxmlLerTodos) Then
          Call fgErroLoadXML(xmlLerTodos, App.EXEName, Me.Name & "", "flCarregarItensAssociados")
       End If
       blnSemAssociacao = False
    Else
       Set objMIU = Nothing
       blnSemAssociacao = True
       Exit Sub
    End If

    Set objMIU = Nothing
    
    With lstItemCaixaTipoOperacao
        For Each xmlDomNode In xmlLerTodos.documentElement.selectNodes("Grupo_ItemCaixaTipoOperacao")
            Set objListItem = lstItemCaixaTipoOperacao.ListItems("K" & xmlDomNode.selectSingleNode("TP_CNPT").Text & _
                                                                 "K" & xmlDomNode.selectSingleNode("IN_ENTR_SAID").Text)
                                                                 
            If Not objListItem Is Nothing Then
                objListItem.SubItems(2) = xmlDomNode.selectSingleNode("DE_ITEM_CAIX").Text
                objListItem.Tag = xmlDomNode.selectSingleNode("TP_OPER").Text & ";" & _
                                  xmlDomNode.selectSingleNode("TP_CAIX").Text & ";" & _
                                  xmlDomNode.selectSingleNode("TP_CNPT").Text & ";" & _
                                  xmlDomNode.selectSingleNode("IN_ENTR_SAID").Text & ";" & _
                                  xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & ";" & _
                                  xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & ";" & _
                                  xmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & ";" & _
                                  "C"
            End If
        Next
    End With

    Set objMIU = Nothing

    Exit Sub

ErrorHandler:
    
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - flCarregarItensAssociados"

End Sub

' Limpa associações para os itens selecionados do ListView.

Private Sub flDesfazerAssociacao()

Dim objListItem                             As MSComctlLib.ListItem
Dim arrTag()                                As String 'Conten CO_ITEM_CAIXA;DH_ULTI_ATLZ

On Error GoTo ErrorHandler

    For Each objListItem In lstItemCaixaTipoOperacao.ListItems
        If objListItem.Selected Then
            
            If Trim(objListItem.SubItems(2)) <> vbNullString Then
            
                objListItem.SubItems(2) = vbNullString
                
                arrTag = Split(objListItem.Tag, ";")
                
                objListItem.Tag = arrTag(POS_TAG_TP_OPER) & ";" & _
                                  arrTag(POS_TAG_TP_CAIX) & ";" & _
                                  arrTag(POS_TAG_TP_CNPT) & ";" & _
                                  arrTag(POS_TAG_IN_ENTR_SAID) & ";" & _
                                  arrTag(POS_TAG_CO_ITEM_CAIX) & ";" & _
                                  arrTag(POS_TAG_CO_ITEM_CAIX) & ";" & _
                                  arrTag(POS_TAG_DH_ULTI_ATLZ) & ";" & _
                                  "E"
            End If
        End If
    Next
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - flDesfazerAssociacao"

End Sub

' Carrega configurações iniciais do formulário.

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim strMapaNavegacao    As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.SBR, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmItemCaixa", "flInicializar")
    End If

    Call fgCarregarCombos(Me.cboTipoOperacao, xmlMapaNavegacao, "TipoOperacao", "TP_OPER", "NO_TIPO_OPER")
    
    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0

End Sub

' Aciona a atualização da tabela de associações item caixa x tipo de operação.

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim xmlDomNode          As MSXML2.IXMLDOMNode
Dim xmlSalvar           As MSXML2.DOMDocument40
Dim objListItem         As MSComctlLib.ListItem
Dim lngItem             As Long
Dim blnIncluir          As Boolean
Dim arrChaves()         As String
Dim arrTag()            As String 'Conten CO_ITEM_CAIXA;DH_ULTI_ATLZ
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlSalvar = CreateObject("MSXML2.DOMDocument.4.0")
    fgAppendNode xmlSalvar, "", "Repeat_ItemCaixaTipoOperacao", ""

    blnIncluir = False
    For Each objListItem In lstItemCaixaTipoOperacao.ListItems
        If objListItem.Tag <> vbNullString Then
            
            arrTag = Split(objListItem.Tag, ";")
                        
            Set xmlDomNode = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixaTipoOperacao")
            
            If arrTag(POS_TAG_OPER) <> "C" Then
                
                blnIncluir = True
                
                Select Case arrTag(POS_TAG_OPER)
                    Case "A"
                        xmlDomNode.selectSingleNode("@Operacao").Text = "Alterar"
                    Case "I"
                        xmlDomNode.selectSingleNode("@Operacao").Text = "Incluir"
                    Case "E"
                        xmlDomNode.selectSingleNode("@Operacao").Text = "Excluir"
                End Select
                
                arrTag = Split(objListItem.Tag, ";")
                xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text = arrTag(POS_TAG_CO_ITEM_CAIX)
                xmlDomNode.selectSingleNode("TP_OPER").Text = arrTag(POS_TAG_TP_OPER)
                xmlDomNode.selectSingleNode("TP_CNPT").Text = arrTag(POS_TAG_TP_CNPT)
                xmlDomNode.selectSingleNode("IN_ENTR_SAID").Text = arrTag(POS_TAG_IN_ENTR_SAID)
                xmlDomNode.selectSingleNode("CO_ITEM_CAIX_OLD").Text = arrTag(POS_TAG_CO_ITEM_CAIX_OLD)
                xmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text = arrTag(POS_TAG_DH_ULTI_ATLZ)
    
                fgAppendXML xmlSalvar, "Repeat_ItemCaixaTipoOperacao", xmlDomNode.xml, "Repeat_ItemCaixaTipoOperacao"
            
            End If
        End If
    Next

    If Not blnIncluir Then
        MsgBox "Nenhum Item de Caixa foi associado", vbCritical, Me.Caption
        Set xmlSalvar = Nothing
        Exit Sub
    End If
    
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    Call objMIU.Executar(xmlSalvar.xml, _
                         vntCodErro, _
                         vntMensagemErro)
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    flCarregarItensAssociados CLng("0" & fgObterCodigoCombo(cboTipoOperacao.Text))
    
    fgCursor
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - flSalvar")
    
    flCarregarItensAssociados CLng("0" & fgObterCodigoCombo(cboTipoOperacao.Text))
    
End Sub

Private Sub cboTipoOperacao_Click()

On Error GoTo ErrorHandler

    If cboTipoOperacao.ListIndex = -1 Then Exit Sub
    
    DoEvents
    fgCursor True
    
    treItemCaixa.Enabled = True
    flCarregarItensAssociados CLng("0" & fgObterCodigoCombo(cboTipoOperacao.Text))

    fgCursor

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_Load"

End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler
    
    Me.Icon = mdiSBR.Icon
    
    fgCursor True
    fgCenterMe Me
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    treItemCaixa.Enabled = False
    
    Call flInicializar
    Call flCarregarCombinacoesAssociacoes
    Call optTipoCaixa_Click(1)

    Me.Show
    DoEvents
    
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - Form_Load"

End Sub

Private Sub lstItemCaixaTipoOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstItemCaixaTipoOperacao, ColumnHeader.Index)

    Exit Sub

ErrorHandler:
     mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - lstItemCaixaTipoOperacao_ColumnClick"

End Sub

Private Sub optTipoCaixa_Click(Index As Integer)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim xmlDomNode          As MSXML2.IXMLDOMNode
Dim strLerTodos         As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant
    
On Error GoTo ErrorHandler

    Call flCarregarCombinacoesAssociacoes

    fgCursor True
    Set xmlDomNode = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa")
    xmlDomNode.selectSingleNode("@Operacao").Text = "LerTodos"
    Select Case Index
        Case enumTipoCaixa.CaixaFuturo
             xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaFuturo
        Case enumTipoCaixa.CaixaSubReserva
             xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaSubReserva
    End Select

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    strLerTodos = objMIU.Executar(xmlDomNode.xml, _
                                  vntCodErro, _
                                  vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strLerTodos <> "" Then
       Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
       If Not xmlLerTodos.loadXML(strLerTodos) Then
          Call fgErroLoadXML(xmlLerTodos, App.EXEName, Me.Name & "", "optTipoCaixa_Click")
       End If
    Else
       Set objMIU = Nothing
       fgCursor
       Exit Sub
    End If

    Set objMIU = Nothing

    fgCarregarTreItemCaixa treItemCaixa, xmlLerTodos, Me
    
    If cboTipoOperacao.ListIndex > -1 Then
        flCarregarItensAssociados CLng("0" & fgObterCodigoCombo(cboTipoOperacao.Text))
    End If
    
    Set xmlDomNode = Nothing
    
    fgCursor
        
    Exit Sub
    
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - optTipoCaixa_Click"
    
End Sub

Private Sub tlbAssociacao_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "associar"
            Call flAssociarItemCaixa
        Case "desfazer"
            Call flDesfazerAssociacao
    End Select
    
    Exit Sub

ErrorHandler:
   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbAssociacao_ButtonClick"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo ErrorHandler

    fgCursor True
    
    Select Case Button.Key
        Case "Salvar"
            Call flSalvar
        Case "Sair"
            Unload Me
    End Select
    
    fgCursor

    Exit Sub

ErrorHandler:
   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbCadastro_ButtonClick"

End Sub

Private Sub treItemCaixa_Collapse(ByVal Node As MSComctlLib.Node)
    
On Error GoTo ErrorHandler

    If Node.Expanded = True And Node.children <> 0 Then
        Node.Image = "ItemGrupoAberto"
    ElseIf Node.Expanded = False And Node.children <> 0 Then
        Node.Image = "ItemGrupoFechado"
    ElseIf Node.children = 0 Then
        Node.Image = "ItemElementar"
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - treItemCaixa_NodeClick"

End Sub

Private Sub treItemCaixa_Expand(ByVal Node As MSComctlLib.Node)
    
On Error GoTo ErrorHandler

    If Node.Expanded = True And Node.children <> 0 Then
        Node.Image = "ItemGrupoAberto"
    ElseIf Node.Expanded = False And Node.children <> 0 Then
        Node.Image = "ItemGrupoFechado"
    ElseIf Node.children = 0 Then
        Node.Image = "ItemElementar"
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - treItemCaixa_NodeClick"

End Sub

Private Sub treItemCaixa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo ErrorHandler

    If cboTipoOperacao.ListIndex > -1 Then
        treItemCaixa.ToolTipText = "Selecione um Item de Caixa Elementar"
    Else
        treItemCaixa.ToolTipText = "Selecione um Tipo Operação"
    End If

    Exit Sub

ErrorHandler:
   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - treItemCaixa_MouseMove"
        
End Sub

