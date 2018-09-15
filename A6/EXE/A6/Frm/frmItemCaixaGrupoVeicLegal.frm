VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemCaixaGrupoVeicLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Itens de Caixa x Grupo de Veículo Legal"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6540
   Tag             =   "Relação Itens Caixa"
   Begin VB.OptionButton optTipoCaixa 
      Caption         =   "Caixa Futuro"
      Height          =   255
      Index           =   2
      Left            =   1620
      TabIndex        =   6
      Tag             =   "CaixaFuturo"
      Top             =   750
      Width           =   1245
   End
   Begin VB.OptionButton optTipoCaixa 
      Caption         =   "Sub Reserva"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Tag             =   "SubReserva"
      Top             =   750
      Width           =   1335
   End
   Begin VB.Frame fraMoldura 
      Height          =   4575
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   6135
      Begin MSComctlLib.ListView lstItemCaixa 
         Height          =   4200
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   7408
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   9702
         EndProperty
      End
   End
   Begin VB.ComboBox cboGrupoVeicLegal 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4035
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   4380
      TabIndex        =   2
      Top             =   5430
      Width           =   2175
      _ExtentX        =   3836
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
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   180
      Top             =   5400
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
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":195E
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":1A58
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":1B52
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixaGrupoVeicLegal.frx":2094
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGrupoVeicLegal 
      AutoSize        =   -1  'True
      Caption         =   "Grupo de Veículos Legais"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   300
      Width           =   1845
   End
End
Attribute VB_Name = "frmItemCaixaGrupoVeicLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a administração da associação entre itens de caixa,
' e grupos de veículos legais.

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlItemCaixaGrupoVeicLegal          As MSXML2.DOMDocument40

Private strOperacao                         As String
Private Const strFuncionalidade             As String = "frmItemCaixaGrupoVeicLegal"

' Aciona a atualização da tabela de item caixa x grupo de veículo legal.

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A6MIU.clsMIU
#End If

Dim xmlSalvar               As MSXML2.DOMDocument40
Dim xmlNode                 As MSXML2.IXMLDOMNode
Dim lstListItem             As MSComctlLib.ListItem
Dim lngItem                 As Long
Dim blnExisteOperacao       As Boolean
Dim lngGrupoVeiculoLegal    As Long
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If lstItemCaixa.ListItems.Count = 0 Then
       Exit Sub
    End If

    Set xmlSalvar = CreateObject("MSXML2.DOMDocument.4.0")

    fgAppendNode xmlSalvar, "", "Repeat_ItemCaixaGrupoVeicLegal", ""
    
    lngGrupoVeiculoLegal = fgObterCodigoCombo(cboGrupoVeicLegal.Text)

    With lstItemCaixa
    
         For Each lstListItem In .ListItems
         
            If lstListItem.Tag = enumTipoOperacao.Incluir Then
               blnExisteOperacao = True
               Set xmlNode = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixaGrupoVeicLegal")
                
               xmlNode.selectSingleNode("@Operacao").Text = "Incluir"
               xmlNode.selectSingleNode("CO_ITEM_CAIX").Text = Mid(lstListItem.Key, 2)
               xmlNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text = lngGrupoVeiculoLegal
            
               fgAppendXML xmlSalvar, "Repeat_ItemCaixaGrupoVeicLegal", xmlNode.xml
               
            ElseIf lstListItem.Tag = enumTipoOperacao.Excluir Then
                blnExisteOperacao = True
                For Each xmlNode In xmlItemCaixaGrupoVeicLegal.documentElement.selectNodes("Grupo_ItemCaixaGrupoVeicLegal")
                    
                    If xmlNode.selectSingleNode("CO_ITEM_CAIX").Text = Mid(lstListItem.Key, 2) And _
                       xmlNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text = lngGrupoVeiculoLegal Then
                       
                       xmlNode.selectSingleNode("@Operacao").Text = "Excluir"
                       fgAppendXML xmlSalvar, "Repeat_ItemCaixaGrupoVeicLegal", xmlNode.xml
                       Exit For
                    End If
                
                Next
            
            End If
         
         Next lstListItem
    
    End With

    If blnExisteOperacao Then
        Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
        Call objMIU.Executar(xmlSalvar.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        flLimparOperacao
        flInicializar
        fgCursor
        MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
        If optTipoCaixa(enumTipoCaixa.CaixaSubReserva).Value = True Then
           Call optTipoCaixa_Click(optTipoCaixa(enumTipoCaixa.CaixaSubReserva).Index)
        Else
           Call optTipoCaixa_Click(optTipoCaixa(enumTipoCaixa.CaixaFuturo).Index)
        End If
    
    Else
        fgCursor
        MsgBox "Não existe Operação a ser executada.", vbInformation, Me.Caption
    End If
    
    Set objMIU = Nothing
    Set xmlSalvar = Nothing
    Set xmlNode = Nothing
    
    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlSalvar = Nothing
    Set xmlNode = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

' Localiza um item no listview, a partir de sua propriedade Key.

Private Sub flLocalizaItemListView(ByVal pvChave As Variant, _
                                   ByVal plstListView As MSComctlLib.ListView)

Dim lstListItem As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    With plstListView
         For Each lstListItem In .ListItems
             If pvChave = Mid(lstListItem.Key, 2) Then
                lstListItem.Checked = True
                lstListItem.Tag = enumTipoOperacao.None
             End If
         Next lstListItem
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLocalizaItemListView", 0

End Sub

' Carrega Listview com itens de caixa.

Private Sub flCarregarlstItemCaixa(ByVal penumTipoCaixa As enumTipoCaixa)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim xmlDomItem          As MSXML2.IXMLDOMNode
Dim lstListItem         As MSComctlLib.ListItem
Dim xmlItemCaixa        As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    lstItemCaixa.ListItems.Clear
    
    'Setar Filtros - Tipo de Caixa
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa/TP_CAIX").Text = penumTipoCaixa
    
    'Setar Filtros - Somente Nivel 01
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa/CO_NIVEL_01").Text = "S"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa/@Operacao").Text = "LerTodos"
    
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    Set xmlItemCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlItemCaixa.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectNodes("//Grupo_Propriedades/Grupo_ItemCaixa").Item(0).xml, vntCodErro, vntMensagemErro)) Then
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        GoTo Fim
    End If

    Set objMIU = Nothing
        
    
    fgLockWindow lstItemCaixa.hwnd

    With lstItemCaixa
        
        For Each xmlDomItem In xmlItemCaixa.documentElement.selectNodes("/Repeat_ItemCaixa/Grupo_ItemCaixa")
            If xmlDomItem.selectSingleNode("DE_ITEM_CAIX").Text <> gstrItemGenerico Then
                Set lstListItem = .ListItems.Add(, "K" & xmlDomItem.selectSingleNode("CO_ITEM_CAIX").Text, xmlDomItem.selectSingleNode("DE_ITEM_CAIX").Text)
                lstListItem.Tag = enumTipoOperacao.None
            End If
        Next
    End With
    
    Set xmlItemCaixa = Nothing
    
    If lstItemCaixa.ListItems.Count > 0 Then
       lstItemCaixa.ListItems(1).Selected = False
    End If

Fim:
    fgLockWindow 0
    Exit Sub

ErrorHandler:
    fgLockWindow 0
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarlstItemCaixa", 0

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
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.SBR, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmItemCaixaGrupoVeicLegal", "flInicializar")
    End If
    
    Set objMIU = Nothing

    Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0
    
End Sub

Private Sub cboGrupoVeicLegal_Click()

On Error GoTo ErrorHandler

    lstItemCaixa.ListItems.Clear
    optTipoCaixa.Item(1).Value = False
    optTipoCaixa.Item(2).Value = False

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboGrupoVeicLegal_Click"
        
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    Me.Icon = mdiSBR.Icon
    
    fgCursor True
    fgCenterMe Me
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Me.Show
    DoEvents
    
    strOperacao = ""

    flInicializar
    
    Call fgCarregarCombos(Me.cboGrupoVeicLegal, xmlMapaNavegacao, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA")
    
    fgCursor

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixaGrupoVeicLegal - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmItemCaixaGrupoVeicLegal = Nothing

End Sub

Private Sub lstItemCaixa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
On Error GoTo ErrorHandler

    If Item.Checked Then
        
       If Item.Tag = enumTipoOperacao.None Then
       
          Item.Tag = enumTipoOperacao.Incluir
          
       ElseIf Item.Tag = enumTipoOperacao.Excluir Then
          
          Item.Tag = enumTipoOperacao.None
       
       End If
       
    Else    'Unchecked
    
       If Item.Tag = enumTipoOperacao.Incluir Then
       
          Item.Tag = enumTipoOperacao.None
          
       ElseIf Item.Tag = enumTipoOperacao.None Then
          
          Item.Tag = enumTipoOperacao.Excluir
       
       End If
    
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixaGrupoVeicLegal - lstItemCaixa_ItemCheck()"

End Sub

Private Sub optTipoCaixa_Click(Index As Integer)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim xmlNode             As MSXML2.IXMLDOMNode
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    If cboGrupoVeicLegal.ListIndex < 0 Then
       Call flLimparOperacao
       Exit Sub
    End If
    
    fgCursor True
    
    fgLockWindow lstItemCaixa.hwnd
    
    If Index = enumTipoCaixa.CaixaSubReserva Then
       flCarregarlstItemCaixa enumTipoCaixa.CaixaSubReserva
    Else
       flCarregarlstItemCaixa enumTipoCaixa.CaixaFuturo
    End If
    
    Set xmlItemCaixaGrupoVeicLegal = CreateObject("MSXML2.DOMDocument.4.0")
    xmlItemCaixaGrupoVeicLegal.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//frmItemCaixaGrupoVeicLegal/Grupo_Propriedades/Grupo_ItemCaixaGrupoVeicLegal").xml
    xmlItemCaixaGrupoVeicLegal.documentElement.selectSingleNode("@Operacao").Text = "LerTodos"
    xmlItemCaixaGrupoVeicLegal.documentElement.selectSingleNode("CO_ITEM_CAIX").Text = vbNullString
    xmlItemCaixaGrupoVeicLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text = IIf(fgObterCodigoCombo(cboGrupoVeicLegal.Text) = "", 0, fgObterCodigoCombo(cboGrupoVeicLegal.Text))

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    xmlItemCaixaGrupoVeicLegal.loadXML objMIU.Executar(xmlItemCaixaGrupoVeicLegal.selectSingleNode("//Grupo_ItemCaixaGrupoVeicLegal").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If xmlItemCaixaGrupoVeicLegal.xml <> "" Then
        For Each xmlNode In xmlItemCaixaGrupoVeicLegal.documentElement.childNodes
        
            flLocalizaItemListView xmlNode.selectSingleNode("CO_ITEM_CAIX").Text, lstItemCaixa
        
        Next xmlNode
    End If
    
    fgLockWindow 0
    
    fgCursor

    Exit Sub

ErrorHandler:
    Set xmlNode = Nothing
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixaGrupoVeicLegal - optTipoCaixa_Click()"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True
    
    Select Case Button.Key
    Case "Salvar"
        flSalvar
            
    Case "Sair"
        Unload Me
    
    End Select
    
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "tlbCadastro_ButtonClick"
    
    If optTipoCaixa(enumTipoCaixa.CaixaSubReserva).Value = True Then
       Call optTipoCaixa_Click(optTipoCaixa(enumTipoCaixa.CaixaSubReserva).Index)
    Else
       Call optTipoCaixa_Click(optTipoCaixa(enumTipoCaixa.CaixaFuturo).Index)
    End If
    
End Sub

' Limpa Tags dos itens do listview, inidicando que não há nenhuma operação a ser executada.

Private Sub flLimparOperacao()

Dim lstListItem                             As MSComctlLib.ListItem
    
On Error GoTo ErrorHandler

    For Each lstListItem In lstItemCaixa.ListItems
        lstListItem.Tag = enumTipoOperacao.None
    Next

Exit Sub

ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparOperacao", 0

End Sub

