VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultaMensagemSisbacen 
   Caption         =   "Consulta Operação Sisbacen"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   12660
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFiltro 
      Height          =   945
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      Begin VB.ComboBox cboMensagem 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Width           =   6645
      End
      Begin VB.Label Label3 
         Caption         =   "Mensagem"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Width           =   1140
      End
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Index           =   0
      Left            =   240
      Top             =   6540
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
            Picture         =   "frmConsultaMensagemSisbacen.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   2685
      Left            =   120
      TabIndex        =   3
      Top             =   1005
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   7455
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   635
      ButtonWidth     =   2487
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar Filtro"
            Key             =   "AplicarFiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
            Key             =   "showfiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Index           =   1
      Left            =   120
      Top             =   0
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
            Picture         =   "frmConsultaMensagemSisbacen.frx":19F2
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":1B04
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":1C16
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":1F68
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":22BA
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":260C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaMensagemSisbacen.frx":295E
            Key             =   "sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaMensagemSisbacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

'Constante para o Mapa de Navegacao
Private Const strFuncionalidade             As String = "frmConsultaMensagemSisbacen"
Private Const intGrupoCAM                   As Integer = 21
Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1
Private strFiltroXML                        As String
Private blnPrimeiraConsulta                 As Boolean
Private blnUtilizaFiltro                    As Boolean
Private blnFiltroAplicado                   As Boolean
Private blnOrigemBotaoRefresh               As Boolean

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    
    DoEvents

    Call flInicializar
    blnPrimeiraConsulta = True
    blnFiltroAplicado = False
    blnUtilizaFiltro = (tlbButtons.Buttons("AplicarFiltro").value = tbrPressed)
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaMensagemSisbacen
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    
    blnPrimeiraConsulta = False
    
    fgCursor
                
Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me

        .lvwMensagem.Top = .fraFiltro.Height + 100
        .lvwMensagem.Left = .fraFiltro.Left
        .lvwMensagem.Height = .Height - .lvwMensagem.Top - 950
        .lvwMensagem.Width = .Width - 400

    End With

End Sub

Private Sub flLimparLista()

    Me.lvwMensagem.ListItems.Clear
    
End Sub

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
    
    vntCodErro = 0
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    Call flCarregaComboMensagem
    
    Set objMIU = Nothing
    Set objMensagem = Nothing
 
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Set objMensagem = Nothing
    Set xmlMapaNavegacao = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

Private Sub flCarregaComboMensagem()

On Error GoTo ErrorHandler
    
    With cboMensagem
        .Clear
        .AddItem "CAM0042 - IF consulta contratos em ser", 0
        .AddItem "CAM0043 - IF consulta eventos de um dia", 1
        .AddItem "CAM0044 - IF consulta detalhamento de contrato interbancário", 2
        .AddItem "CAM0045 - IF consulta eventos de um contrato do mercado primário", 3
        .AddItem "CAM0046 - Corretora consulta eventos de um contrato intermediário no mercado primário", 4
        .AddItem "CAM0047 - IF consulta histórico de incorporações", 5
        .AddItem "CAM0048 - IF consulta contratos da incorporação", 6
        .AddItem "CAM0049 - IF consulta cadeia de incorporações de um contrato", 7
        .AddItem "CAM0050 - IF consulta posição de câmbio por moeda", 8
        .AddItem "CAM0052 - IF consulta instruções de pagamento", 9
        .ListIndex = -1
    End With
    
Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaComboMensagem", 0

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

Dim strSelecaoVisual                        As String
Dim strSelecaoFiltro                        As String
Dim objDomDocument                          As MSXML2.DOMDocument
Dim strErroCamposObrigatorios               As String
    
On Error GoTo ErrorHandler

    fgCursor True
    blnUtilizaFiltro = True
    blnFiltroAplicado = True

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If blnPrimeiraConsulta Then
            blnPrimeiraConsulta = False
            
            If blnOrigemBotaoRefresh Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
                Exit Sub
            Else
                Call tlbButtons_ButtonClick(tlbButtons.Buttons("showfiltro"))
                Exit Sub
            End If
        End If
        
        'Pressiona o botão << Aplicar Filtro >> apenas se o filtro for selecionado diretamente
        If Not blnOrigemBotaoRefresh Then
            blnUtilizaFiltro = True
            tlbButtons.Buttons("AplicarFiltro").value = tbrPressed
        End If
        
        'Valida os dados Obrigatórios
        Set objDomDocument = CreateObject("MSXML2.DOMDocument")
        objDomDocument.loadXML strFiltroXML
        If objDomDocument.documentElement.selectSingleNode("//DataIni") Is Nothing Then
            strErroCamposObrigatorios = strErroCamposObrigatorios & "Obrigatório a Seleção da Data " & vbNewLine
        End If
        If objDomDocument.documentElement.selectSingleNode("//CodigoMensagem") Is Nothing Then
            strErroCamposObrigatorios = strErroCamposObrigatorios & "Obrigatório a Seleção do Código da Mensagem CAM " & vbNewLine
        End If
        
        'Se nenhum dos campos obrigatórios for selecionado sair do filtro
        If strErroCamposObrigatorios <> "" Then
            frmMural.Caption = Me.Caption
            frmMural.Display = strErroCamposObrigatorios
            frmMural.Show vbModal
            blnUtilizaFiltro = False
            lvwMensagem.ListItems.Clear
            cboMensagem.ListIndex = -1
            cboMensagem.Enabled = False
            Set objDomDocument = Nothing
            Exit Sub
        Else
            cboMensagem.Enabled = True
        End If
        
        Call flCarregaList
        
        If (blnUtilizaFiltro = True) Then
            fgSearchItemCombo cboMensagem, , objDomDocument.documentElement.selectSingleNode("//CodigoMensagem").Text
        End If
        
        Set objDomDocument = Nothing
        
    End If
    
    blnUtilizaFiltro = False
    fgCursor
    
Exit Sub
ErrorHandler:
    fgCursor
    Set objDomDocument = Nothing
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub cboMensagem_Click()
On Error GoTo ErrorHandler
    
    fgCursor True
    
    If blnUtilizaFiltro = False Then
        Call flCarregaList
        blnFiltroAplicado = False
    End If
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboMensagem_Click", Me.Caption

End Sub

Private Sub flCarregaList()

#If EnableSoap = 1 Then
    Dim objMensagem        As MSSOAPLib30.SoapClient30
    Dim objXMLMensagem     As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem        As A8MIU.clsMensagem
    Dim objXMLMensagem     As A8MIU.clsMensagem
#End If

    Dim strFiltro               As String
    Dim vntCodErro              As Variant
    Dim vntMensagemErro         As Variant
    Dim objDomDocument          As MSXML2.DOMDocument
    Dim objDomAux               As MSXML2.DOMDocument
    Dim objDomDocumentFiltroXml As MSXML2.DOMDocument
    Dim objNode                 As MSXML2.IXMLDOMNode
    Dim objNodeAux              As MSXML2.IXMLDOMNode
    Dim objListItem             As ListItem
    Dim lngList                 As Long
    Dim strNumeroControle       As String
    Dim strCodigoMensagem       As String

On Error GoTo ErrorHandler

    strFiltro = flMontarXMLFiltroPesquisa
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set objXMLMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    Set objDomDocument = New MSXML2.DOMDocument
    Set objDomAux = New MSXML2.DOMDocument
    Set objDomDocumentFiltroXml = CreateObject("MSXML2.DOMDocument")
        
    flFormataListView strFiltro
    
    objDomDocument.loadXML objMensagem.ObterDetalheMensagem(strFiltro, vntCodErro, vntMensagemErro)
    
    For Each objNode In objDomDocument.selectNodes("Repeat_DetalheMensagem/*")
        
        strNumeroControle = objNode.selectSingleNode("NU_CTRL_IF").Text
        Set objListItem = lvwMensagem.ListItems.Add(, strNumeroControle, "")
        lngList = lvwMensagem.ListItems.Count
        lvwMensagem.ListItems.Item(lngList).Tag = objNode.selectSingleNode("CO_TEXT_XML").Text
        
        objDomAux.loadXML objMensagem.ObterXMLMensagem(objNode.selectSingleNode("CO_TEXT_XML").Text, vntCodErro, vntMensagemErro)
        
        objDomDocumentFiltroXml.loadXML strFiltro
        If blnUtilizaFiltro = False Then
            strCodigoMensagem = Left(cboMensagem.Text, 7)
        Else
            strCodigoMensagem = Left(objDomDocumentFiltroXml.selectSingleNode("//CodigoMensagem").Text, 7)
        End If
        
        For Each objNodeAux In objDomAux.selectNodes("SISMSG/" & strCodigoMensagem & "R1/*")
            If InStr(1, objNodeAux.nodeName, "Repet") = 0 Then
                If lvwMensagem.ColumnHeaders(objNodeAux.nodeName).SubItemIndex = 0 Then
                    objListItem.Text = objNodeAux.Text
                Else
                    objListItem.SubItems(lvwMensagem.ColumnHeaders(objNodeAux.nodeName).SubItemIndex) = objNodeAux.Text
                End If
            End If
        Next
    Next
    
    Set objMensagem = Nothing
    Set objXMLMensagem = Nothing
    Set objDomDocument = Nothing
    Set objDomAux = Nothing
    Set objDomDocumentFiltroXml = Nothing
    Set objNode = Nothing
    Set objNodeAux = Nothing
    Set objListItem = Nothing
    
Exit Sub
ErrorHandler:
    
    If Err.Number = 35601 Then Resume Next
    
    Set objMensagem = Nothing
    Set objXMLMensagem = Nothing
    Set objDomDocument = Nothing
    Set objDomAux = Nothing
    Set objDomDocumentFiltroXml = Nothing
    Set objNode = Nothing
    Set objNodeAux = Nothing
    Set objListItem = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaList", 0

End Sub

Private Sub flFormataListView(strFiltro As String)

On Error GoTo ErrorHandler

#If EnableSoap = 1 Then
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem        As A8MIU.clsMensagem
#End If

    Dim objDomDocument                      As MSXML2.DOMDocument
    Dim objNode                             As MSXML2.IXMLDOMNode
    Dim vntCodErro                          As Variant
    Dim vntMensagemErro                     As Variant
    Dim objDomDocumentFiltroXml             As MSXML2.DOMDocument
    

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    Set objDomDocument = CreateObject("MSXML2.DOMDocument")
    Set objDomDocumentFiltroXml = CreateObject("MSXML2.DOMDocument")
    
    objDomDocumentFiltroXml.loadXML strFiltro
    
    If blnUtilizaFiltro = False Then
        objDomDocument.loadXML objMensagem.LerMensagem(Left(cboMensagem.Text, 7) & "R1", 0, 0, vntCodErro, vntMensagemErro, 0)
    Else
        objDomDocument.loadXML objMensagem.LerMensagem(objDomDocumentFiltroXml.selectSingleNode("//CodigoMensagem").Text, 0, 0, vntCodErro, vntMensagemErro, 0)
    End If
    
    lvwMensagem.ListItems.Clear
    lvwMensagem.ColumnHeaders.Clear
    
    For Each objNode In objDomDocument.selectNodes("Repeat_Mensagem/*")
        If objNode.selectSingleNode("IN_NIVE_REPE").Text = 1 Then
            If UCase(Left(objNode.selectSingleNode("NO_TAG").Text, 5)) <> "REPET" _
            And UCase(Left(objNode.selectSingleNode("NO_TAG").Text, 6)) <> "/REPET" Then
                lvwMensagem.ColumnHeaders.Add , objNode.selectSingleNode("NO_TAG").Text, objNode.selectSingleNode("DE_TAG").Text
            End If
        End If
    Next
    
    Set objMensagem = Nothing
    Set objDomDocument = Nothing
    Set objNode = Nothing
    Set objDomDocumentFiltroXml = Nothing
      
Exit Sub
ErrorHandler:
    
    Set objMensagem = Nothing
    Set objDomDocument = Nothing
    Set objNode = Nothing
    Set objDomDocumentFiltroXml = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Monta o XML com os dados de filtro para seleção de operações
Private Function flMontarXMLFiltroPesquisa() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlFiltros.loadXML strFiltroXML
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")

    If blnUtilizaFiltro = False Then
        If Not xmlFiltros.documentElement.selectSingleNode("//CodigoMensagem") Is Nothing Then
            xmlFiltros.documentElement.selectSingleNode("//CodigoMensagem").Text = Left(cboMensagem.Text, 7) & "R1"
        End If
    Else
        If Not xmlFiltros.documentElement.selectSingleNode("//CodigoMensagem") Is Nothing Then
            xmlFiltros.documentElement.selectSingleNode("//CodigoMensagem").Text = xmlFiltros.documentElement.selectSingleNode("//CodigoMensagem").Text & "R1"
        End If
    End If
    
    flMontarXMLFiltroPesquisa = xmlFiltros.xml
    
    Set xmlFiltros = Nothing

Exit Function
ErrorHandler:
    Set xmlFiltros = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "showfiltro"
            Set objFiltro = Nothing
            Set objFiltro = New frmFiltro
            Set objFiltro.FormOwner = Me
            objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaMensagemSisbacen
            objFiltro.Show vbModal
        Case "refresh"
            fgCursor True
            If blnFiltroAplicado = True Then
                blnOrigemBotaoRefresh = True
                objFiltro.fgCarregarPesquisaAnterior
                blnOrigemBotaoRefresh = False
            Else
                If blnUtilizaFiltro = False Then
                    Call flCarregaList
                    blnFiltroAplicado = False
                End If
            End If
            fgCursor
        Case gstrSair
            Unload Me
    End Select

End Sub

Private Sub lvwMensagem_DblClick()

On Error GoTo ErrorHandler

    fgCursor True

    If Not lvwMensagem.SelectedItem Is Nothing Then
        With frmDetalheMensagemCAM
            .CodigoXml = lvwMensagem.SelectedItem.Tag
            .Show vbModal
        End With
    End If
    
    fgCursor
    
Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_DblClick", Me.Caption

End Sub
