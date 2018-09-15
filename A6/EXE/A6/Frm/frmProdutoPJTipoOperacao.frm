VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProdutoPJTipoOperacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Produto x Tipo de Operação"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   14370
   Tag             =   "Tipos de Operação"
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   6090
      TabIndex        =   10
      Top             =   45
      Width           =   8205
      Begin VB.OptionButton optFiltro 
         Caption         =   "Produtos Não Associados"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Produtos Associados"
         Height          =   195
         Index           =   1
         Left            =   2662
         TabIndex        =   12
         Top             =   270
         Width           =   1905
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Produtos Associados com Fim Vigência"
         Height          =   195
         Index           =   2
         Left            =   4950
         TabIndex        =   11
         Top             =   270
         Width           =   3150
      End
   End
   Begin VB.Frame fraMoldura 
      Height          =   740
      Index           =   2
      Left            =   6090
      TabIndex        =   7
      Top             =   8040
      Width           =   8205
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   248
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53936129
         CurrentDate     =   37622
         MaxDate         =   73050
         MinDate         =   37622
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   5100
         TabIndex        =   3
         Top             =   255
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   53936129
         CurrentDate     =   37622
         MaxDate         =   73050
         MinDate         =   37622
      End
      Begin MSComctlLib.Toolbar tblAplicar 
         Height          =   330
         Left            =   6660
         TabIndex        =   4
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Aplicar"
               Key             =   "Aplicar"
               ImageIndex      =   2
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Label lblDataFimVigencia 
         AutoSize        =   -1  'True
         Caption         =   "Data Fim Vigência"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   308
         Width           =   1365
      End
      Begin VB.Label lblDataInicioVigencia 
         AutoSize        =   -1  'True
         Caption         =   "Data Início Vigência"
         Height          =   195
         Left            =   255
         TabIndex        =   8
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.Frame fraMoldura 
      Caption         =   "Tipo de Operações"
      Height          =   8730
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   45
      Width           =   5895
      Begin MSComctlLib.ListView lstTipoOperacao 
         Height          =   8325
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   14684
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1413
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   68792
         EndProperty
      End
   End
   Begin VB.Frame fraMoldura 
      Caption         =   "Produtos "
      Height          =   7275
      Index           =   1
      Left            =   6090
      TabIndex        =   5
      Top             =   720
      Width           =   8205
      Begin MSComctlLib.ListView lstProduto 
         Height          =   6930
         Left            =   135
         TabIndex        =   1
         Tag             =   "Produtos PJ"
         Top             =   225
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   12224
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
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
            Text            =   "Código"
            Object.Width           =   1413
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   6985
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Início Vigência"
            Object.Width           =   2189
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fim Vigência"
            Object.Width           =   2118
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   105
      Top             =   8700
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
            Picture         =   "frmProdutoPJTipoOperacao.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdutoPJTipoOperacao.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdutoPJTipoOperacao.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdutoPJTipoOperacao.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdutoPJTipoOperacao.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdutoPJTipoOperacao.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdutoPJTipoOperacao.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   11370
      TabIndex        =   14
      Top             =   8880
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProdutoPJTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a administração da associação entre produtos PJ e
' tipos de operação.

Option Explicit

Private xmlProduto                          As MSXML2.DOMDocument40
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private intPosLinha                         As Integer
Private lngTipoOperacaoAlterado             As Long
Private datDataServidor                     As Date

' Aciona a atualização da tabela de associação produto Pj x tipo de operação.

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsProdutoPJTipoOperacao
#End If

Dim xmlDomSalvar        As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim objListItem         As MSComctlLib.ListItem
Dim lngItem             As Long
Dim blnExisteOperacao   As Boolean
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    blnExisteOperacao = False

    Set xmlDomSalvar = CreateObject("MSXML2.DOMDocument.4.0")

    fgAppendNode xmlDomSalvar, "", "Repeat_TipoOperacaoProduto", ""

    For Each objListItem In lstProduto.ListItems
        If objListItem.Tag = enumTipoOperacao.Incluir Or _
            objListItem.Tag = enumTipoOperacao.Alterar Then

            Set xmlNode = xmlProduto.selectSingleNode("Repeat_ProdutoPJ/Grupo_ProdutoPJ")

            If objListItem.Tag = enumTipoOperacao.Incluir Then

                xmlNode.selectSingleNode("@Operacao").Text = "Incluir"
                xmlNode.selectSingleNode("TP_OPER").Text = lngTipoOperacaoAlterado
                xmlNode.selectSingleNode("CO_PROD").Text = Mid(objListItem.Key, 2)

            ElseIf objListItem.Tag = enumTipoOperacao.Alterar Then
            
                For Each xmlNode In xmlProduto.documentElement.selectNodes("//Repeat_ProdutoPJ/Grupo_ProdutoPJ")
                    
                    If xmlNode.selectSingleNode("TP_OPER").Text = lngTipoOperacaoAlterado And _
                       xmlNode.selectSingleNode("CO_PROD").Text = Mid(objListItem.Key, 2) Then
                       
                       xmlNode.selectSingleNode("@Operacao").Text = "Alterar"
                       
                       Exit For
                    End If
                Next
            End If
            
            xmlNode.selectSingleNode("DT_INIC_VIGE").Text = Format(objListItem.SubItems(2), "YYYYMMDD")
            xmlNode.selectSingleNode("DT_FIM_VIGE").Text = IIf(Trim(objListItem.SubItems(3)) = "", "", Format(objListItem.SubItems(3), "YYYYMMDD"))
            fgAppendXML xmlDomSalvar, "Repeat_TipoOperacaoProduto", xmlNode.xml
            blnExisteOperacao = True

        End If
    Next
    
    If blnExisteOperacao Then
        Set objMIU = fgCriarObjetoMIU("A6MIU.clsProdutoPJTipoOperacao")
        Call objMIU.Executar(xmlDomSalvar.xml, _
                             vntCodErro, _
                             vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        fgCursor
        
        MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
        If optFiltro.Item(0).Value Then
            lngItem = 0
        ElseIf optFiltro.Item(1).Value Then
            lngItem = 1
        ElseIf optFiltro.Item(2).Value Then
            lngItem = 2
        End If
        lngTipoOperacaoAlterado = 0
        flVerificarProdutos lngItem
    Else
        MsgBox "Não existe Operação a ser executada.", vbInformation, Me.Caption
    End If
    
    Set objMIU = Nothing
    Set xmlDomSalvar = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlDomSalvar = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "frmProdutoPJTipoOperacao - flSalvar", 0

End Sub

' Aciona a exclusão de uma associação produto PJ x tipo de operação.

Private Sub flExcluir()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsProdutoPJTipoOperacao
#End If

Dim xmlDomSalvar        As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim lngItem             As Long
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    If lstProduto.SelectedItem Is Nothing Then
        frmMural.Display = "Selecione uma Associação de Produto a ser excluída."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If
    
    If MsgBox("Confirma a exclusão da Associação Tipo de Operação X Produto selecionada ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If

    fgCursor True
    Set xmlDomSalvar = CreateObject("MSXML2.DOMDocument.4.0")

    fgAppendNode xmlDomSalvar, "", "Repeat_TipoOperacaoProduto", ""

    For Each xmlNode In xmlProduto.documentElement.selectNodes("//Repeat_ProdutoPJ/Grupo_ProdutoPJ")
        If xmlNode.selectSingleNode("TP_OPER").Text = Mid(lstTipoOperacao.SelectedItem.Key, 2) And _
           xmlNode.selectSingleNode("CO_PROD").Text = Mid(lstProduto.SelectedItem.Key, 2) Then
           
           xmlNode.selectSingleNode("@Operacao").Text = "Excluir"
           Exit For
        End If
    Next
    
    fgAppendXML xmlDomSalvar, "Repeat_TipoOperacaoProduto", xmlNode.xml
    
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsProdutoPJTipoOperacao")
    Call objMIU.Executar(xmlDomSalvar.xml, _
                         vntCodErro, _
                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    If optFiltro.Item(0).Value Then
        lngItem = 0
    ElseIf optFiltro.Item(1).Value Then
        lngItem = 1
    ElseIf optFiltro.Item(2).Value Then
        lngItem = 2
    End If
        
    lngTipoOperacaoAlterado = 0
    flVerificarProdutos lngItem
    
    Set objMIU = Nothing
    Set xmlDomSalvar = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlDomSalvar = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "frmProdutoPJTipoOperacao - flExcluir", 0

End Sub

' Carrega lista de produtos PJ.

Private Sub flCarregarlstProduto()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    flInicializarFuncoes

    With lstProduto
        .ListItems.Clear
        For Each objDomNode In xmlProduto.selectNodes("Repeat_ProdutoPJ/*")

            Set objListItem = .ListItems.Add(, "K" & objDomNode.selectSingleNode("CO_PROD").Text, objDomNode.selectSingleNode("CO_PROD").Text)
            objListItem.SubItems(1) = objDomNode.selectSingleNode("DE_PROD").Text
            objListItem.SubItems(2) = fgDtXML_To_Interface(IIf(objDomNode.selectSingleNode("DT_INIC_VIGE").Text = "00:00:00", "", objDomNode.selectSingleNode("DT_INIC_VIGE").Text))
            objListItem.SubItems(3) = fgDtXML_To_Interface(IIf(objDomNode.selectSingleNode("DT_FIM_VIGE").Text = "00:00:00", "", objDomNode.selectSingleNode("DT_FIM_VIGE").Text))

        Next
    End With
    
    Set objDomNode = Nothing
    Set objListItem = Nothing

    Exit Sub

ErrorHandler:
    Set objDomNode = Nothing
    Set objListItem = Nothing
    
    fgRaiseError App.EXEName, Me.Name, "frmProdutoPJTipoOperacao - flCarregarlstProduto", 0

End Sub

' Carrega lista de tipos de operação.

Private Sub flCarregarlstTipoOperacao()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    With lstTipoOperacao
        .ListItems.Clear
        For Each objDomNode In xmlMapaNavegacao.selectNodes("ProdutoPJTipoOperacao/Repeat_TipoOperacao/*")
            Set objListItem = .ListItems.Add(, "K" & objDomNode.selectSingleNode("TP_OPER").Text, objDomNode.selectSingleNode("TP_OPER").Text)
            objListItem.SubItems(1) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
        Next
    End With
    
    Set objDomNode = Nothing
    Set objListItem = Nothing
    
    Exit Sub
ErrorHandler:

    Set objDomNode = Nothing
    Set objListItem = Nothing

    fgRaiseError App.EXEName, Me.Name, "frmProdutoPJTipoOperacao - flCarregarlstTipoOperacao", 0
End Sub

' Carrega configurações iniciais do formulário.

Private Sub flInicializar(ByVal plStatusProduto As enumStatusProduto)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsProdutoPJTipoOperacao
#End If

Dim strMapaNavegacao    As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsProdutoPJTipoOperacao")
    strMapaNavegacao = objMIU.ObterProdutoPJTipoOperacao(plStatusProduto, _
                                                         vntCodErro, _
                                                         vntMensagemErro)
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlProduto = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmProdutoPJTipoOperacao", "flInicializar")
    End If
    
    If Not xmlMapaNavegacao.documentElement.selectSingleNode("Repeat_ProdutoPJ") Is Nothing Then
    
        If Not xmlProduto.loadXML(xmlMapaNavegacao.documentElement.selectSingleNode("Repeat_ProdutoPJ").xml) Then
            Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmProdutoPJTipoOperacao", "flInicializar")
        End If
        
        datDataServidor = fgDtXML_To_Date(xmlProduto.documentElement.selectSingleNode("Grupo_ProdutoPJ/DT_BASE_DADO").Text)
    End If
    
    Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmProdutoPJTipoOperacao - flInicializar", 0

End Sub

Private Sub dtpFim_Change()

On Error GoTo ErrorHandler

    Call fgDataVigenciaFimChange(dtpInicio, dtpFim)

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - dtpFim_Change"
   
End Sub

Private Sub dtpFim_KeyPress(KeyAscii As Integer)
    
    KeyAscii = 0

End Sub

Private Sub dtpInicio_Change()

On Error GoTo ErrorHandler

    Call fgDataVigenciaInicioChange(dtpInicio, dtpFim)

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - dtpInicio_Change"
   
End Sub

Private Sub dtpInicio_KeyPress(KeyAscii As Integer)
    
    KeyAscii = 0

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Me.Icon = mdiSBR.Icon
    
    fgCursor True
    fgCenterMe Me
    
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Me.Show
    DoEvents
        
    flInicializar enumStatusProduto.NaoAssociado
    
    flCarregarlstTipoOperacao
    flCarregarlstProduto
    fgCursor
    
    Exit Sub
ErrorHandler:

    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - Form_Load")
    
End Sub

Private Sub lstProduto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstProduto, ColumnHeader.Index)
    
    Exit Sub
ErrorHandler:
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - lstProduto_ColumnClick")
    
End Sub

Private Sub lstProduto_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim strDataInicio                           As String
Dim strDataFim                              As String

On Error GoTo ErrorHandler

    fgCursor True
    
    Item.Selected = True
    intPosLinha = Item.Index

    tblAplicar.Enabled = True
    
    dtpFim.Enabled = True
    dtpInicio.Enabled = True
    
    If Item.SubItems(2) = vbNullString Then
        strDataInicio = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    Else
        strDataInicio = fgDt_To_Xml(CDate(Item.SubItems(2)))
    End If
    
    If Item.SubItems(3) = vbNullString Then
        strDataFim = datDataVazia
    Else
        strDataFim = fgDt_To_Xml(CDate(Item.SubItems(3)))
    End If

    Call fgCarregaDataVigencia(dtpInicio, dtpFim, strDataInicio, strDataFim)
    
    If optFiltro.Item(0).Value Then
        dtpInicio.Enabled = True
        tlbCadastro.Buttons("Excluir").Enabled = False
    Else
        dtpInicio.Enabled = False
        tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
    End If
    
    fgCursor

    Exit Sub
ErrorHandler:
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - lstProduto_ItemClick")

End Sub

Private Sub lstTipoOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lstTipoOperacao, ColumnHeader.Index)
    
    Exit Sub
ErrorHandler:
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - lstTipoOperacao_ColumnClick")
    
End Sub

Private Sub lstTipoOperacao_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim lngItem                                 As Long

On Error GoTo ErrorHandler

    If gblnPerfilManutencao Then
        If lngTipoOperacaoAlterado > 0 Then
            If MsgBox("Associação de Tipo de Operação com Produto sofreu alteração. Deseja Salvar? ", vbYesNo) = vbYes Then
                fgCursor True
                flSalvar
            End If
        End If
    End If
    
    If optFiltro.Item(0).Value Then
        lngItem = 0
    ElseIf optFiltro.Item(1).Value Then
        lngItem = 1
    ElseIf optFiltro.Item(2).Value Then
        lngItem = 2
    End If

    flVerificarProdutos lngItem
    fgCursor

    Exit Sub

ErrorHandler:
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - lstTipoOperacao_ItemClick")

End Sub

Private Sub optFiltro_Click(Index As Integer)

On Error GoTo ErrorHandler
    
    If gblnPerfilManutencao Then
        
        If lngTipoOperacaoAlterado > 0 Then
            If MsgBox("Associação de Tipo de Operação com Produto sofreu alteração. Deseja Salvar? ", vbYesNo) = vbYes Then
                fgCursor True
                flSalvar
            End If
        End If
    
        flVerificarProdutos Index
    
    End If
    
    fgCursor

    Exit Sub
ErrorHandler:
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - optFiltro_Click")

End Sub

' Aciona a leitura de produtos PJ respeitando sua associação ou não a um tipo de operação.

Private Sub flVerificarProdutos(ByVal plItem As Long)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsProdutoPJTipoOperacao
#End If

Dim lngStatusProduto    As enumStatusProduto
Dim strProdutos         As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    If Trim(lstTipoOperacao.SelectedItem.Key) = "" Then
        frmMural.Display = "Selecionar um Tipo de Operação."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If

    lstProduto.ListItems.Clear

    Select Case plItem
        Case 0
            lngStatusProduto = enumStatusProduto.NaoAssociado
        Case 1
            lngStatusProduto = enumStatusProduto.AssociadoSemFimVigencia
        Case 2
            lngStatusProduto = enumStatusProduto.AssociadoComFimVigencia
    End Select

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsProdutoPJTipoOperacao")
    strProdutos = objMIU.ObterProdutoPorStatus(lngStatusProduto, _
                                               CLng("0" & Mid(lstTipoOperacao.SelectedItem.Key, 2)), _
                                               vntCodErro, _
                                               vntMensagemErro)
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If Trim(strProdutos) = "" Then
        Exit Sub
    End If
    Set xmlProduto = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlProduto.loadXML(strProdutos) Then
        Call fgErroLoadXML(xmlProduto, App.EXEName, "frmProdutoPJTipoOperacao", "flVerificarProdutos")
    End If

    flCarregarlstProduto

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "frmProdutoPJTipoOperacao - flVerificarProdutos", 0

End Sub

Private Sub tblAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    If dtpFim.Value < dtpInicio.Value And Not IsNull(dtpFim.Value) Then
        frmMural.Display = "Data de Fim de Vigência não pode ser menor que a Data de Início."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If

    If dtpInicio.Value < datDataServidor And optFiltro(0) Then
        frmMural.Display = "Data de Início de Vigência não pode ser menor que a Data Atual."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        dtpInicio.Value = datDataServidor
        Exit Sub
    End If

    If dtpFim.Value < datDataServidor And Not IsNull(dtpFim.Value) Then
        frmMural.Display = "Data de Fim de Vigência não pode ser menor que a Data Atual."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        dtpFim.Value = datDataServidor
        Exit Sub
    End If

    If lstProduto.ListItems.Count = 0 Then
        frmMural.Display = "Não existem produtos a serem associados."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If
    
    lstProduto.ListItems(intPosLinha).SubItems(2) = dtpInicio.Value
    If IsNull(dtpFim.Value) Then
        lstProduto.ListItems(intPosLinha).SubItems(3) = ""
    Else
        lstProduto.ListItems(intPosLinha).SubItems(3) = dtpFim.Value
    End If

    'Controlar o Status da Coluna
    If optFiltro.Item(0).Value Then
        lstProduto.ListItems(intPosLinha).Tag = enumTipoOperacao.Incluir
    Else
        lstProduto.ListItems(intPosLinha).Tag = enumTipoOperacao.Alterar
    End If

    lngTipoOperacaoAlterado = Mid(lstTipoOperacao.SelectedItem.Key, 2)

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tblAplicar_ButtonClick"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim lngItem                                 As Long

On Error GoTo ErrorHandler

    Select Case Button.Key
    Case "Excluir"
        flExcluir
    Case "Salvar"
        If lngTipoOperacaoAlterado > 0 Then
            fgCursor True
            flSalvar
        End If
    Case "Sair"
        Unload Me
    End Select
    
    fgCursor
    
    Exit Sub

ErrorHandler:
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmProdutoPJTipoOperacao - tlbCadastro_ButtonClick")
    
    If optFiltro.Item(0).Value Then
        lngItem = 0
    ElseIf optFiltro.Item(1).Value Then
        lngItem = 1
    ElseIf optFiltro.Item(2).Value Then
        lngItem = 2
    End If
    lngTipoOperacaoAlterado = 0
    
    fgCursor True
    flVerificarProdutos lngItem
    fgCursor

End Sub

' Inicializa funcionalidades e campos da tela.

Private Sub flInicializarFuncoes()

On Error GoTo ErrorHandler

    dtpInicio.Value = fgDataHoraServidor(Data)
    If dtpFim.MinDate > fgDataHoraServidor(Data) Then
        dtpFim.Value = dtpFim.MinDate
    Else
        dtpFim.Value = fgDataHoraServidor(Data)
    End If
    dtpFim.Value = Null
    
    dtpInicio.Enabled = False
    dtpFim.Enabled = False
    
    tblAplicar.Enabled = False
    tlbCadastro.Buttons("Excluir").Enabled = False
    lngTipoOperacaoAlterado = 0

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "frmProdutoPJTipoOperacao - flInicializarFuncoes", 0

End Sub
