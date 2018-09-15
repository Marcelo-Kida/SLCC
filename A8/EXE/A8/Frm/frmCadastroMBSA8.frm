VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroConversaoMBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segregação Acesso - Controle Acesso Dados"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11850
   Begin VB.Frame Frame2 
      Caption         =   "Consultar"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4860
      Width           =   4335
      Begin VB.CommandButton cmdOcorrencias 
         Caption         =   "Divergência A8 X MBS"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton cmdConsultaTodas 
         Caption         =   "Todas as Associações"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   2220
         TabIndex        =   9
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame fraTipoInformacao 
      Caption         =   "Tipo Associação"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   60
      Width           =   11595
      Begin VB.OptionButton optTipoInformacao 
         Caption         =   "Tipo Back Office"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1800
      End
      Begin VB.OptionButton optTipoInformacao 
         Caption         =   "Grupo Usuário"
         Height          =   255
         Index           =   4
         Left            =   8760
         TabIndex        =   3
         Top             =   360
         Width           =   1800
      End
      Begin VB.OptionButton optTipoInformacao 
         Caption         =   "Grupo Veículo Legal"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   1800
      End
      Begin VB.OptionButton optTipoInformacao 
         Caption         =   "Local Liquidação"
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhe"
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   11595
      Begin VB.TextBox txtNome 
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   6975
      End
      Begin VB.ComboBox cboInformacao 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1260
         Width           =   6975
      End
      Begin NumBox.Number numCodigo 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelecao     =   0   'False
         AceitaNegativo  =   0   'False
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Associação"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome MBS"
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código MBS"
         Height          =   195
         Left            =   1620
         TabIndex        =   12
         Top             =   360
         Width           =   885
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4920
      Top             =   4920
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
            Picture         =   "frmCadastroMBSA8.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroMBSA8.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroMBSA8.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroMBSA8.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroMBSA8.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroMBSA8.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroMBSA8.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstConversao 
      Height          =   2115
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   3731
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
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   7845
      TabIndex        =   10
      Top             =   5055
      Width           =   3975
      _ExtentX        =   7011
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
Attribute VB_Name = "frmCadastroConversaoMBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pelo cadastramento de conversão dos cadastro MBS para o Controle do SLCC,
'' através da camada de controle de caso de uso MIU.
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private xmlLerTodos                         As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmCadastroConversaoMBS"
Private penumTipoInformacao                 As enumTipoInformacao

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private blnTodasAssociacoes                 As Boolean

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Sub flDefinirTamanhoMaximoCampos()
    
    With xmlMapaNavegacao.documentElement
        txtNome.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_ControleAcessDado/NO_GRUP_GS/@Tamanho").Text
    End With

End Sub

'Formata o Cabeçalho e o tamanho de cada coluna do List View

Private Sub flFormataListView()

    lstConversao.ColumnHeaders.Clear
    lstConversao.ColumnHeaders.Add , , "Código MBS", 1100, lvwColumnLeft
    lstConversao.ColumnHeaders.Add , , "Nome MBS", 2000, lvwColumnLeft
    lstConversao.ColumnHeaders.Add , , "Tipo Associação", 2000, lvwColumnLeft
    lstConversao.ColumnHeaders.Add , , "Código Associação", 1800, lvwColumnLeft
    lstConversao.ColumnHeaders.Add , , "Descrição Associação", 2200, lvwColumnLeft
    lstConversao.ColumnHeaders.Add , , "Data Início", 1100, lvwColumnLeft
    lstConversao.ColumnHeaders.Add , , "Data Fim ", 1100, lvwColumnLeft

End Sub

Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lstConversao.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstConversao.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstConversao_ItemClick objListItem
           lstConversao.ListItems(strKeyItemSelected).EnsureVisible
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
    
    Set objListItem = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

'' Salva as alterações efetuadas através da camada controladora de casos de uso
'' MIU, método A8MIU.clsMIU.Executar

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu             As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu             As A8MIU.clsMIU
#End If

Dim strRetorno             As String
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
    
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strOperacao = "Incluir" Then
           strKeyItemSelected = "K" & numCodigo.Valor
        End If
        
        strOperacao = gstrOperAlterar
        numCodigo.Enabled = False
    Else
        flLimpaCampos
    End If
    
    Set objMiu = Nothing
    
    Call flCarregaListView(penumTipoInformacao)
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption

    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing
    
    If strOperacao <> gstrOperExcluir Then
        With xmlLer.documentElement
             strKeyItemSelected = "K" & .selectSingleNode("//CO_GRUP_GS").Text
        End With
    End If
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString
Private Function flValidarCampos() As String
    
    If numCodigo = 0 Then
        flValidarCampos = "Digite o Código do Grupo MBS."
        numCodigo.SetFocus
        Exit Function
    End If
    
    If Len(txtNome.Text) = 0 Then
        flValidarCampos = "Digite o Nome do Grupo MBS."
        txtNome.SetFocus
        Exit Function
    End If
    
    If cboInformacao.ListIndex < 0 Then
        flValidarCampos = "Informe o Tipo da Informação do Grupo MBS."
        cboInformacao.SetFocus
        Exit Function
    End If
    
    flValidarCampos = ""

End Function

'Limpar as informações para uma nova inclusão

Private Sub flLimpaCampos()
        
    strOperacao = "Incluir"
    
    numCodigo.Valor = 0
    numCodigo.Enabled = True
    
    txtNome.Text = ""
        
    fraTipoInformacao.Enabled = True
    
    cboInformacao.Clear
    
    flCarregarCombo penumTipoInformacao
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    txtNome.Locked = False
    numCodigo.Enabled = True
    
    cmdOcorrencias.Enabled = True

End Sub

'' Preenche a interface de acordo com o documento XML

Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu             As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu             As A8MIU.clsMIU
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
        
    numCodigo.Enabled = False
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//CO_GRUP_GS").Text = lstConversao.SelectedItem.Text
    End With
    
    Set objMiu = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlLer.loadXML objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
   
    With xmlLer.documentElement
    
        numCodigo.Valor = CLng("0" & .selectSingleNode("CO_GRUP_GS").Text)
        txtNome.Text = .selectSingleNode("NO_GRUP_GS").Text
        fgSearchItemCombo cboInformacao, , .selectSingleNode("CO_INFO").Text
        
    End With
        
Exit Sub
    
ErrorHandler:
    
    Set objMiu = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0

End Sub

' Comverte os parametros da interface para um documento XML

Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
    
         .selectSingleNode("//@Operacao").Text = strOperacao
    
         If strOperacao = "Incluir" Then
            .selectSingleNode("//CO_GRUP_GS").Text = numCodigo.Valor
         ElseIf strOperacao = gstrOperExcluir Then
            Exit Function
         End If
          
         .selectSingleNode("//NO_GRUP_GS").Text = txtNome.Text
         
         If optTipoInformacao(enumTipoInformacao.GrupoVeiculoLegal).value = True Then
            .selectSingleNode("//TP_INFO").Text = enumTipoInformacao.GrupoVeiculoLegal
         ElseIf optTipoInformacao(enumTipoInformacao.GrupoUsuario).value = True Then
            .selectSingleNode("//TP_INFO").Text = enumTipoInformacao.GrupoUsuario
         ElseIf optTipoInformacao(enumTipoInformacao.TipoBackOffice).value = True Then
            .selectSingleNode("//TP_INFO").Text = enumTipoInformacao.TipoBackOffice
         ElseIf optTipoInformacao(enumTipoInformacao.LocalLiquidacao).value = True Then
            .selectSingleNode("//TP_INFO").Text = enumTipoInformacao.LocalLiquidacao
         End If
         
         .selectSingleNode("//CO_INFO").Text = fgObterCodigoCombo(cboInformacao.List(cboInformacao.ListIndex))
        
    End With
    
    Exit Function
    
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function

'' Carrega as propriedades necessárias a interface frmCadastroConversaoMBS, através da
'' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu             As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu             As A8MIU.clsMIU
#End If

Dim strMapaNavegacao       As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    Set objMiu = fgCriarObjetoMIU("A8MIU.clsMIU")
    strMapaNavegacao = objMiu.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmCadastroConversaoMBS", "flInicializar")
    End If
    
    If xmlLer Is Nothing Then
        Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
        If Not xmlLer.loadXML(xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_ControleAcessDado").xml) Then
            Call fgErroLoadXML(xmlLer, App.EXEName, "frmCadastroConversaoMBS", "flInicializar")
        End If
    End If
    
    Exit Sub

ErrorHandler:

    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0
    
End Sub

' Preenche o combo de acordo com o documento XML

Private Sub flCarregarCombo(ByVal penumTipoInformacao As enumTipoInformacao)

#If EnableSoap = 1 Then
    Dim objMiu             As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu             As A8MIU.clsMIU
#End If

Dim xmlLerTodos            As MSXML2.DOMDocument40
Dim strLerTodos            As String
Dim strPropriedades        As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    cboInformacao.Enabled = True
    
    Set objMiu = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Select Case penumTipoInformacao
            
           Case enumTipoInformacao.GrupoVeiculoLegal
           
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_VIGE").Text = "S"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_SEGR").Text = "N"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/@Operacao").Text = "LerTodos"
                strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal").xml
                strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
                
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
                
                If strLerTodos = "" Then Exit Sub
                Call xmlLerTodos.loadXML(strLerTodos)
           
                Call fgCarregarCombos(cboInformacao, xmlLerTodos, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA")
           
           Case enumTipoInformacao.GrupoUsuario
           
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/TP_VIGE").Text = "S"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/TP_SEGR").Text = "N"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/@Operacao").Text = "LerTodos"
                strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario").xml
                strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
                
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
                
                If strLerTodos = "" Then Exit Sub
                Call xmlLerTodos.loadXML(strLerTodos)
           
                Call fgCarregarCombos(cboInformacao, xmlLerTodos, "GrupoUsuario", "CO_GRUP_USUA", "NO_GRUP_USUA")
           
           Case enumTipoInformacao.TipoBackOffice
           
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice/TP_VIGE").Text = "S"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice/TP_SEGR").Text = "N"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice/@Operacao").Text = "LerTodos"
                strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice").xml
                strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
                
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
                
                If strLerTodos = "" Then Exit Sub
                Call xmlLerTodos.loadXML(strLerTodos)
           
                Call fgCarregarCombos(cboInformacao, xmlLerTodos, "TipoBackOffice", "TP_BKOF", "DE_BKOF")
           
           Case enumTipoInformacao.LocalLiquidacao
           
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_LocalLiquidacao/TP_VIGE").Text = "S"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_LocalLiquidacao/TP_SEGR").Text = "N"
                xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_LocalLiquidacao/@Operacao").Text = "LerTodos"
                strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_LocalLiquidacao").xml
                strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
                
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
                
                If strLerTodos = "" Then Exit Sub
                Call xmlLerTodos.loadXML(strLerTodos)
           
                Call fgCarregarCombos(cboInformacao, xmlLerTodos, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU")
    
    End Select
    
    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Exit Sub

ErrorHandler:

    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarCombo", 0

End Sub

'' Carrega os cadastros já existentes e preenche o listview com os mesmos,
'' através da classe controladora de caso de uso MIU, método  A8MIU.clsMIU.Executar

Private Sub flCarregaListView(ByVal penumTipoInformacao As enumTipoInformacao)

#If EnableSoap = 1 Then
    Dim objMiu             As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu             As A8MIU.clsMIU
#End If

Dim xmlDomNode             As MSXML2.IXMLDOMNode
Dim objListItem            As MSComctlLib.ListItem
Dim strPropriedades        As String
Dim strLerTodos            As String
Dim strTipoInformacao      As String
Dim xmlLerTodos            As MSXML2.DOMDocument40
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    lstConversao.ListItems.Clear
    
    Set xmlDomNode = Nothing
    
    Set objMiu = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ControleAcessDado/TP_INFO").Text = penumTipoInformacao
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ControleAcessDado/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ControleAcessDado").xml
    strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)

    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_ControleAcessDado/*")
        With xmlDomNode
         
            Select Case .selectSingleNode("TP_INFO").Text
                   Case enumTipoInformacao.GrupoUsuario
                        strTipoInformacao = "Grupo Usuário"
                   Case enumTipoInformacao.GrupoVeiculoLegal
                        strTipoInformacao = "Grupo Veículo Legal"
                   Case enumTipoInformacao.LocalLiquidacao
                        strTipoInformacao = "Local Liquidação"
                   Case enumTipoInformacao.TipoBackOffice
                        strTipoInformacao = "Tipo Back Office"
            End Select
            
            Set objListItem = lstConversao.ListItems.Add(, "K" & .selectSingleNode("CO_GRUP_GS").Text, .selectSingleNode("CO_GRUP_GS").Text)
            objListItem.SubItems(1) = .selectSingleNode("NO_GRUP_GS").Text
            objListItem.SubItems(2) = strTipoInformacao
            objListItem.SubItems(3) = .selectSingleNode("CO_INFO").Text
            objListItem.SubItems(4) = .selectSingleNode("DE_INFO").Text
            objListItem.SubItems(5) = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE").Text)
            
            If CStr(.selectSingleNode("DT_FIM_VIGE").Text) <> gstrDataVazia Then
                objListItem.SubItems(6) = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text)
            End If
            
        End With
    Next
    
    Set objMiu = Nothing
    Set xmlLerTodos = Nothing

    Exit Sub
    
ErrorHandler:

    Set objMiu = Nothing
    Set xmlLerTodos = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0
    
End Sub

Private Sub cmdConsultaTodas_Click()
     
On Error GoTo ErrorHandler

    fgCursor True

    flLimpaCampos
    
    cboInformacao.Clear
    
    optTipoInformacao(1).value = False
    optTipoInformacao(2).value = False
    optTipoInformacao(3).value = False
    optTipoInformacao(4).value = False
    
    tlbCadastro.Buttons("Limpar").Enabled = False
    tlbCadastro.Buttons(gstrSalvar).Enabled = False
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    cmdOcorrencias.Enabled = False
    txtNome.Locked = True
    numCodigo.Enabled = False
    
    Call flCarregaListView(TodasInformacoes)
    
    blnTodasAssociacoes = True
    
    fgCursor

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cmdConsultaTodas_Click"

End Sub

Private Sub lstConversao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lstConversao.Sorted = True
    lstConversao.SortKey = ColumnHeader.Index - 1

    If lstConversao.SortOrder = lvwAscending Then
        lstConversao.SortOrder = lvwDescending
    Else
        lstConversao.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstConversao_ColumnClick"

End Sub

Private Sub lstConversao_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
            
    Call fgCursor(True)
    
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface
    
    strKeyItemSelected = Item.Key
    
    numCodigo.Enabled = False
    
    If numCodigo.Valor > 0 Then
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    End If
    
    If blnTodasAssociacoes = True Then
       tlbCadastro.Buttons("Limpar").Enabled = False
       tlbCadastro.Buttons(gstrSalvar).Enabled = False
       tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
       cmdOcorrencias.Enabled = False
       txtNome.Locked = True
       numCodigo.Enabled = False
       cboInformacao.Clear
       cboInformacao.AddItem Item.SubItems(3) & " - " & Item.SubItems(4)
       cboInformacao.ListIndex = 0
    End If
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroConversaoMBS - lstConversao_ItemClick", Me.Caption
    
    flLimpaCampos
    If Not blnTodasAssociacoes Then
       Call flCarregaListView(penumTipoInformacao)
    Else
       cmdConsultaTodas_Click
    End If

End Sub

Private Sub optTipoInformacao_Click(Index As Integer)

On Error GoTo ErrorHandler

    DoEvents
    blnTodasAssociacoes = False
    
    fgLockWindow Me.hwnd
    penumTipoInformacao = Index
    Call fgCursor(True)
    Call flInicializar
    Call flFormataListView
    Call flLimpaCampos
    Call flCarregarCombo(penumTipoInformacao)
    Call flDefinirTamanhoMaximoCampos
    Call flCarregaListView(penumTipoInformacao)
    
    Call fgCursor(False)
    fgLockWindow 0
    
    Exit Sub

ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroConversaoMBS - optTipoInformacao_Click", Me.Caption

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
            If Frame1.Enabled = True And numCodigo.Enabled = True Then
               numCodigo.SetFocus
            End If
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
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroConversaoMBS - tlbCadastro_ButtonClick", Me.Caption
    
    Call flCarregaListView(penumTipoInformacao)
    
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub

Private Function flOcorrencias() As String

#If EnableSoap = 1 Then
    Dim objControleAcessDado    As MSSOAPLib30.SoapClient30
#Else
    Dim objControleAcessDado    As A8MIU.clsControleAcessDado
#End If

Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    Set objControleAcessDado = fgCriarObjetoMIU("A8MIU.clsControleAcessDado")
    flOcorrencias = objControleAcessDado.VerificaControleAcessDado(CInt(penumTipoInformacao), vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcessDado = Nothing
    Exit Function
    
ErrorHandler:
    
    Set objControleAcessDado = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flOcorrencias", 0
    
End Function

Private Sub cmdOcorrencias_Click()

On Error GoTo ErrorHandler

    frmTipoInformacao.pstrXMLOcorrencias = flOcorrencias
    frmTipoInformacao.Show , mdiLQS
    
    Exit Sub

ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroConversaoMBS - cmdOcorrencias_Click", Me.Caption
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCursor
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    
    optTipoInformacao(1).value = True
    
    Exit Sub

ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroConversaoMBS - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmCadastroConversaoMBS = Nothing

End Sub

