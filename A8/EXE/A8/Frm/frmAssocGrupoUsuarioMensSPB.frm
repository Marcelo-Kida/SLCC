VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAssocGrupoUsuarioMensSPB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segregação Acesso - Grupo Usuário X Mensagens SPB"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   14535
   Begin VB.Frame fraAcessoPadrao 
      Caption         =   "Alterar Acesso"
      Height          =   615
      Left            =   12480
      TabIndex        =   8
      Top             =   120
      Width           =   1995
      Begin VB.OptionButton optConsultar 
         Caption         =   "&Consultar"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   250
         Width           =   975
      End
      Begin VB.OptionButton optEnviar 
         Caption         =   "&Enviar"
         Height          =   195
         Left            =   1140
         TabIndex        =   9
         Top             =   250
         Width           =   795
      End
   End
   Begin VB.Frame fraPrincipal 
      Height          =   4875
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   14475
      Begin MSComctlLib.ListView lvwMsg 
         Height          =   4335
         Left            =   60
         TabIndex        =   6
         Top             =   420
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwMsgAutorizadas 
         Height          =   4335
         Left            =   7740
         TabIndex        =   7
         Top             =   420
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Associar ->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   6060
         TabIndex        =   13
         Top             =   360
         Width           =   1575
         Begin VB.CommandButton cmdAdicionar 
            Caption         =   "Enviar->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   1740
            Width           =   1335
         End
         Begin VB.CommandButton cmdAdicionarTodas 
            Caption         =   "Enviar->>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   2100
            Width           =   1335
         End
         Begin VB.CommandButton cmdAdicionar 
            Caption         =   "Consultar->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   540
            Width           =   1335
         End
         Begin VB.CommandButton cmdAdicionarTodas 
            Caption         =   "Consultar->>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   900
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "<-Desassociar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   6060
         TabIndex        =   14
         Top             =   3480
         Width           =   1575
         Begin VB.CommandButton cmdRemover 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdRemoverTodas 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   780
            Width           =   1335
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Mensagens associadas"
         Height          =   255
         Left            =   7740
         TabIndex        =   11
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Mensagens para associação"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.ComboBox cboGrupoUsuario 
      Height          =   315
      Left            =   6900
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   5535
   End
   Begin VB.ComboBox cboTipoBackOffice 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Width           =   5835
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   5400
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
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssocGrupoUsuarioMensSPB.frx":195E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   12420
      TabIndex        =   4
      Top             =   5640
      Width           =   2055
      _ExtentX        =   3625
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
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo Usuário"
      Height          =   195
      Left            =   6900
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Back Office"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "frmAssocGrupoUsuarioMensSPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'' Objeto responsável por associar tipos de mensagens SPB a grupos de usuário e
'' tipo de acesso(Consulta/Envio), através da camada de controle de caso de uso
'' MIU.
''
'' São consideradas especificamente classes de destino:
''    A8MIU.clsMIU
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLerTodos                         As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmAssocGrupoUsuarioMensSPB"

Private Const TIPOACESSO_ENVIAR             As String = "Enviar"
Private Const TIPOACESSO_CONSULTAR          As String = "Consultar"

'' Configura os cabeçalhos do listview de acordo com os campos necessários
Private Sub flFormataListView()

On Error GoTo ErrorHandler

    lvwMsg.ColumnHeaders.Clear
    lvwMsg.ColumnHeaders.Add , , "Código Mensagem", 1600, lvwColumnLeft
    lvwMsg.ColumnHeaders.Add , , "Descrição da Mensagem", 4140, lvwColumnLeft
    
    lvwMsgAutorizadas.ColumnHeaders.Clear
    lvwMsgAutorizadas.ColumnHeaders.Add , , "Código Mensagem", 1600, lvwColumnLeft
    lvwMsgAutorizadas.ColumnHeaders.Add , , "Descrição da Mensagem", 3770, lvwColumnLeft
    lvwMsgAutorizadas.ColumnHeaders.Add , , "Acesso ", 1200, lvwColumnLeft

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormataListView", 0

End Sub

'' Salva as alterações através da camada de controle de casos de uso MIU, através
'' do método:A8MIU.clsMIU.Executar
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim xmlSalvar              As MSXML2.DOMDocument40
Dim lstItem                As MSComctlLib.ListItem
Dim objDomNode             As MSXML2.IXMLDOMNode
Dim blnSalvar              As Boolean
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    fgCursor True

    Set xmlSalvar = CreateObject("MSXML2.DOMDocument.4.0")
    fgAppendNode xmlSalvar, "", "Repeat_GrupoUsuarioMensSPB", ""
    blnSalvar = False

    For Each lstItem In lvwMsgAutorizadas.ListItems
        If lstItem.Tag = enumTipoOperacao.Incluir Then
            Set objDomNode = xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_GrupoUsuarioMensSPB")
            objDomNode.selectSingleNode("@Operacao").Text = "Incluir"
            objDomNode.selectSingleNode("TP_BKOF").Text = fgObterCodigoCombo(cboTipoBackOffice.Text)
            objDomNode.selectSingleNode("CO_GRUP_USUA").Text = fgObterCodigoCombo(cboGrupoUsuario.Text)
            objDomNode.selectSingleNode("CO_MESG_SPB").Text = Mid$(lstItem.Key, 2)
            objDomNode.selectSingleNode("TP_ACES").Text = IIf(lstItem.SubItems(2) = TIPOACESSO_CONSULTAR, enumTipoAcessoMensagemSPB.Consultar, enumTipoAcessoMensagemSPB.Enviar)
            fgAppendXML xmlSalvar, "Repeat_GrupoUsuarioMensSPB", objDomNode.xml
            blnSalvar = True
        ElseIf lstItem.Tag = enumTipoOperacao.Alterar Then
            Set objDomNode = xmlLerTodos.selectSingleNode("/Repeat_SistemaMensagem/Grupo_SistemaMensagem[CO_MESG_SPB='" & Mid$(lstItem.Key, 2) & "  ']")
            objDomNode.selectSingleNode("@Operacao").Text = "Alterar"
            objDomNode.selectSingleNode("TP_ACES").Text = IIf(lstItem.SubItems(2) = TIPOACESSO_CONSULTAR, enumTipoAcessoMensagemSPB.Consultar, enumTipoAcessoMensagemSPB.Enviar)
            fgAppendXML xmlSalvar, "Repeat_GrupoUsuarioMensSPB", objDomNode.xml
            blnSalvar = True
        End If
        
    Next lstItem
    
    For Each lstItem In lvwMsg.ListItems
        If lstItem.Tag = enumTipoOperacao.Excluir Then
            If xmlLerTodos.xml = Empty Then Exit For
            Set objDomNode = xmlLerTodos.selectSingleNode("/Repeat_SistemaMensagem/Grupo_SistemaMensagem[CO_MESG_SPB='" & Mid$(lstItem.Key, 2) & "  ']")
            objDomNode.selectSingleNode("@Operacao").Text = "Excluir"
            fgAppendXML xmlSalvar, "Repeat_GrupoUsuarioMensSPB", objDomNode.xml
            blnSalvar = True
        End If
    Next lstItem

    If blnSalvar Then
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
        objMIU.Executar xmlSalvar.xml, vntCodErro, vntMensagemErro
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        flCarregarListView
        Set objMIU = Nothing
        fgCursor
        MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    Else
        fgCursor
    End If

Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    fgCursor
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'' Obtém as propriedades e demais dados dinâmicos da tela, através de interação
'' com a camada de controle de caso de uso MIU, método  A8MIU.clsMiu.
'' ObterMapaNavegacao
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    If xmlLerTodos Is Nothing Then
        Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    End If
    
    'Carregar Tipo BackOffice
    xmlLerTodos.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice").xml
    xmlLerTodos.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
    xmlLerTodos.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("TP_SEGR").Text = "N"
    
    xmlLerTodos.loadXML objMIU.Executar(xmlLerTodos.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call fgCarregarCombos(Me.cboTipoBackOffice, xmlLerTodos, "TipoBackOffice", "TP_BKOF", "DE_BKOF")
    
    'Carregar Grupos de Usuário
    xmlLerTodos.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario").xml
    xmlLerTodos.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
    xmlLerTodos.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("TP_SEGR").Text = "N"
    
    xmlLerTodos.loadXML objMIU.Executar(xmlLerTodos.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call fgCarregarCombos(Me.cboGrupoUsuario, xmlLerTodos, "GrupoUsuario", "CO_GRUP_USUA", "NO_GRUP_USUA")
    
    Call fgCursor(False)
    Set objMIU = Nothing
    
    Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

Private Sub cboGrupoUsuario_Click()

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    If cboGrupoUsuario.ListIndex > -1 And cboTipoBackOffice.ListIndex > -1 Then
        flCarregarListView
    End If
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboGrupoUsuario_Click", Me.Caption

End Sub

'' Executa a operação selecionada para cada um dos itens da tela
Private Sub flProcessaSelecao(ByRef pobjLvw As MSComctlLib.ListView, _
                              ByVal pblnSelected As Boolean, _
                              ByVal peTipoOperacao As enumTipoOperacao, _
                     Optional ByVal penumTipoAcessoMensagemSPB As enumTipoAcessoMensagemSPB)

Dim objCol                                  As VBA.Collection
Dim lstItem                                 As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    Set objCol = New VBA.Collection
    
    For Each lstItem In pobjLvw.ListItems
        If Not pblnSelected Or lstItem.Selected Then
            objCol.Add lstItem, lstItem.Key
        End If
    Next lstItem
    
    For Each lstItem In objCol
        flMoveItemLvw peTipoOperacao, lstItem.Key, penumTipoAcessoMensagemSPB
    Next lstItem
    
    Set objCol = Nothing

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flProcessaSelecao", 0

End Sub

Private Sub cboTipoBackOffice_Click()

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    If cboGrupoUsuario.ListIndex > -1 And cboTipoBackOffice.ListIndex > -1 Then
        flCarregarListView
    End If
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboTipoBackOffice_Click", Me.Caption

End Sub

Private Sub cmdAdicionar_Click(Index As Integer)

On Error GoTo ErrorHandler

    fgCursor True
    flProcessaSelecao lvwMsg, True, enumTipoOperacao.Incluir, Index
    fgCursor

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cmdAdicionar_Click", Me.Caption

End Sub

Private Sub cmdAdicionarTodas_Click(Index As Integer)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    flProcessaSelecao lvwMsg, False, enumTipoOperacao.Incluir, Index
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name & " - cmdAdicionarTodas_Click", Me.Caption)

End Sub

Private Sub cmdRemover_Click()

On Error GoTo ErrorHandler
    
    fgCursor True
    flProcessaSelecao lvwMsgAutorizadas, True, enumTipoOperacao.Excluir
    fgCursor

Exit Sub
ErrorHandler:
    
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cmdRemover_Click", Me.Caption
    
End Sub

Private Sub cmdRemoverTodas_Click()

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    flProcessaSelecao lvwMsgAutorizadas, False, enumTipoOperacao.Excluir
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name & " - cmdRemoverTodas_Click", Me.Caption)

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Call fgCursor(True)

    Call fgCenterMe(Me)
    Set Me.Icon = mdiLQS.Icon
    DoEvents
    optConsultar.value = True
    flFormataListView
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Call flInicializar
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

'' Preenche os listviews com as mensagems as quais o grupo tem acesso ou não
Private Sub flCarregarListView()

Dim objNodeElement                          As MSXML2.IXMLDOMElement
Dim strXPath                                As String

On Error GoTo ErrorHandler

    flObterGrupoUsuaMensSPB

    lvwMsg.ListItems.Clear
    lvwMsgAutorizadas.ListItems.Clear

    For Each objNodeElement In xmlMapaNavegacao.selectNodes("//Grupo_Dados/Repeat_SistemaMensagem/*")
    
        strXPath = "/Repeat_SistemaMensagem" & _
                   "/Grupo_SistemaMensagem[CO_MESG_SPB='" & objNodeElement.selectSingleNode("CO_MESG").Text & "  ']"
                   
        If Not xmlLerTodos.xml = Empty Then
            If xmlLerTodos.selectSingleNode(strXPath) Is Nothing Then
               flAdicionarItemLvw lvwMsg, objNodeElement
            Else
               flAdicionarItemLvw lvwMsgAutorizadas, objNodeElement, CLng("0" & xmlLerTodos.selectSingleNode(strXPath & "/TP_ACES").Text)
            End If
        Else
            flAdicionarItemLvw lvwMsg, objNodeElement
        End If
        
    Next objNodeElement
    
    Exit Sub

ErrorHandler:
   
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarListView", 0

End Sub

'' Move um tipo de mensagem de uma listagem para outra
Private Sub flMoveItemLvw(ByVal pOperacao As enumTipoOperacao, _
                          ByVal strKey As String, _
                          ByVal penumTipoAcessoMensagemSPB As enumTipoAcessoMensagemSPB)

Dim lstItemOrigem                           As MSComctlLib.ListItem
Dim eTag                                    As enumTipoOperacao

On Error GoTo ErrorHandler

    Select Case pOperacao
        Case enumTipoOperacao.Incluir
            Set lstItemOrigem = lvwMsg.ListItems(strKey)
            If lstItemOrigem.Tag = enumTipoOperacao.None Then
                eTag = enumTipoOperacao.Incluir
            Else
                eTag = enumTipoOperacao.None
            End If
            Call flCopiarItemPara(lstItemOrigem, lvwMsgAutorizadas, eTag, penumTipoAcessoMensagemSPB)
            lvwMsg.ListItems.Remove strKey
            
        Case enumTipoOperacao.Excluir
            Set lstItemOrigem = lvwMsgAutorizadas.ListItems(strKey)
            If lstItemOrigem.Tag = enumTipoOperacao.None Or lstItemOrigem.Tag = enumTipoOperacao.Alterar Then
                eTag = enumTipoOperacao.Excluir
            Else
                eTag = enumTipoOperacao.None
            End If
            Call flCopiarItemPara(lstItemOrigem, lvwMsg, eTag, penumTipoAcessoMensagemSPB)
            lvwMsgAutorizadas.ListItems.Remove strKey
    End Select
    
Exit Sub
ErrorHandler:
    
    Set lstItemOrigem = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flMoveItemLvw", 0

End Sub

'' Chamada por flMoveItemLvw para copiar um item de um listview para outro antes
'' de remove-lo
Private Sub flCopiarItemPara(ByRef lstItemOrigem As MSComctlLib.ListItem, _
                             ByRef objLvwDestino As MSComctlLib.ListView, _
                             ByVal eTag As enumTipoOperacao, _
                             ByVal penumTipoAcessoMensagemSPB As enumTipoAcessoMensagemSPB)
    
Dim lstItemDestino                          As MSComctlLib.ListItem
Dim eTipoAcesso                             As enumTipoAcessoMensagemSPB

On Error GoTo ErrorHandler

    Set lstItemDestino = objLvwDestino.ListItems.Add(, lstItemOrigem.Key)
    
    lstItemDestino.Text = lstItemOrigem.Text
    lstItemDestino.SubItems(1) = lstItemOrigem.SubItems(1)
    If objLvwDestino Is lvwMsgAutorizadas Then
        lstItemDestino.SubItems(2) = IIf(penumTipoAcessoMensagemSPB = enumTipoAcessoMensagemSPB.Consultar, TIPOACESSO_CONSULTAR, TIPOACESSO_ENVIAR)
        'lstItemDestino.SubItems(2) = IIf(optConsultar.Value, TIPOACESSO_CONSULTAR, TIPOACESSO_ENVIAR)
        
        'Se o acesso é none, o item já existe, verificar se o acesso deve ser alterado
        If eTag = enumTipoOperacao.None Then
            eTipoAcesso = xmlLerTodos.selectSingleNode("/Repeat_SistemaMensagem/Grupo_SistemaMensagem[CO_MESG_SPB='" & Mid$(lstItemOrigem.Key, 2) & "  ']/TP_ACES").Text
            If lstItemDestino.SubItems(2) <> IIf(eTipoAcesso = enumTipoAcessoMensagemSPB.Consultar, TIPOACESSO_CONSULTAR, TIPOACESSO_ENVIAR) Then
                eTag = enumTipoOperacao.Alterar
            End If
        Else
               eTag = enumTipoOperacao.Incluir
        End If
    End If
    lstItemDestino.SmallIcon = lstItemOrigem.SmallIcon
    lstItemDestino.Tag = eTag

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flCopiarItemPara", 0
    
End Sub

'' Cria um novo listitem para um tipo de mensagem no momento de preenchimento da
'' tela
Private Sub flAdicionarItemLvw(ByRef objLvw As MSComctlLib.ListView, _
                               ByRef objNodeElement As MSXML2.IXMLDOMElement, _
                      Optional ByVal eTipoAcesso As enumTipoAcessoMensagemSPB)

Dim lstItem                                 As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    Set lstItem = objLvw.ListItems.Add(, "K" & objNodeElement.selectSingleNode("CO_MESG").Text)
    lstItem.Text = objNodeElement.selectSingleNode("CO_MESG").Text
    lstItem.SubItems(1) = objNodeElement.selectSingleNode("NO_MESG").Text
    If objLvw Is lvwMsgAutorizadas Then
        lstItem.SubItems(2) = IIf(eTipoAcesso = enumTipoAcessoMensagemSPB.Consultar, TIPOACESSO_CONSULTAR, TIPOACESSO_ENVIAR)
    End If
    lstItem.Tag = enumTipoOperacao.None
    lstItem.SmallIcon = 8

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flAdicionarItemLvw", 0

End Sub

'' Obtém as mensagens as quais um grupo de usuário tem acesso através da camada de
'' controle de casos de uso MIU.
Private Sub flObterGrupoUsuaMensSPB()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strGrupoUsuarMensSPB    As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuarioMensSPB")
        .selectSingleNode("@Operacao").Text = "LerTodos"
        .selectSingleNode("CO_GRUP_USUA").Text = fgObterCodigoCombo(cboGrupoUsuario.Text)
        .selectSingleNode("TP_BKOF").Text = fgObterCodigoCombo(cboTipoBackOffice.Text)
    End With
    
    Set xmlLerTodos = Nothing
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strGrupoUsuarMensSPB = objMIU.Executar(xmlMapaNavegacao.selectSingleNode("//Grupo_GrupoUsuarioMensSPB").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If Len(strGrupoUsuarMensSPB) > 0 Then
        xmlLerTodos.loadXML strGrupoUsuarMensSPB
    End If

Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flObterGrupoUsuaMensSPB", 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmAssocGrupoUsuarioMensSPB = Nothing

End Sub

Private Sub lvwMsg_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lvwMsg.Sorted = True
    lvwMsg.SortKey = ColumnHeader.Index - 1

    If lvwMsg.SortOrder = lvwAscending Then
        lvwMsg.SortOrder = lvwDescending
    Else
        lvwMsg.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwMsg_ColumnClick"
    
End Sub

Private Sub lvwMsgAutorizadas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lvwMsgAutorizadas.Sorted = True
    lvwMsgAutorizadas.SortKey = ColumnHeader.Index - 1

    If lvwMsgAutorizadas.SortOrder = lvwAscending Then
        lvwMsgAutorizadas.SortOrder = lvwDescending
    Else
        lvwMsgAutorizadas.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwMsgAutorizadas_ColumnClick"

End Sub

Private Sub flMudaAcesso(ByVal eTipoAcesso As enumTipoAcessoMensagemSPB)

Dim lstItem                                 As MSComctlLib.ListItem
Dim strTipoAcesso                           As String

On Error GoTo ErrorHandler

    For Each lstItem In lvwMsgAutorizadas.ListItems
        If lstItem.Selected Then
            strTipoAcesso = IIf(optConsultar.value, TIPOACESSO_CONSULTAR, TIPOACESSO_ENVIAR)
            If lstItem.SubItems(2) <> strTipoAcesso Then
                If lstItem.Tag = enumTipoOperacao.None Then
                    lstItem.Tag = enumTipoOperacao.Alterar
                ElseIf lstItem.Tag = enumTipoOperacao.Alterar Then
                    lstItem.Tag = enumTipoOperacao.None
                End If
                lstItem.SubItems(2) = IIf(eTipoAcesso = enumTipoAcessoMensagemSPB.Consultar, TIPOACESSO_CONSULTAR, TIPOACESSO_ENVIAR)
            End If
        End If
    Next lstItem

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMudaAcesso", 0

End Sub

Private Sub lvwMsgAutorizadas_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    If lvwMsgAutorizadas.ListItems.Count = 0 Then Exit Sub
    
    If Item.SubItems(2) = TIPOACESSO_CONSULTAR And optEnviar.value = True Then
       optConsultar.value = True
    ElseIf Item.SubItems(2) = TIPOACESSO_ENVIAR And optConsultar.value = True Then
       optEnviar.value = True
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwMsgAutorizadas_ItemClick"
    
End Sub

Private Sub optConsultar_Click()

On Error GoTo ErrorHandler

    flMudaAcesso enumTipoAcessoMensagemSPB.Consultar

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - optConsultar_Click"

End Sub

Private Sub optEnviar_Click()

On Error GoTo ErrorHandler

    flMudaAcesso enumTipoAcessoMensagemSPB.Enviar

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - optEnviar_Click"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True
    
    Select Case Button.Key
    Case "Salvar"
        flSalvar
    Case "Sair"
        fgCursor False
        Unload Me
    End Select
    
    fgCursor
Exit Sub
ErrorHandler:
    
    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbCadastro_ButtonClick", Me.Caption

End Sub
