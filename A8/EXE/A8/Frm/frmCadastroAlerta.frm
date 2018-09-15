VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroAlerta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parametrização de Alertas"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11370
   Begin VB.Frame fraDetalhe 
      Caption         =   "Detalhe"
      Height          =   4515
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   11145
      Begin VB.CheckBox chkIndicadorEnvioEmail 
         Caption         =   "Enviar e-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5445
         TabIndex        =   19
         Top             =   900
         Width           =   1425
      End
      Begin VB.CheckBox chkIndicadorAlertaTela 
         Caption         =   "Exibir em tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   885
         Width           =   1485
      End
      Begin VB.ComboBox cboFatoGeradorAlerta 
         Height          =   315
         ItemData        =   "frmCadastroAlerta.frx":0000
         Left            =   120
         List            =   "frmCadastroAlerta.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   465
         Width           =   5085
      End
      Begin VB.Frame fraTela 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   120
         TabIndex        =   14
         Top             =   885
         Width           =   5085
         Begin MSComctlLib.ListView lstGrupoUsuario 
            Height          =   2775
            Left            =   120
            TabIndex        =   15
            Top             =   510
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descricao"
               Object.Width           =   6350
            EndProperty
         End
         Begin VB.Label lblAlerta 
            AutoSize        =   -1  'True
            Caption         =   "Grupos de usuários para exibição do alerta"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   16
            Top             =   300
            Width           =   3015
         End
      End
      Begin VB.Frame fraEmail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   5310
         TabIndex        =   10
         Top             =   885
         Width           =   5715
         Begin VB.CheckBox chkEnviaMensagem 
            Caption         =   "Enviar mensagem anexada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   3195
            Width           =   2685
         End
         Begin VB.TextBox txtListaEmail 
            Height          =   735
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   465
            Width           =   5385
         End
         Begin VB.TextBox txtAssuntoEmail 
            Height          =   495
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1440
            Width           =   5385
         End
         Begin VB.TextBox txtTextoEmail 
            Height          =   885
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2190
            Width           =   5385
         End
         Begin VB.Label lblAlerta 
            AutoSize        =   -1  'True
            Caption         =   "Assunto e-mail"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   13
            Top             =   1230
            Width           =   1020
         End
         Begin VB.Label lblAlerta 
            AutoSize        =   -1  'True
            Caption         =   "Texto e-mail"
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   12
            Top             =   1980
            Width           =   855
         End
         Begin VB.Label lblAlerta 
            AutoSize        =   -1  'True
            Caption         =   "Lista de endereços de e-mails (separar por ponto e vírgula)"
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   11
            Top             =   255
            Width           =   4155
         End
      End
      Begin VB.Label lblAlerta 
         AutoSize        =   -1  'True
         Caption         =   "Fato Gerador Alerta"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   255
         Width           =   1380
      End
   End
   Begin VB.ComboBox cboTipoBackOffice 
      Height          =   315
      ItemData        =   "frmCadastroAlerta.frx":0011
      Left            =   4980
      List            =   "frmCadastroAlerta.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   4245
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   7320
      TabIndex        =   6
      Top             =   6960
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   120
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":0015
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":0127
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":0A01
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":12DB
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":1BB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":248F
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":2D69
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":3643
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":3F1D
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":4237
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":4551
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":486B
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":4B85
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":4E9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":51B9
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":54D3
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":5825
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":5937
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":5C51
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":5F6B
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroAlerta.frx":6285
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstAlerta 
      Height          =   1875
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fato Gerador Alerta"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Exibir em tela"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Enviar e-mail"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Lista de endereços de e-mails"
         Object.Width           =   8644
      EndProperty
   End
   Begin VB.Label lblAlerta 
      AutoSize        =   -1  'True
      Caption         =   "Alertas Cadastrados"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   1410
   End
   Begin VB.Label lblTipoBackOffice 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Back Office"
      Height          =   195
      Left            =   3660
      TabIndex        =   8
      Top             =   180
      Width           =   1200
   End
End
Attribute VB_Name = "frmCadastroAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 10:54:26
'-------------------------------------------------
'' Objeto responsavel pelo cadastramento dos alertas emitidos pelo sistema,
'' através da camada de controle de caso de uso MIU.
''
'' Classes de destino consideradas especificamente
''    A8MIU.clsFatoGeraAlerBkOffice
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40

Private strKeyItemSelected                  As String

Private strOperacao                         As String
Private strTipoBackOffice                   As String
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const strFuncionalidade             As String = "frmCadastroAlerta"

'' Reposiciona as seleções do usuário quando o objeto é atualizado
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lstAlerta.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstAlerta.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstAlerta_ItemClick objListItem
           lstAlerta.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimparDetalhe
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

'' Carrega os tipos de alertas existentes e preenche o listview com os mesmos,
'' através da classe controladora de caso de uso MIU, método  A8MIU.
'' clsFatoGeraAlerBkOffice.LerTodos
Private Sub flCarregarAlertasListview()

#If EnableSoap = 1 Then
    Dim objFatoGeraAlerBkOffice     As MSSOAPLib30.SoapClient30
#Else
    Dim objFatoGeraAlerBkOffice     As A8MIU.clsFatoGeraAlerBkOffice
#End If

Dim xmlLerTodos                     As MSXML2.DOMDocument40
Dim xmlDomNode                      As IXMLDOMNode

Dim objListItem                     As ListItem
Dim vntCodErro                      As Variant
Dim vntMensagemErro                 As Variant

On Error GoTo ErrorHandler

    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    Set objFatoGeraAlerBkOffice = fgCriarObjetoMIU("A8MIU.clsFatoGeraAlerBkOffice")
    
    Me.lstAlerta.ListItems.Clear
    
    Call xmlLerTodos.loadXML( _
         objFatoGeraAlerBkOffice.LerTodos(CInt(strTipoBackOffice), vntCodErro, vntMensagemErro))

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If xmlLerTodos.xml <> vbNullString Then
        For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_FatoGeraAlerBkOffice/*")
            Set objListItem = lstAlerta.ListItems.Add(, "K" & xmlDomNode.selectSingleNode("CO_FATO_GERA_ALER").Text & _
                                                        "K" & strTipoBackOffice, _
                                                        xmlDomNode.selectSingleNode("DE_FATO_GERA_ALER").Text)
            
            Select Case xmlDomNode.selectSingleNode("IN_EMIS_ALER_TELA").Text
            Case enumIndicadorSimNao.sim
                objListItem.SubItems(1) = "Sim"
            Case enumIndicadorSimNao.nao
                objListItem.SubItems(1) = "Não"
            End Select
            
            Select Case xmlDomNode.selectSingleNode("IN_ENVI_EMAIL").Text
            Case enumIndicadorSimNao.sim
                objListItem.SubItems(2) = "Sim"
            Case enumIndicadorSimNao.nao
                objListItem.SubItems(2) = "Não"
            End Select
        
            objListItem.SubItems(3) = xmlDomNode.selectSingleNode("DE_EMAIL_ENVI_ALER").Text
        Next
    End If

    Set objFatoGeraAlerBkOffice = Nothing
    Set xmlLerTodos = Nothing
    
    Call flLimparDetalhe
    
    Exit Sub

ErrorHandler:

    Set objFatoGeraAlerBkOffice = Nothing
    Set xmlLerTodos = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarAlertasListview", 0

End Sub

'' Carrega os detalhes dos grupos de usuários e preenche o listview com os mesmos,
'' através da classe controladora de caso de uso MIU, metodo   A8MIU.
'' clsFatoGeraAlerBkOffice.LerGruposUsuariosAlerta
Private Sub flCarregarGruposUsuariosDetalhe()

#If EnableSoap = 1 Then
    Dim objFatoGeradorAlerta        As MSSOAPLib30.SoapClient30
#Else
    Dim objFatoGeradorAlerta        As A8MIU.clsFatoGeraAlerBkOffice
#End If

Dim xmlLerTodos                     As MSXML2.DOMDocument40
Dim xmlDomNode                      As MSXML2.IXMLDOMNode
Dim vntCodErro                      As Variant
Dim vntMensagemErro                 As Variant

On Error GoTo ErrorHandler

    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    Set objFatoGeradorAlerta = fgCriarObjetoMIU("A8MIU.clsFatoGeraAlerBkOffice")
    
    Call xmlLerTodos.loadXML(objFatoGeradorAlerta.LerGruposUsuariosAlerta(CInt(fgObterCodigoCombo(Me.cboFatoGeradorAlerta.Text)), _
                                                                          CInt(strTipoBackOffice), _
                                                                          vntCodErro, _
                                                                          vntMensagemErro))
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
                                                      
    If lstGrupoUsuario.ListItems.Count > 0 Then
        For Each xmlDomNode In xmlLerTodos.selectNodes("/Repeat_GrupoUsuario/*")
            With xmlDomNode
                If fgExisteItemLvw(Me.lstGrupoUsuario, "K" & .selectSingleNode("CO_GRUP_USUA").Text) Then
                   lstGrupoUsuario.ListItems("K" & .selectSingleNode("CO_GRUP_USUA").Text).Checked = True
                End If
            End With
        Next
    End If
    
    Set xmlLerTodos = Nothing
    Set objFatoGeradorAlerta = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    Set xmlLerTodos = Nothing
    Set objFatoGeradorAlerta = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarGruposUsuariosDetalhe", 0
    
End Sub

'' Carrega os grupos de usuários e preenche o listview com os mesmos, através da
'' classe controladora de caso de uso MIU, método  A8MIU.clsMiu.ObterMapaNavegacao
Private Sub flCarregarGruposUsuariosListView()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim xmlLerTodos            As MSXML2.DOMDocument40
Dim xmlDomNode             As MSXML2.IXMLDOMNode
Dim objListItem            As MSComctlLib.ListItem
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlLerTodos.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, "frmGrupoUsuario", vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlLerTodos, App.EXEName, Me.Name, "flCarregarGruposUsuariosListView")
    End If
    
    xmlLerTodos.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/TP_VIGE").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/TP_SEGR").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/@Operacao").Text = "LerTodos"
    Call xmlLerTodos.loadXML(objMIU.Executar(xmlLerTodos.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario").xml, _
                                             vntCodErro, _
                                             vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    lstGrupoUsuario.ListItems.Clear
    
    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_GrupoUsuario/*")
        With xmlDomNode
            Set objListItem = lstGrupoUsuario.ListItems.Add(, "K" & .selectSingleNode("CO_GRUP_USUA").Text, .selectSingleNode("CO_GRUP_USUA").Text)
            objListItem.SubItems(1) = .selectSingleNode("NO_GRUP_USUA").Text
        End With
    Next

    Set xmlLerTodos = Nothing
    Exit Sub
    
ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarGruposUsuariosListView", 0

End Sub

'' Define o tamanho máximo dos campos da tela que podem ser preenchidos pelo
'' usuário
Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement
        txtAssuntoEmail.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_FatoGeraAlerBkOffice/TX_ASSU_EMAIL_ENVI_ALER/@Tamanho").Text
        txtListaEmail.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_FatoGeraAlerBkOffice/DE_EMAIL_ENVI_ALER/@Tamanho").Text
        txtTextoEmail.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_FatoGeraAlerBkOffice/TX_EMAIL_ENVI_ALER/@Tamanho").Text
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0

End Sub

'' Carrega as propriedades necessárias a interface frmCadastroAlerta, através da
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
    
    Call fgCursor(True)
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    Call fgCarregarCombos(Me.cboTipoBackOffice, xmlMapaNavegacao, "TipoBackOffice", "TP_BKOF", "DE_BKOF")
    
    If cboTipoBackOffice.ListCount > 0 Then
        cboTipoBackOffice.ListIndex = 0
    End If
    
    Call fgCarregarCombos(Me.cboFatoGeradorAlerta, xmlMapaNavegacao, "FatoGeradorAlerta", "CO_FATO_GERA_ALER", "DE_FATO_GERA_ALER")
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlLer.loadXML(xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_FatoGeraAlerBkOffice").xml)
    
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

'' Comverte os parametros da interface para um documento XML
Private Sub flInterfaceToXml(ByRef xmlEnvio As MSXML2.DOMDocument40, ByVal strOperacao As String)

On Error GoTo ErrorHandler

    With xmlEnvio.documentElement
        .selectSingleNode("@Operacao").Text = strOperacao
        
        .selectSingleNode("CO_FATO_GERA_ALER").Text = fgObterCodigoCombo(Me.cboFatoGeradorAlerta.Text)
        .selectSingleNode("TP_BKOF").Text = strTipoBackOffice
        .selectSingleNode("DE_EMAIL_ENVI_ALER").Text = fgLimpaCaracterInvalido(txtListaEmail.Text)
        .selectSingleNode("TX_ASSU_EMAIL_ENVI_ALER").Text = fgLimpaCaracterInvalido(txtAssuntoEmail.Text)
        .selectSingleNode("TX_EMAIL_ENVI_ALER").Text = fgLimpaCaracterInvalido(txtTextoEmail.Text)
        
        .selectSingleNode("IN_ENVI_EMAIL_ANEX").Text = IIf(chkEnviaMensagem.value = vbChecked, _
                                                            enumIndicadorSimNao.sim, _
                                                            enumIndicadorSimNao.nao)
        
        .selectSingleNode("IN_ENVI_EMAIL").Text = IIf(chkIndicadorEnvioEmail.value = vbChecked, _
                                                            enumIndicadorSimNao.sim, _
                                                            enumIndicadorSimNao.nao)
        .selectSingleNode("IN_EMIS_ALER_TELA").Text = IIf(chkIndicadorAlertaTela.value = vbChecked, _
                                                            enumIndicadorSimNao.sim, _
                                                            enumIndicadorSimNao.nao)
    End With
    
    Exit Sub

ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Sub

'' Limpa o detalhe de um alerta
Private Sub flLimparDetalhe()

Dim objListItem                             As ListItem

On Error GoTo ErrorHandler

    cboFatoGeradorAlerta.ListIndex = -1
    chkIndicadorAlertaTela.value = vbUnchecked
    chkIndicadorEnvioEmail.value = vbUnchecked
    chkEnviaMensagem.value = vbUnchecked
    txtListaEmail.Text = vbNullString
    txtAssuntoEmail.Text = vbNullString
    txtTextoEmail.Text = vbNullString
    
    For Each objListItem In lstGrupoUsuario.ListItems
        objListItem.Checked = False
    Next

    cboFatoGeradorAlerta.Enabled = True
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    
    Set objListItem = Nothing

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparDetalhe", 0
    
End Sub

'' Salva as alterações efetuadas através da camada controladora de casos de uso
'' MIU, método A8MIU.clsFatoGeraAlerBkOffice.Salvar
Private Function flSalvar() As Boolean
    
#If EnableSoap = 1 Then
    Dim objFatoGeraAlerBkOffice     As MSSOAPLib30.SoapClient30
#Else
    Dim objFatoGeraAlerBkOffice     As A8MIU.clsFatoGeraAlerBkOffice
#End If

Dim xmlSalvar                       As MSXML2.DOMDocument40
Dim xmlUsuarios                     As MSXML2.DOMDocument40
Dim xmlAux                          As MSXML2.DOMDocument40
Dim objListItem                     As ListItem

Dim strRetorno                      As String
Dim strFatoGerador                  As String

Dim vntCodErro                      As Variant
Dim vntMensagemErro                 As Variant

On Error GoTo ErrorHandler

    flSalvar = False

    If strOperacao = gstrOperExcluir Then
        If MsgBox("Confirma a exclusão do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    
    strRetorno = flValidarCampos()
    
    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Function
    End If
    
    Call fgCursor(True)
    
    Set xmlSalvar = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlUsuarios = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlSalvar, "", "Repeat_CadastroAlerta", "")
    
    Call flInterfaceToXml(xmlLer, strOperacao)
    Call fgAppendXML(xmlSalvar, "Repeat_CadastroAlerta", xmlLer.xml)
    
    strFatoGerador = fgObterCodigoCombo(cboFatoGeradorAlerta.Text)
    
    Call fgAppendNode(xmlUsuarios, "", "Repeat_GrupoUsuario", "")
    
    If strOperacao <> gstrOperExcluir Then
        For Each objListItem In lstGrupoUsuario.ListItems
            If objListItem.Checked Then
                Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
                Call fgAppendNode(xmlAux, "", "GrupoUsuario", "")
                Call fgAppendNode(xmlAux, "GrupoUsuario", "TP_BKOF", strTipoBackOffice)
                Call fgAppendNode(xmlAux, "GrupoUsuario", "CO_FATO_GERA_ALER", strFatoGerador)
                Call fgAppendNode(xmlAux, "GrupoUsuario", "CO_GRUP_USUA", objListItem.Text)
                Call fgAppendXML(xmlUsuarios, "Repeat_GrupoUsuario", xmlAux.xml)
                Set xmlAux = Nothing
            End If
        Next
    End If

    Call fgAppendXML(xmlSalvar, "Repeat_CadastroAlerta", xmlUsuarios.xml)
    
    Set objFatoGeraAlerBkOffice = fgCriarObjetoMIU("A8MIU.clsFatoGeraAlerBkOffice")
    Call objFatoGeraAlerBkOffice.Salvar(xmlSalvar.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objFatoGeraAlerBkOffice = Nothing
    Set xmlSalvar = Nothing
    Set xmlUsuarios = Nothing
    
    Call fgCursor(False)
    Call flCarregarAlertasListview
    
    Select Case strOperacao
    Case gstrOperAlterar, gstrOperIncluir
    
        If strOperacao = gstrOperIncluir Then
           strKeyItemSelected = "K" & strFatoGerador & "K" & strTipoBackOffice
        End If
    
        If fgExisteItemLvw(Me.lstAlerta, "K" & strFatoGerador & "K" & strTipoBackOffice) Then
            lstAlerta.ListItems("K" & strFatoGerador & "K" & strTipoBackOffice).Selected = True
            Call lstAlerta_ItemClick(lstAlerta.ListItems("K" & strFatoGerador & "K" & strTipoBackOffice))
        End If
        
    Case gstrOperExcluir
        Call flLimparDetalhe
        strOperacao = gstrOperIncluir
    End Select
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    flSalvar = True
    
    Exit Function

ErrorHandler:
    Call fgCursor(False)
    Set objFatoGeraAlerBkOffice = Nothing
    Set xmlSalvar = Nothing
    Set xmlUsuarios = Nothing
    
    If strOperacao <> gstrOperExcluir Then
       strKeyItemSelected = "K" & strFatoGerador & "K" & strTipoBackOffice
    End If

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, Me.Name, "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString
Private Function flValidarCampos() As String
    
Dim objListItem                             As ListItem
Dim blnGrupoUsuarioSel                      As Boolean

Dim arrEmail()                              As String
Dim lngCont                                 As Integer

On Error GoTo ErrorHandler

    If cboTipoBackOffice.Text = vbNullString Then
        flValidarCampos = "Selecione o Tipo de BackOffice."
        cboTipoBackOffice.SetFocus
        Exit Function
    End If
    
    If cboFatoGeradorAlerta.Text = vbNullString Then
        flValidarCampos = "Selecione o Fato Gerador de Alerta."
        cboFatoGeradorAlerta.SetFocus
        Exit Function
    End If
    
    If strOperacao <> gstrOperExcluir Then
        If chkIndicadorAlertaTela.value = vbChecked Then
            blnGrupoUsuarioSel = False
            For Each objListItem In lstGrupoUsuario.ListItems
                If objListItem.Checked Then
                    blnGrupoUsuarioSel = True
                    Exit For
                End If
            Next
        
            If Not blnGrupoUsuarioSel Then
                flValidarCampos = "Selecione pelo menos um Grupo de Usuário para receber o alerta na tela."
                Exit Function
            End If
        End If
        
        If chkIndicadorEnvioEmail.value = vbChecked Then
            If Trim$(txtListaEmail.Text) = vbNullString Then
                flValidarCampos = "Informe o(s) endereço(s) de e-mail a ser(em) enviado(s)."
                txtListaEmail.SetFocus
                Exit Function
            End If
        
            arrEmail = Split(Replace(txtListaEmail, " ", vbNullString), ";", , vbBinaryCompare)
            
            For lngCont = 0 To UBound(arrEmail)
                If Not fgValidarEmail(arrEmail(lngCont)) Then
                    flValidarCampos = "E-mail informado inválido : " & arrEmail(lngCont) & "."
                    txtListaEmail.SetFocus
                    Exit Function
                End If
            Next
                
            If Trim$(txtAssuntoEmail.Text) = vbNullString Then
                flValidarCampos = "Informe o assunto do e-mail a ser enviado."
                txtAssuntoEmail.SetFocus
                Exit Function
            End If
        
            If Trim$(txtTextoEmail.Text) = vbNullString Then
                flValidarCampos = "Informe o texto do e-mail a ser enviado."
                txtTextoEmail.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If strOperacao = gstrOperIncluir Then
        If Me.lstAlerta.ListItems.Count > 0 Then
            If fgExisteItemLvw(Me.lstAlerta, "K" & fgObterCodigoCombo(cboFatoGeradorAlerta.Text) & "K" & strTipoBackOffice) Then
                flValidarCampos = "Este Fato Gerador de Alerta já está cadastrado para o Tipo de BackOffice selecionado."
                cboFatoGeradorAlerta.SetFocus
                Exit Function
            End If
        End If
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'' Preenche a interface de acordo com o documento XML
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objFatoGeraAlerBkOffice     As MSSOAPLib30.SoapClient30
#Else
    Dim objFatoGeraAlerBkOffice     As A8MIU.clsFatoGeraAlerBkOffice
#End If

Dim arrItemKey()                    As String
Dim vntCodErro                      As Variant
Dim vntMensagemErro                 As Variant

On Error GoTo ErrorHandler

    arrItemKey = Split(lstAlerta.SelectedItem.Key, "K")
    
    xmlLer.documentElement.selectSingleNode("//@Operacao").Text = gstrOperLer
    
    Set objFatoGeraAlerBkOffice = fgCriarObjetoMIU("A8MIU.clsFatoGeraAlerBkOffice")
    Call xmlLer.loadXML(objFatoGeraAlerBkOffice.Ler(CInt(arrItemKey(1)), CInt(arrItemKey(2)), vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objFatoGeraAlerBkOffice = Nothing
   
    If xmlLer.xml <> vbNullString Then
        With xmlLer.documentElement
        
            If .selectSingleNode("IN_EMIS_ALER_TELA").Text <> vbNullString Then
                chkIndicadorAlertaTela.value = IIf(.selectSingleNode("IN_EMIS_ALER_TELA").Text = enumIndicadorSimNao.sim, _
                                                            vbChecked, _
                                                            vbUnchecked)
            End If
            
            If .selectSingleNode("IN_ENVI_EMAIL").Text <> vbNullString Then
                chkIndicadorEnvioEmail.value = IIf(.selectSingleNode("IN_ENVI_EMAIL").Text = enumIndicadorSimNao.sim, _
                                                            vbChecked, _
                                                            vbUnchecked)
            End If
            
            If .selectSingleNode("IN_ENVI_EMAIL_ANEX").Text <> vbNullString Then
                chkEnviaMensagem.value = IIf(.selectSingleNode("IN_ENVI_EMAIL_ANEX").Text = enumIndicadorSimNao.sim, _
                                                            vbChecked, _
                                                            vbUnchecked)
            End If
            
            Call chkIndicadorAlertaTela_Click
            Call chkIndicadorEnvioEmail_Click
            
            txtListaEmail.Text = .selectSingleNode("DE_EMAIL_ENVI_ALER").Text
            txtAssuntoEmail.Text = .selectSingleNode("TX_ASSU_EMAIL_ENVI_ALER").Text
            txtTextoEmail.Text = .selectSingleNode("TX_EMAIL_ENVI_ALER").Text
            
            Call fgSearchItemCombo(Me.cboFatoGeradorAlerta, , .selectSingleNode("CO_FATO_GERA_ALER").Text)
        End With
    End If
    
    Exit Sub

ErrorHandler:
    
    Set objFatoGeraAlerBkOffice = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0
    
End Sub

Private Sub cboTipoBackOffice_Click()

On Error GoTo ErrorHandler

    If cboTipoBackOffice.Text <> vbNullString Then
        strOperacao = gstrOperIncluir
        strTipoBackOffice = fgObterCodigoCombo(cboTipoBackOffice.Text)
        Call flCarregarAlertasListview
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboTipoBackOffice_Click"

End Sub

Private Sub chkIndicadorAlertaTela_Click()
    fraTela.Enabled = IIf(chkIndicadorAlertaTela.value = vbChecked, True, False)
End Sub

Private Sub chkIndicadorEnvioEmail_Click()
    fraEmail.Enabled = IIf(chkIndicadorEnvioEmail.value = vbChecked, True, False)
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

    Call fgCursor(True)

    Call fgCenterMe(Me)
    Set Me.Icon = mdiLQS.Icon
    'Alteração Temporária deverá ser retirado o TipoBackOffice desta Tela
    lblTipoBackOffice.Visible = False
    cboTipoBackOffice.Visible = False
    
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Call flInicializar
    Call flDefinirTamanhoMaximoCampos
    Call flCarregarGruposUsuariosListView
    Call fgCursor(False)

    strOperacao = gstrOperIncluir
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    Set xmlLer = Nothing
    Set frmCadastroAlerta = Nothing

End Sub

Private Sub lstAlerta_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    fgCursor True
    strOperacao = gstrOperAlterar
    
    strKeyItemSelected = Item.Key
    
    Call flLimparDetalhe
    Call flXmlToInterface
    Call flCarregarGruposUsuariosDetalhe
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    cboFatoGeradorAlerta.Enabled = False
    fgCursor

    Exit Sub
    
ErrorHandler:
    
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
End Sub

Private Sub lstGrupoUsuario_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then Item.Selected = True
End Sub

'' Salva ou exclui o registro selecionado ,limpa os campos ,para permitir a
'' entrada de um novo registro, ou fecha a tela
Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Select Case Button.Key
    Case "Limpar"
        strOperacao = gstrOperIncluir
        Call flLimparDetalhe
    Case gstrOperExcluir
        strOperacao = gstrOperExcluir
        Call flSalvar
    Case gstrSalvar
        If flSalvar Then
            If strOperacao = gstrOperAlterar Then
               flPosicionaItemListView
            End If
        End If
    Case gstrSair
        Call fgCursor(False)
        Unload Me
    End Select
    
    Call fgCursor(False)
    
Exit Sub

ErrorHandler:

    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
    Call flCarregarAlertasListview
    
    If strOperacao = gstrOperExcluir Then
        flLimparDetalhe
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If
    
End Sub

