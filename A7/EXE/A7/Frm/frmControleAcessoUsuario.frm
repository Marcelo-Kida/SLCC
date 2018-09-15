VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControleAcessoUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Controle de Acesso Usuário"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10155
   Begin MSComCtl2.DTPicker dtpDataFiltro 
      Height          =   345
      Left            =   4800
      TabIndex        =   6
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   60227585
      CurrentDate     =   38444
   End
   Begin VB.ComboBox cboSistema 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtUsuario 
      Height          =   345
      Left            =   120
      MaxLength       =   8
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin MSComctlLib.ListView lstControleAcesso 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   10610
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuário"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sistema"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Último Acesso"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFuncoes 
      Height          =   330
      Left            =   135
      TabIndex        =   7
      Top             =   7005
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   582
      ButtonWidth     =   1826
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atualizar"
            Key             =   "Atualizar"
            ImageKey        =   "Atualizar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   7080
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":0112
            Key             =   "Testar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":0564
            Key             =   "Camara"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":09B6
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":0CD0
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":1022
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":1134
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":144E
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":1768
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleAcessoUsuario.frx":1A82
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmControleAcessoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pelo cadastramento e manutenção das regras de transporte do sistema A7.
Option Explicit

Private Enum enumSistemaControleAcesso
    
    A6 = 1
    A7 = 2
    A8 = 3
    A6T = 4
    A7T = 5
    A8T = 6
    A8P = 7

End Enum

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmControleAcessoUsuario"

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True
    
    fgCenterMe Me
    Me.Icon = mdiBUS.Icon
    Me.Show
    
    DoEvents
        
    flInicializar
    flCarregaComboSistema
    
    cboSistema.ListIndex = 0
    dtpDataFiltro.Value = fgDataHoraServidor(DataAux)
    dtpDataFiltro.Value = Null
    
    fgCursor
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmControleAcessoUsuario - Form_Load")

End Sub

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlNode             As MSXML2.IXMLDOMNode
Dim strPropriedades     As String
Dim strLerTodos         As String
Dim xmlLerTodos         As MSXML2.DOMDocument40
Dim dtmDataServidor     As Date
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set objMiu = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.loadXML objMiu.ObterMapaNavegacao(1, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmControleAcessoUsuario", "flInicializar", 0
    
End Sub

Private Sub lstControleAcesso_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lstControleAcesso, ColumnHeader.Index)
    
    Exit Sub
ErrorHandler:

    mdiBUS.uctLogErros.MostrarErros Err, "frmControleAcessoUsuario - lstRecebimento_ColumnClick"

End Sub

Private Sub tlbFuncoes_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Select Case Button.Key
        
        Case "Atualizar"
            
            fgCursor True
            Call flCarregaLista
            fgCursor False
            
        Case "Sair"
            Unload Me
    End Select
    
    Exit Sub

ErrorHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmControleAcessoUsuario - tlbFuncoes_ButtonClick"

End Sub

Private Sub flCarregaLista()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlNode             As MSXML2.IXMLDOMNode
Dim xmlControleAcesso   As MSXML2.DOMDocument40
Dim strControleAcesso   As String
Dim objListItem         As MSComctlLib.ListItem
Dim strSistema          As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
        
    Select Case cboSistema.ListIndex
        Case enumSistemaControleAcesso.A6
            strSistema = "A6"
        Case enumSistemaControleAcesso.A6T
            strSistema = "A6T"
        Case enumSistemaControleAcesso.A7
            strSistema = "A7"
        Case enumSistemaControleAcesso.A7T
            strSistema = "A7T"
        Case enumSistemaControleAcesso.A8
            strSistema = "A8"
        Case enumSistemaControleAcesso.A8T
            strSistema = "A8T"
        Case enumSistemaControleAcesso.A8P
            strSistema = "A8P"
    End Select
    
    xmlMapaNavegacao.selectSingleNode("//Grupo_ControleAcessoSistemaUsuariuo/@Operacao").Text = "LerTodos"
    xmlMapaNavegacao.selectSingleNode("//CO_USUA_ACES").Text = txtUsuario
    xmlMapaNavegacao.selectSingleNode("//SG_SIST").Text = strSistema
    
    If dtpDataFiltro.Value <> Null Or dtpDataFiltro.Value <> "" Then
        xmlMapaNavegacao.selectSingleNode("//DT_ULTI_ACES").Text = fgDt_To_Xml(dtpDataFiltro.Value)
    Else
        xmlMapaNavegacao.selectSingleNode("//DT_ULTI_ACES").Text = ""
    End If
    
    strControleAcesso = objMiu.Executar(xmlMapaNavegacao.selectSingleNode("//Grupo_ControleAcessoSistemaUsuariuo").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lstControleAcesso.ListItems.Clear
            
    If strControleAcesso = vbNullString Then
        Set objMiu = Nothing
        Exit Sub
    End If

    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")

    xmlControleAcesso.loadXML strControleAcesso

    For Each xmlNode In xmlControleAcesso.selectNodes("//Repeat_ControleAcessoUsuario/*")

        Set objListItem = lstControleAcesso.ListItems.Add(, "", xmlNode.selectSingleNode("CO_USUA_ACES").Text)
        
        Select Case Trim$(xmlNode.selectSingleNode("SG_SIST").Text)
            Case "A6": strSistema = "A6 - Sub Reserva"
            Case "A6T": strSistema = "A6 - Trilha Auditoria"
            Case "A7": strSistema = "A7 - BUS"
            Case "A7T": strSistema = "A7 - Trilha Auditoria"
            Case "A8": strSistema = "A8 - LQS"
            Case "A8T": strSistema = "A8 - Trilha Auditoria"
            Case "A8P": strSistema = "A8 - Processamento"
        End Select
        
        objListItem.SubItems(1) = strSistema
        objListItem.SubItems(2) = fgDtXML_To_Interface(xmlNode.selectSingleNode("DT_ULTI_ACES").Text)
        
        If DateDiff("d", fgDtXML_To_Date(xmlNode.selectSingleNode("DT_ULTI_ACES").Text), fgDataHoraServidor(DataHoraAux)) > 30 Then
            
            objListItem.ForeColor = vbRed
            objListItem.ListSubItems.Item(1).ForeColor = vbRed
            objListItem.ListSubItems.Item(2).ForeColor = vbRed
                
            objListItem.Bold = True
            objListItem.ListSubItems.Item(1).Bold = True
            objListItem.ListSubItems.Item(2).Bold = True
        
        End If

    Next

    Set objMiu = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregaLista", 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        
        fgCursor True
        flCarregaLista
        fgCursor
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmControleAcessoUsuario - Form_KeyDown")

End Sub

Private Sub flCarregaComboSistema()

On Error GoTo ErrorHandler
        
    cboSistema.Clear
    
    cboSistema.AddItem "<Todos>"
    
    cboSistema.AddItem "A6 - Sub Reserva"
    cboSistema.AddItem "A7 - BUS"
    cboSistema.AddItem "A8 - LQS"
    cboSistema.AddItem "A6 - Trilha Auditoria"
    cboSistema.AddItem "A7 - Trilha Auditoria"
    cboSistema.AddItem "A8 - Trilha Auditoria"
    cboSistema.AddItem "A8 - Processamento"
   
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmControleAcessoUsuario - Form_KeyDown")

End Sub

