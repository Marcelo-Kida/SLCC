VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInclusaoRegra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Mensagem Disponíveis"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   2610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInclusaoRegra.frx":0000
            Key             =   "Evento"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   5085
      TabIndex        =   3
      Top             =   4050
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   6300
      TabIndex        =   2
      Top             =   4050
      Width           =   1140
   End
   Begin MSComctlLib.ListView lstTipoMensagem 
      Height          =   3435
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo Formato Saída"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selecione um Tipo de Mensagem:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   225
      Width           =   2880
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmInclusaoRegra.frx":031A
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "frmInclusaoRegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição e seleção de tipos de mensagens disponíveis para a configuração de regra de transporte.
Option Explicit

Public CodigoBanco                          As Long
Public SistemaOrigem                        As String

Public Event EventoEscolhido(ByVal pstrTipoMensagem As String, ByVal pstrDescricaoTipoMensagem As String, ByVal plngTipoFormatoMensagemSaida As Long, ByVal plngNaturezaMensagem As Long)

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
On Error GoTo ErrorHandler

    If lstTipoMensagem.SelectedItem Is Nothing Then Exit Sub
    
    RaiseEvent EventoEscolhido(Mid$(lstTipoMensagem.SelectedItem.Key, 8), _
                               lstTipoMensagem.SelectedItem.SubItems(1), _
                               CLng(Mid$(lstTipoMensagem.SelectedItem.Key, 4, 4)), _
                               lstTipoMensagem.SelectedItem.Tag)
    
    Unload Me

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - cmdOK_Click"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Me.Icon = mdiBUS.Icon
    
    cmdOk.Enabled = False
    flCarregarlistEvento
    
    fgCursor
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmInclusãoRegra - frmLoad")

End Sub

'Carregar listview com tipos de mensagem disponíveis.
Private Sub flCarregarlistEvento()

#If EnableSoap = 1 Then
    Dim objRegraTraducao    As MSSOAPLib30.SoapClient30
#Else
    Dim objRegraTraducao    As A7Miu.clsRegraTransporte
#End If

Dim xmlNode                 As MSXML2.IXMLDOMNode
Dim strLerTodos             As String
Dim xmlRegra                As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    fgCursor True
    
    Set objRegraTraducao = fgCriarObjetoMIU("A7Miu.clsRegraTransporte")
    strLerTodos = objRegraTraducao.ObterEventosDisponiveis(CodigoBanco, SistemaOrigem, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objRegraTraducao = Nothing

    If strLerTodos = "" Then
        fgCursor
        Exit Sub
    End If
    
    Set xmlRegra = CreateObject("MSXML2.DOMDocument.4.0")
    xmlRegra.loadXML (strLerTodos)
    
    lstTipoMensagem.ListItems.Clear
    lstTipoMensagem.HideSelection = False
        
    For Each xmlNode In xmlRegra.selectNodes("//Grupo_Evento")
        With lstTipoMensagem.ListItems.Add(, "EVE" & _
                                             Format(xmlNode.selectSingleNode("TP_FORM_MESG_SAID").Text, "0000") & _
                                             xmlNode.selectSingleNode("TP_MESG").Text, _
                                             xmlNode.selectSingleNode("TP_MESG").Text, , "Evento")
            .SubItems(1) = xmlNode.selectSingleNode("NO_TIPO_MESG").Text
            .SubItems(2) = flTipoSaidaToSTR(CLng(xmlNode.selectSingleNode("TP_FORM_MESG_SAID").Text))
            .Tag = xmlNode.selectSingleNode("TP_NATZ_MESG").Text
        End With
    Next
    
    lstTipoMensagem.SelectedItem.Selected = False
    
    Set xmlRegra = Nothing
    fgCursor
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    Set xmlRegra = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Caption, "flCarregarEvento", 0
    
End Sub

Private Sub lstTipoMensagem_DblClick()
    
On Error GoTo ErrorHandler

    If lstTipoMensagem.SelectedItem Is Nothing Then Exit Sub
    
    Call cmdOK_Click

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - lstTipoMensagem_DblClick"

End Sub

Private Sub lstTipoMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

    cmdOk.Enabled = True

End Sub

'Converter o domínio numérico de tipo de saida para literais.
Private Function flTipoSaidaToSTR(lngTipoSaida As Long) As String
    
    Select Case lngTipoSaida
        Case enumTipoSaidaMensagem.SaidaXML
            flTipoSaidaToSTR = "XML"
        Case enumTipoSaidaMensagem.SaidaString
            flTipoSaidaToSTR = "String"
        Case enumTipoSaidaMensagem.SaidaCSV
            flTipoSaidaToSTR = "CSV"
        Case enumTipoSaidaMensagem.SaidaStringXML
            flTipoSaidaToSTR = "String + XML"
        Case enumTipoSaidaMensagem.SaidaCSVXML
            flTipoSaidaToSTR = "CSV + XML"
    End Select

End Function

