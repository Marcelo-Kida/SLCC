VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReprocessaMensagem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprocessamento de Mensagens"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   Begin A7.ctlMenu ctlMenu1 
      Left            =   3450
      Top             =   6030
      _ExtentX        =   2831
      _ExtentY        =   556
   End
   Begin RichTextLib.RichTextBox rtfMensagem 
      Height          =   1815
      Left            =   75
      TabIndex        =   7
      Top             =   4170
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmReprocessaMensagem.frx":0000
   End
   Begin VB.Frame fraSistema 
      Height          =   615
      Left            =   75
      TabIndex        =   2
      Top             =   -75
      Width           =   12435
      Begin VB.ComboBox cboSistema 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Sistema"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3315
      Left            =   75
      TabIndex        =   0
      Top             =   585
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5847
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   5640
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":0082
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":0194
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":0A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":1348
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":1C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":24FC
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":2DD6
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":36B0
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":3F8A
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":42A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":48D8
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":4BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":4F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":5226
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":5540
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":5892
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":59A4
            Key             =   "Confirmar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":5CBE
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":5FD8
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":62F2
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReprocessaMensagem.frx":660C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   8430
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   582
      ButtonWidth     =   2540
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "Atualizar"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reprocessar"
            Key             =   "Reprocessar"
            Object.ToolTipText     =   "Reprocessar Mensagem"
            ImageKey        =   "AlterarAgendamento"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Arquivar"
            Key             =   "Arquivar"
            ImageKey        =   "Sistema"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                 "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfErro 
      Height          =   1815
      Left            =   75
      TabIndex        =   8
      Top             =   6345
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmReprocessaMensagem.frx":695E
   End
   Begin VB.Label Label3 
      Caption         =   "Mensagem de Erro"
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
      Left            =   75
      TabIndex        =   6
      Top             =   6105
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Mensagem Original"
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
      Left            =   75
      TabIndex        =   5
      Top             =   3945
      Width           =   2055
   End
End
Attribute VB_Name = "frmReprocessaMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição de mensagens paradas em filas de erro e o seu reprocessamento.
Option Explicit

Private Const FILTRO_TODOS                  As Integer = 0
Private Const FILTRO_A6                     As Integer = 1
Private Const FILTRO_A7                     As Integer = 2
Private Const FILTRO_A8                     As Integer = 3

Private Const COL_FILA_ORIGEM               As Integer = 0
Private Const COL_CODIGO_ERRO               As Integer = 1
Private Const COL_FILA_ERRO                 As Integer = 2
Private Const COL_DATA                      As Integer = 3

Private xmlMensagens                        As MSXML2.DOMDocument40

'Carregar combo para filtro de sistema.
Private Sub flCarregaCombo()
    
    With cboSistema
        .Clear
        .AddItem "Todos", FILTRO_TODOS
        .AddItem "A6 - LQS", FILTRO_A6
        .AddItem "A7 - BUS", FILTRO_A7
        .AddItem "A8 - LQS", FILTRO_A8
        .ListIndex = FILTRO_TODOS
    End With

End Sub

'Configurar listview de mensagens rejeitadas
Private Sub flConfiguraListView()

    With lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Fila de Origem / Aplicativo Origem", 3000
        .Add , , "Código do Erro", 3000
        .Add , , "Fila de Erro", 3000
        .Add , , "Data/Hora", 3000
    End With
    
End Sub

Private Sub cboSistema_Click()

On Error GoTo ErrorHandler

    flCarregarLista

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboSistema_Click"

End Sub

Private Sub ctlMenu1_ClickMenu(ByVal Retorno As Long)

    On Error GoTo ErrorHandler

    Select Case Retorno
        Case enumTipoSelecao.MarcarTodas, enumTipoSelecao.DesmarcarTodas
            Call fgMarcarDesmarcarTodas(lvwMensagem, Retorno)
    End Select
    
    Exit Sub
    
ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, Me.Name & " - ctlMenu1_ClickMenu", Me.Caption
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        
        fgCursor True
        
        flCarregarLista
        
        fgCursor
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmMonitoracao - Form_KeyDown")

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCenterMe Me
    
    Me.Icon = mdiBUS.Icon
    
    flConfiguraListView
    flCarregaCombo

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Private Sub lvwMensagem_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    If Split(Item.Key, "|")(2) = "S" Then
        MsgBox "Mensagem não pode ser reprocessada", vbCritical, "A7 - BUS"
    End If
    
Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - lvwMensagem_ItemCheck"

End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    flExibirMensagem Split(Item.Key, "|")(1)

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - lvwMensagem_ItemClick"

End Sub

'Exibir informações da mensagem selecionada.
Private Sub flExibirMensagem(ByVal pstrChave As String)

Dim objDOMNode                              As MSXML2.IXMLDOMNode
Dim strMensagem                             As String
Dim strErro                                 As String

Dim xmlAux                                  As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")

    Set objDOMNode = xmlMensagens.selectSingleNode("//FILA_ERRO[MESSAGE_ID='" & pstrChave & "']")
    
    If xmlAux.loadXML(objDOMNode.selectSingleNode("TX_MESG_ORIG").firstChild.xml) Then
        strMensagem = xmlAux.xml
    Else
        strMensagem = objDOMNode.selectSingleNode("TX_MESG_ORIG").Text
    End If
    
    If xmlAux.loadXML(objDOMNode.selectSingleNode("TX_MESG_ERRO").firstChild.xml) Then
        strErro = xmlAux.xml
    Else
        strErro = objDOMNode.selectSingleNode("TX_MESG_ERRO").Text
    End If
    
    rtfErro.Text = strErro
    
    rtfMensagem.Text = strMensagem

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flExibirMensagem", 0

End Sub

Private Sub lvwMensagem_KeyDown(KeyCode As Integer, Shift As Integer)

    Form_KeyDown KeyCode, 0

End Sub

Private Sub lvwMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        ctlMenu1.ShowMenuMarcarDesmarcar
    End If

    Exit Sub
    
ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, Me.Name & " - lvwMensagem_MouseDown", Me.Caption

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strSelecaoFiltro                        As String
Dim strResultadoConfirmacao                 As String

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
                
        Case "Reprocessar"
            If (MsgBox("Deseja reprocessar as mensagens selecionadas?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes) Then
                flReprocessar Button.Key
            End If
            
            flCarregarLista
            
        Case "Arquivar"
            
            If (MsgBox("Deseja arquivar as mensagens selecionadas?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes) Then
                flReprocessar Button.Key
            End If
            
            flCarregarLista
            
        Case gstrAtualizar
            flCarregarLista
        
        Case gstrSair
            Unload Me
            
    End Select
    
    fgCursor
    
    Exit Sub

ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmReprocessaMensagem - tlbFiltro_ButtonClick", Me.Caption

End Sub

'Carregar listview com as mensagens rejeitadas.
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objFilaErro         As MSSOAPLib30.SoapClient30
#Else
    Dim objFilaErro         As A7Miu.clsFilaErro
#End If

Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strMensagens            As String
Dim objDOMNode              As MSXML2.IXMLDOMNode
Dim lstItem                 As ListItem
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    fgCursor True
    
    rtfMensagem.Text = ""
    rtfErro.Text = ""

    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlDomFiltros, "", "Repeat_Filtros", ""
    fgAppendNode xmlDomFiltros, "Repeat_Filtros", "Repeat_Sistema", ""
    
    If cboSistema.ListIndex = FILTRO_TODOS Then
        fgAppendNode xmlDomFiltros, "Repeat_Sistema", "Grupo_Sistema", "A6"
        fgAppendNode xmlDomFiltros, "Repeat_Sistema", "Grupo_Sistema", "A7"
        fgAppendNode xmlDomFiltros, "Repeat_Sistema", "Grupo_Sistema", "A8"
    Else
        fgAppendNode xmlDomFiltros, "Repeat_Sistema", "Grupo_Sistema", fgObterCodigoCombo(cboSistema)
    End If
    
    Set objFilaErro = fgCriarObjetoMIU("A7MIU.clsFilaErro")
    strMensagens = objFilaErro.LerMensagem(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objFilaErro = Nothing
    
    lvwMensagem.ListItems.Clear
    Set xmlMensagens = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlMensagens.loadXML(strMensagens) Then
        'Preenche o Listview
        For Each objDOMNode In xmlMensagens.documentElement.childNodes
                
            If objDOMNode.selectSingleNode("@MensagemDesconhecida") Is Nothing Then
                Set lstItem = lvwMensagem.ListItems.Add(, "|" & objDOMNode.selectSingleNode("MESSAGE_ID").Text & "|N")
            Else
                Set lstItem = lvwMensagem.ListItems.Add(, "|" & objDOMNode.selectSingleNode("MESSAGE_ID").Text & "|S")
            End If
            lstItem.Text = objDOMNode.selectSingleNode("NO_FILA_ORIG_MQSE").Text
            lstItem.SubItems(COL_CODIGO_ERRO) = flObterDescricaoErro(objDOMNode.selectSingleNode("TX_MESG_ERRO").Text)
            lstItem.SubItems(COL_FILA_ERRO) = objDOMNode.selectSingleNode("NO_FILA_ERRO_MQSE").Text
            lstItem.SubItems(COL_DATA) = fgDtHrXML_To_Interface(objDOMNode.selectSingleNode("DH_MESG_ERRO").Text)
            
        Next objDOMNode
    End If

    fgCursor

    Set xmlDomFiltros = Nothing

Exit Sub
ErrorHandler:

    fgCursor

    Set xmlDomFiltros = Nothing
    Set objFilaErro = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - flCarregarLista"

End Sub

'Obter a descrição do erro de acordo com o código informado.
Private Function flObterDescricaoErro(ByVal pstrErro As String) As String

Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strRetorno                              As String

On Error GoTo ErrorHandler

    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlErro.loadXML(pstrErro) Then
        strRetorno = pstrErro
    Else
        strRetorno = xmlErro.documentElement.selectSingleNode("Grupo_ErrorInfo/Number").Text & " - " & xmlErro.documentElement.selectSingleNode("Grupo_ErrorInfo/Description").Text
    End If

    flObterDescricaoErro = strRetorno

    Set xmlErro = Nothing

Exit Function
ErrorHandler:

    Set xmlErro = Nothing

   'fgRaiseError App.EXEName,TypeName(me),"flObterDescricaoErro",0
   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - flObterDescricaoErro"

End Function

'Reprocessar as mensagens selecionadas.
Private Function flReprocessar(ByVal pstrAcao As String) As Boolean

#If EnableSoap = 1 Then
    Dim objFilaErro         As MSSOAPLib30.SoapClient30
#Else
    Dim objFilaErro         As A7Miu.clsFilaErro
#End If

Dim xmlLoteProcessamento    As MSXML2.DOMDocument40
Dim xmlNode                 As MSXML2.DOMDocument40
Dim lstItem                 As ListItem
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set xmlLoteProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlLoteProcessamento, "", "Repeat_Mensagem", ""
    
    For Each lstItem In lvwMensagem.ListItems
        If lstItem.Checked Then
            
            Set xmlNode = CreateObject("MSXML2.DOMDocument.4.0")
            
            fgAppendNode xmlNode, "", "Grupo_Mensagem", ""
            
            If Split(lstItem.Key, "|")(2) = "S" And pstrAcao = "Reprocessar" Then
                fgAppendAttribute xmlNode, "Grupo_Mensagem", "Acao", ""
            Else
                fgAppendAttribute xmlNode, "Grupo_Mensagem", "Acao", pstrAcao
            End If
            
            fgAppendNode xmlNode, "Grupo_Mensagem", "MESSAGE_ID", Split(lstItem.Key, "|")(1)
            fgAppendNode xmlNode, "Grupo_Mensagem", "NO_FILA_ERRO_MQSE", lstItem.SubItems(COL_FILA_ERRO)
                
            fgAppendXML xmlLoteProcessamento, "Repeat_Mensagem", xmlNode.xml
            Set xmlNode = Nothing
        End If
        
    Next lstItem
    
    Set objFilaErro = fgCriarObjetoMIU("A7MIU.clsFilaErro")
    
    objFilaErro.ReprocessaMensagens xmlLoteProcessamento.xml, vntCodErro, vntMensagemErro
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objFilaErro = Nothing
    Set xmlLoteProcessamento = Nothing

    Exit Function
ErrorHandler:

    Set xmlNode = Nothing
    Set objFilaErro = Nothing
    Set xmlLoteProcessamento = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flReprocessar", 0

End Function

