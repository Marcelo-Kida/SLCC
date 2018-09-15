VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMensagemRejeitada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log de Mensagens Rejeitadas"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10515
   Begin VB.Frame fraDetalhe 
      Caption         =   "Mensagem Recebida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   60
      TabIndex        =   10
      Top             =   4200
      Width           =   10395
      Begin RichTextLib.RichTextBox rtfErro 
         Height          =   1035
         Left            =   135
         TabIndex        =   12
         Top             =   2640
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   1826
         _Version        =   393217
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMensagemRejeitada.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfMensagem 
         Height          =   2055
         Left            =   135
         TabIndex        =   11
         Top             =   255
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   3625
         _Version        =   393217
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMensagemRejeitada.frx":0084
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo da Rejeição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   13
         Top             =   2400
         Width           =   1785
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   45
      TabIndex        =   3
      Top             =   60
      Width           =   10380
      Begin MSComCtl2.DTPicker dtpDataRecebimento 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59965441
         CurrentDate     =   37827
      End
      Begin MSComctlLib.Toolbar tlbFiltro 
         Height          =   330
         Left            =   4860
         TabIndex        =   4
         Top             =   450
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Atualizar"
               Key             =   "Filtro"
               Object.ToolTipText     =   "Aplicar filtro"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpHoraDe 
         Height          =   315
         Left            =   2010
         TabIndex        =   1
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59965442
         CurrentDate     =   37827
      End
      Begin MSComCtl2.DTPicker dtpHoraAte 
         Height          =   315
         Left            =   3495
         TabIndex        =   2
         Top             =   450
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59965442
         CurrentDate     =   37827.9583333333
      End
      Begin VB.Label lblGeral 
         AutoSize        =   -1  'True
         Caption         =   "Data Recebimento"
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
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label lblGeral 
         AutoSize        =   -1  'True
         Caption         =   "Horário"
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
         Index           =   6
         Left            =   2010
         TabIndex        =   6
         Top             =   210
         Width           =   630
      End
      Begin VB.Label lblGeral 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Index           =   7
         Left            =   3315
         TabIndex        =   5
         Top             =   510
         Width           =   90
      End
   End
   Begin MSComctlLib.ListView lstRecebimento 
      Height          =   3045
      Left            =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   5371
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlIcons"
      SmallIcons      =   "imlIcons"
      ColHdrIcons     =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Identificador Mensagem"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ocorrência"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data/Hora Recebimento"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data/Hora Put na Fila"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Aplicativo Put Na Fila"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1485
      Top             =   7950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0106
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0218
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0532
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0884
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0996
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0CB0
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":0FCA
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":12E4
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":15FE
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":1950
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":1CA2
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMensagemRejeitada.frx":1FF4
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandosForm 
      Height          =   330
      Left            =   9675
      TabIndex        =   9
      Top             =   7980
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   582
      ButtonWidth     =   1191
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Fechar formulário"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMensagemRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição de mensagens rejeitadas pelo sistema A7.
Option Explicit
Dim blnBaseHistorica                        As Boolean

'Carregar ListView com mensagens rejeitadas de acordo com o filtro informado.
Private Sub flCarregarListView()

#If EnableSoap = 1 Then
    Dim objMensagemRejeitada    As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagemRejeitada    As A7Miu.clsMensagemRejeitada
#End If

Dim objDOMNode                  As MSXML2.IXMLDOMNode
Dim objEventos                  As MSXML2.DOMDocument40
Dim strMensagem                 As String
Dim strDtHrDe                   As String
Dim strDtHrAte                  As String
Dim objListItem                 As ListItem
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler
        
    strDtHrDe = Format(Me.dtpDataRecebimento.Value, gstrMascaraDataXml) & _
                       Format(dtpHoraDe.Hour, "00") & _
                       Format(dtpHoraDe.Minute, "00") & _
                       Format(dtpHoraDe.Second, "00")
                       
    strDtHrAte = Format(Me.dtpDataRecebimento.Value, gstrMascaraDataXml) & _
                       Format(dtpHoraAte.Hour, "00") & _
                       Format(dtpHoraAte.Minute, "00") & _
                       Format(dtpHoraAte.Second, "00")
        
    Set objMensagemRejeitada = fgCriarObjetoMIU("A7Miu.clsMensagemRejeitada")
    Set objEventos = CreateObject("MSXML2.DOMDocument.4.0")
        
    Me.lstRecebimento.ListItems.Clear
    lstRecebimento.HideSelection = False
    Call fgCursor(True)

    strMensagem = objMensagemRejeitada.LerTodos(strDtHrDe, strDtHrAte, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem <> "" Then
        If Not objEventos.loadXML(strMensagem) Then
            Call fgErroLoadXML(objEventos, App.EXEName, "frmMensagemRejeitada", "flCarregarListView")
        End If
        
        If Not objEventos.selectSingleNode("//OWNER") Is Nothing Then
            blnBaseHistorica = (objEventos.selectSingleNode("//OWNER").Text = "A7HIST")
        End If
        
        For Each objDOMNode In objEventos.documentElement.selectNodes("//Repeat_MensagemRejeitada/*")
            
            Set objListItem = lstRecebimento.ListItems.Add(, "K" & Format(objDOMNode.selectSingleNode("CO_TEXT_XML").Text, "000000000"), objDOMNode.selectSingleNode("CO_MESG_MQSE").Text)
            
            objListItem.Tag = objDOMNode.selectSingleNode("TX_DTLH_OCOR_ERRO").Text
            
            objListItem.SubItems(1) = objDOMNode.selectSingleNode("DE_OCOR_MESG").Text
            objListItem.SubItems(2) = Format(fgDtHrStr_To_DateTime(objDOMNode.selectSingleNode("DH_RECB_MESG").Text), gstrMascaraDataHoraDtp)
            objListItem.SubItems(3) = Format(fgDtHrStr_To_DateTime(objDOMNode.selectSingleNode("DH_ENTR_FILA_MQSE").Text), gstrMascaraDataHoraDtp)
            objListItem.SubItems(4) = objDOMNode.selectSingleNode("NO_ARQU_ENTR_FILA_MQSE").Text
            
        Next
    Else
        
        rtfMensagem.Text = ""
        rtfErro.Text = ""
        
    End If
    
    Set objMensagemRejeitada = Nothing
    Set objEventos = Nothing
    Set objDOMNode = Nothing
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:

    Set objMensagemRejeitada = Nothing
    Set objEventos = Nothing
    Set objDOMNode = Nothing

    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmMensagemRejeitada - flCarregarListView")
    
End Sub

Private Sub dtpHoraAte_CallbackKeyDown(ByVal KeyCode As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal CallbackField As String, _
                                       CallbackDate As Date)

    MsgBox CallbackDate

End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler
        
    Call fgCenterMe(Me)
    Me.Icon = mdiBUS.Icon
    Me.Show
    DoEvents
    
    Me.dtpDataRecebimento.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    Me.dtpHoraDe.Value = CDate("00:00:00")
    Me.dtpHoraAte.Value = Time
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    mdiBUS.uctLogErros.MostrarErros Err, "frmMensagemRejeitada - Form_Load"

End Sub

Private Sub lstRecebimento_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lstRecebimento, ColumnHeader.Index)
    
    Exit Sub
ErrorHandler:

    mdiBUS.uctLogErros.MostrarErros Err, "frmMensagemRejeitada - lstRecebimento_ColumnClick"
    
End Sub

Private Sub lstRecebimento_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
#If EnableSoap = 1 Then
    Dim objMensagemRejeitada    As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagemRejeitada    As A7Miu.clsMensagemRejeitada
#End If
    
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant
    
On Error GoTo ErrorHandler
    
    fgCursor True
    
    If Me.lstRecebimento.SelectedItem Is Nothing Then Exit Sub
    
    Set objMensagemRejeitada = fgCriarObjetoMIU("A7Miu.clsMensagemRejeitada")
    
    rtfMensagem.Text = ""
    
    If blnBaseHistorica Then
        rtfMensagem.Text = objMensagemRejeitada.ObterMensagem(CLng(Mid(Item.Key, 2)) * -1, vntCodErro, vntMensagemErro)
    Else
        rtfMensagem.Text = objMensagemRejeitada.ObterMensagem(CLng(Mid(Item.Key, 2)), vntCodErro, vntMensagemErro)
    End If
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    rtfErro.Text = ""
    rtfErro.Text = Item.Tag
        
    Set objMensagemRejeitada = Nothing
    fgCursor
    Exit Sub
ErrorHandler:

    fgCursor

    Set objMensagemRejeitada = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmMensagemRejeitada - lstRecebimento_ItemClick"
    
End Sub

Private Sub tlbComandosForm_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Unload Me
    
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo ErrorHandler
    fgCursor True
    dtpHoraDe.Refresh
    dtpHoraAte.Refresh
        
    Call flCarregarListView
    fgCursor
    Exit Sub
ErrorHandler:
    fgCursor
    mdiBUS.uctLogErros.MostrarErros Err, "frmMensagemRejeitada - tlbFiltro_ButtonClick"
End Sub
