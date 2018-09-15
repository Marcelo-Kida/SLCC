VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacaoArquivoBMF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ferramentas - Importação Arquivo BMF"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7980
   Begin MSComDlg.CommonDialog cdlgArquivoBMF 
      Left            =   7080
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   7815
      Begin VB.CommandButton cmdSelecaoArquivo 
         Caption         =   "..."
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
         Left            =   6930
         TabIndex        =   5
         Top             =   390
         Width           =   480
      End
      Begin VB.TextBox txtNomeArquivo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivos Importados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   7815
      Begin MSComctlLib.ListView lstLogImportacao 
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome Arquivo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data Importação"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuário"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3630
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      ButtonWidth     =   2884
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar Arquivo"
            Key             =   "importar"
            ImageKey        =   "reprocessar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "refresh"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r      "
            Key             =   "sair"
            ImageKey        =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   6960
      Top             =   3240
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
            Picture         =   "frmImportacaoArquivoBMF.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":1286
            Key             =   "reprocessar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoBMF.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmImportacaoArquivoBMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelecaoArquivo_Click()

On Error GoTo ErrorHandler
    
    cdlgArquivoBMF.DialogTitle = "Selecione o arquivo BMF"
    cdlgArquivoBMF.FileName = ""
    cdlgArquivoBMF.Filter = "*.txt"
    cdlgArquivoBMF.Action = 1
        
    If cdlgArquivoBMF.FileTitle <> "" Then
        txtNomeArquivo.Text = cdlgArquivoBMF.FileName
    End If
      
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmImportacaoArquivoBMF - cmdSelecaoArquivo_Click", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    Call flCarregaHistoricoImpoertacao
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmImportacaoArquivoBMF - Form_Load", Me.Caption

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Select Case Button.Key
        Case "importar"
            Call flImportarArquivo
            Call flCarregaHistoricoImpoertacao
        Case "refresh"
            Call flCarregaHistoricoImpoertacao
        Case "sair"
            Unload Me
    End Select
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub

Private Function flImportarArquivo()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsRemessaFinanceiraBMF
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
        
    If Trim$(txtNomeArquivo) = vbNullString Then
        MsgBox "Selecione um arquivo", vbCritical, Me.Caption
        cmdSelecaoArquivo.SetFocus
        Exit Function
    End If
        
    Call fgCursor(True)
    DoEvents
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraBMF")
    Call objMIU.ProcessaRemessaFinanceiraBMF(txtNomeArquivo.Text, _
                                             vntCodErro, _
                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
  
    Call fgCursor(False)
    
Exit Function
ErrorHandler:
    Call fgCursor(False)
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Function

' Carregar grid com histórico de importação arquivo BMF - D0
Private Sub flCarregaHistoricoImpoertacao()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsRemessaFinanceiraBMF
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim strLeitura                              As String

Dim objListItem                             As MSComctlLib.ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
        
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraBMF")
    strLeitura = objMIU.LerTodosLogRemessa(enumLocalLiquidacao.BMD, _
                                           vntCodErro, _
                                           vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlLerTodos.loadXML(strLeitura) Then
        lstLogImportacao.ListItems.Clear
        For Each objDomNode In xmlLerTodos.documentElement.childNodes
            Set objListItem = lstLogImportacao.ListItems.Add(, , objDomNode.selectSingleNode("NO_ARQU_CAMR").Text)
            objListItem.SubItems(1) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
            objListItem.SubItems(2) = objDomNode.selectSingleNode("CO_USUA_ULTI_ATLZ").Text
        Next objDomNode
    End If

    Set xmlLerTodos = Nothing

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flPreencherHistorico", 0

End Sub

