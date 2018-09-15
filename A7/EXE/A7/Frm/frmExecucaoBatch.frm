VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExecucaoBatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta à Execução dos Processos Batch"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9360
   Begin MSComctlLib.ListView lvwBatch 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tarefa"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data/Hora"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descrição do Erro"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAtualizar 
      Height          =   330
      Left            =   1530
      TabIndex        =   1
      Top             =   5850
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Atualizar"
            Object.ToolTipText     =   "Atualizar"
            ImageKey        =   "Atualizar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   5850
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ButtonWidth     =   2434
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Definir Filtro"
            Key             =   "DefinirFiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageKey        =   "DefinirFiltro"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   6615
      Top             =   5850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":0112
            Key             =   "Erro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":042C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":0746
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":0A98
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":0BAA
            Key             =   "st_aguardando"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":0FFC
            Key             =   "st_postado"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":144E
            Key             =   "st_traduzido"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":18A0
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":1CF2
            Key             =   "st_recebido"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":2144
            Key             =   "st_entregue"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":2596
            Key             =   "st_root_old"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":29E8
            Key             =   "st_root"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":2E3A
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":318C
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":34DE
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":3830
            Key             =   "vermelho"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":3B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":3FD4
            Key             =   "st_cancelado"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandosForm 
      Height          =   330
      Left            =   8595
      TabIndex        =   3
      Top             =   5850
      Width           =   765
      _ExtentX        =   1349
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
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmExecucaoBatch.frx":4426
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":4538
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":4852
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":4BA4
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":4CB6
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":4FD0
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":52EA
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":5604
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":591E
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":5C70
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":5FC2
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExecucaoBatch.frx":6314
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExecucaoBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela apresentação dos logs de execução de rotinas batch.
Option Explicit

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private WithEvents objFiltro                As frmFiltroExecucaoBatch
Attribute objFiltro.VB_VarHelpID = -1
Private xmlFiltro                           As MSXML2.DOMDocument40

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler
    
    If KeyAscii = vbKeyF5 Then
        fgCursor True
        flAtualizar
        fgCursor
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmExecucaoBatch - Form_KeyPress")

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True
    
    fgCenterMe Me
    Me.Icon = mdiBUS.Icon
    Me.Show
    
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Set objFiltro = New frmFiltroExecucaoBatch
    objFiltro.Show vbModal

    fgCursor
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmExecucaoBatch - Form_Load")

End Sub
'Atulizar as informações da tela.
Private Sub flAtualizar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlExecucao         As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlExecucao = CreateObject("MSXML2.DOMDocument.4.0")

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlExecucao.loadXML objMiu.Executar(xmlFiltro.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    lvwBatch.ListItems.Clear
    
    If xmlExecucao.xml = vbNullString Then Exit Sub
    
    For Each xmlNode In xmlExecucao.selectNodes("//Repeat_Execucao/*")
        With lvwBatch.ListItems.Add(, , flCodigoRotinaToString(xmlNode.selectSingleNode(".//CO_ROTI_BATCH").Text), , _
                                    IIf(CLng(xmlNode.selectSingleNode(".//IN_EXEC_SUCE").Text) = enumIndicadorSimNao.sim, "Ok", "Erro"))
            .SubItems(1) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode(".//DH_FIM_EXEC").Text)
            .SubItems(2) = IIf(CLng(xmlNode.selectSingleNode(".//IN_EXEC_SUCE").Text) = enumIndicadorSimNao.sim, "Ok", "Erro")
            .SubItems(3) = xmlNode.selectSingleNode(".//DE_ERRO_EXEC").Text
        End With
    Next
    
    Set xmlExecucao = Nothing
    
    Exit Sub
ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlExecucao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, Me.Name, "flAtualizar Sub", lngCodigoErroNegocio)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objFiltro = Nothing
    Set xmlFiltro = Nothing
    Unload Me
    
End Sub

Private Sub lvwBatch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwBatch, ColumnHeader.Index)
    
    Exit Sub
    
ErrorHandler:

    mdiBUS.uctLogErros.MostrarErros Err, "frmExecucaoBatch - lvwBatch_ColumnClick"

End Sub

Private Sub objFiltro_AplicarFiltro(ByVal pXMLFiltro As MSXML2.DOMDocument40)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Set xmlFiltro = pXMLFiltro

    Call flAtualizar
    
    fgCursor
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    Call fgRaiseError(App.EXEName, Me.Name, "objFiltro_AplicarFiltro Sub", lngCodigoErroNegocio)

End Sub

Private Sub tlbAtualizar_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    If Button.Key = "Atualizar" Then
        fgCursor True
        flAtualizar
        fgCursor
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmExecucaoBatch - tlbAtualizar_ButtonClick")

End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    If Button.Key = "DefinirFiltro" Then
        objFiltro.Show vbModal
    End If
    
    Exit Sub
ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmExecucaoBatch - tlbButtons_ButtonClick")

End Sub

Private Function flCodigoRotinaToString(ByRef plngCodigoRotinaBatch As Long) As String
    
    Select Case plngCodigoRotinaBatch
        Case enumRotinaBatch.ReplicacaoPJPK
            flCodigoRotinaToString = "Replicação PJPK"
        Case enumRotinaBatch.IntegracaoHA
            flCodigoRotinaToString = "Integração HA"
        Case enumRotinaBatch.ExpurgoA6
            flCodigoRotinaToString = "Expurgo Sistema A6"
        Case enumRotinaBatch.ExpurgoA7
            flCodigoRotinaToString = "Expurgo Sistema A7"
        Case enumRotinaBatch.ExpurgoA8
            flCodigoRotinaToString = "Expurgo Sistema A8"
    End Select

End Function

Private Sub tlbComandosForm_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "Sair" Then
        Set objFiltro = Nothing
        Set xmlFiltro = Nothing
        Unload Me
    End If
    

End Sub
