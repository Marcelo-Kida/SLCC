VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFiltroExecucaoBatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro de Execução de Rotinas Batch"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   4485
      Begin VB.ComboBox cboTarefa 
         Height          =   315
         ItemData        =   "frmFiltroLogExecucaoBatch.frx":0000
         Left            =   975
         List            =   "frmFiltroLogExecucaoBatch.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   3375
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmFiltroLogExecucaoBatch.frx":0004
         Left            =   975
         List            =   "frmFiltroLogExecucaoBatch.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpDataInicio 
         Height          =   315
         Left            =   975
         TabIndex        =   3
         Top             =   900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63700993
         CurrentDate     =   36892
         MinDate         =   36892
      End
      Begin MSComCtl2.DTPicker dtpDataFim 
         Height          =   315
         Left            =   2835
         TabIndex        =   4
         Top             =   900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63700993
         CurrentDate     =   36892
         MinDate         =   36892
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   2610
         TabIndex        =   9
         Top             =   990
         Width           =   90
      End
      Begin VB.Label lblSistema 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "Tarefa:"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   225
         Width           =   510
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   2430
      TabIndex        =   5
      Top             =   1395
      Width           =   2100
      _ExtentX        =   3704
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
            Caption         =   "OK"
            Key             =   "OK"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   1290
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
            Picture         =   "frmFiltroLogExecucaoBatch.frx":0008
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":011A
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":0434
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":0786
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":0898
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":0BB2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":0ECC
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":11E6
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":1500
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":1852
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":1BA4
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroLogExecucaoBatch.frx":1EF6
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFiltroExecucaoBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela configuração do filtro da apresentação dos logs de execução de rotinas batch.
Option Explicit

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Public Event AplicarFiltro(ByVal pXMLFiltro As MSXML2.DOMDocument40)

Private Sub dtpDataInicio_Change()
    
    If dtpDataFim.Value < dtpDataInicio.Value Then
        dtpDataFim.Value = dtpDataInicio.Value + 30
        dtpDataFim.MinDate = dtpDataInicio.Value
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            Me.Hide
        Case vbKeyReturn
            tlbCadastro_ButtonClick tlbCadastro.Buttons("OK")
    End Select

End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    fgCursor True
    
    fgCenterMe Me
    
    Me.Icon = mdiBUS.Icon
    
    flCarregarCboStatus
    flCarregarCboTarefa
    flIniciarDatasPeriodo

    fgCursor
    
    Exit Sub
ErrorHandler:
    fgCursor
    mdiBUS.uctLogErros.MostrarErros Err, ("frmFiltroExecucaoBatch - Form_Load")
End Sub

Private Sub flCarregarCboTarefa()

    cboTarefa.Clear
    cboTarefa.AddItem "Todos"
    cboTarefa.AddItem "Replicação PJPK"
    cboTarefa.ItemData(cboTarefa.NewIndex) = enumRotinaBatch.ReplicacaoPJPK
    cboTarefa.AddItem "Integração HA"
    cboTarefa.ItemData(cboTarefa.NewIndex) = enumRotinaBatch.IntegracaoHA
    cboTarefa.AddItem "Expurgo Sistema A6"
    cboTarefa.ItemData(cboTarefa.NewIndex) = enumRotinaBatch.ExpurgoA6
    cboTarefa.AddItem "Expurgo Sistema A7"
    cboTarefa.ItemData(cboTarefa.NewIndex) = enumRotinaBatch.ExpurgoA7
    cboTarefa.AddItem "Expurgo Sistema A8"
    cboTarefa.ItemData(cboTarefa.NewIndex) = enumRotinaBatch.ExpurgoA8
    cboTarefa.ListIndex = 0
    
End Sub

Private Sub flCarregarCboStatus()

    cboStatus.Clear
    cboStatus.AddItem "Ambos"
    cboStatus.AddItem "Ok"
    cboStatus.ItemData(cboStatus.NewIndex) = enumIndicadorSimNao.sim
    cboStatus.AddItem "Erro"
    cboStatus.ItemData(cboStatus.NewIndex) = enumIndicadorSimNao.nao
    cboStatus.ListIndex = 0

End Sub


Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim xmlFiltro                               As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "OK"
            Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
            fgAppendNode xmlFiltro, "", "Filtro", ""
            fgAppendAttribute xmlFiltro, "Filtro", "Operacao", "Ler"
            fgAppendAttribute xmlFiltro, "Filtro", "Objeto", "A7Server.clsExecucaoBatch"
            
            If cboTarefa.ListIndex > 0 Then
                fgAppendNode xmlFiltro, _
                             "Filtro", _
                             "CO_ROTI_BATCH", _
                             cboTarefa.ItemData(cboTarefa.ListIndex)
            End If
            
            If cboStatus.ListIndex > 0 Then
                fgAppendNode xmlFiltro, _
                             "Filtro", _
                             "IN_EXEC_SUCE", _
                             cboStatus.ItemData(cboStatus.ListIndex)
            End If
            
            fgAppendNode xmlFiltro, _
                         "Filtro", _
                         "DT_INIC", _
                         fgDt_To_Xml(dtpDataInicio.Value)
            
            fgAppendNode xmlFiltro, _
                         "Filtro", _
                         "DT_FIM", _
                         fgDt_To_Xml(dtpDataFim.Value)
            
            RaiseEvent AplicarFiltro(xmlFiltro)
            
            Set xmlFiltro = Nothing
            
            Me.Hide
        Case "Cancelar"
            Me.Hide
    End Select

    Exit Sub
ErrorHandler:
    mdiBUS.uctLogErros.MostrarErros Err, ("frmFiltroExecucaoBatch - tlbCadastro_ButtonClick")

End Sub

Private Sub flIniciarDatasPeriodo()

Dim dtFim                                   As Date

    dtFim = DateSerial(Year(Date), Month(Date) + 1, 1)
    dtFim = FormatDateTime(dtFim - 1, vbShortDate)
    
    dtpDataInicio.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpDataFim.Value = dtFim

    dtpDataFim.MinDate = dtpDataInicio.Value

End Sub
