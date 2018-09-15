VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroProcOperAtiv 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   14145
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCadastro 
      Caption         =   "Cadastro"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   9615
      Begin VB.TextBox txtTpLiquidacao 
         Height          =   280
         Left            =   8040
         TabIndex        =   38
         Top             =   2260
         Width           =   600
      End
      Begin VB.TextBox txtTpSituacao 
         Height          =   280
         Left            =   7200
         TabIndex        =   37
         Top             =   2260
         Width           =   600
      End
      Begin VB.TextBox txtTpOperacao 
         Height          =   280
         Left            =   6360
         TabIndex        =   36
         Top             =   2260
         Width           =   600
      End
      Begin VB.CheckBox chkEnviaAlerta 
         Caption         =   "Sim"
         Height          =   375
         Left            =   8500
         TabIndex        =   35
         Top             =   1800
         Width           =   600
      End
      Begin VB.CheckBox chkEstornoPJ 
         Caption         =   "Sim"
         Height          =   375
         Left            =   8500
         TabIndex        =   33
         Top             =   1440
         Width           =   600
      End
      Begin VB.CheckBox chkLancCCContab 
         Caption         =   "Sim"
         Height          =   375
         Left            =   8500
         TabIndex        =   31
         Top             =   1080
         Width           =   600
      End
      Begin VB.CheckBox chkMesgSPB 
         Caption         =   "Sim"
         Height          =   375
         Left            =   5500
         TabIndex        =   29
         Top             =   2550
         Width           =   600
      End
      Begin VB.CheckBox chkMesgRetorno 
         Caption         =   "Sim"
         Height          =   375
         Left            =   5500
         TabIndex        =   27
         Top             =   2160
         Width           =   600
      End
      Begin VB.CheckBox chkVerRegrLib 
         Caption         =   "Sim"
         Height          =   375
         Left            =   5500
         TabIndex        =   25
         Top             =   1800
         Width           =   600
      End
      Begin VB.CheckBox chkVerRegrConc 
         Caption         =   "Sim"
         Height          =   375
         Left            =   5500
         TabIndex        =   23
         Top             =   1440
         Width           =   600
      End
      Begin VB.CheckBox chkVerRegrConf 
         Caption         =   "Sim"
         Height          =   375
         Left            =   5500
         TabIndex        =   21
         Top             =   1080
         Width           =   600
      End
      Begin VB.CheckBox chkRelConfA6 
         Caption         =   "Sim"
         Height          =   375
         Left            =   2000
         TabIndex        =   19
         Top             =   2550
         Width           =   600
      End
      Begin VB.CheckBox chkRelSolA6 
         Caption         =   "Sim"
         Height          =   375
         Left            =   2000
         TabIndex        =   17
         Top             =   2160
         Width           =   600
      End
      Begin VB.CheckBox chkRelPJ 
         Caption         =   "Sim"
         Height          =   375
         Left            =   2000
         TabIndex        =   15
         Top             =   1800
         Width           =   600
      End
      Begin VB.CheckBox chkPrevA6 
         Caption         =   "Sim"
         Height          =   375
         Left            =   2000
         TabIndex        =   13
         Top             =   1440
         Width           =   600
      End
      Begin VB.CheckBox chkPrevPJ 
         Caption         =   "Sim"
         Height          =   375
         Left            =   2000
         TabIndex        =   11
         Top             =   1080
         Width           =   600
      End
      Begin VB.ComboBox cboTpSituacao 
         Height          =   315
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   3000
      End
      Begin VB.ComboBox cboTpLiquidacao 
         Height          =   315
         Left            =   7000
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cboTpOperacao 
         Height          =   315
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3000
      End
      Begin VB.TextBox txtTpProcessamento 
         Height          =   315
         Left            =   7000
         MaxLength       =   30
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envia Alerta"
         Height          =   195
         Index           =   16
         Left            =   7005
         TabIndex        =   34
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estorno PJ"
         Height          =   195
         Index           =   15
         Left            =   7005
         TabIndex        =   32
         Top             =   1520
         Width           =   765
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lanc. CC e Contab."
         Height          =   195
         Index           =   14
         Left            =   7005
         TabIndex        =   30
         Top             =   1150
         Width           =   1395
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem SPB"
         Height          =   195
         Index           =   13
         Left            =   3500
         TabIndex        =   28
         Top             =   2630
         Width           =   1140
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem Retorno"
         Height          =   195
         Index           =   12
         Left            =   3500
         TabIndex        =   26
         Top             =   2260
         Width           =   1395
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verifica Regra Liberação"
         Height          =   195
         Index           =   11
         Left            =   3500
         TabIndex        =   24
         Top             =   1890
         Width           =   1755
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verifica Regra Conciliação"
         Height          =   195
         Index           =   10
         Left            =   3500
         TabIndex        =   22
         Top             =   1520
         Width           =   1875
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verifica Regra Conferido"
         Height          =   195
         Index           =   9
         Left            =   3500
         TabIndex        =   20
         Top             =   1150
         Width           =   1725
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Realizado Conferido A6"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   2630
         Width           =   1665
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Realizado Solicitado A6"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   2260
         Width           =   1680
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realizado PJ"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   1890
         Width           =   930
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Previsão A6"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   1520
         Width           =   855
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Previsão PJ"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1150
         Width           =   840
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Liquidação"
         Height          =   255
         Index           =   3
         Left            =   5100
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo do Processamento"
         Height          =   195
         Index           =   1
         Left            =   5100
         TabIndex        =   4
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo da Operação"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1290
      End
   End
   Begin MSComctlLib.ListView lvwProcOper 
      Height          =   3015
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5318
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
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   0
      TabIndex        =   39
      Top             =   6370
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   582
      ButtonWidth     =   1535
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Excluir"
            Key             =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   5000
      Top             =   6350
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
            Picture         =   "frmCadastroProcOperAtiv.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":0112
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":042C
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":077E
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":0890
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":0BAA
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":0EC4
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroProcOperAtiv.frx":11DE
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadastroProcOperAtiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'' Exibe todos os dados do cadastro de Veículo Legal

Option Explicit

Private strOperacao                         As String
'
Private xmlLer                              As MSXML2.DOMDocument40
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private blnEditMode                         As Boolean
'
Private Const strFuncionalidade             As String = "frmCadastroProcOperAtiv"
'
Private Const COL_TP_OPER                   As Integer = 0
Private Const COL_CO_TP_OPER                As Integer = 1
Private Const COL_SITU_PROC                 As Integer = 2
Private Const COL_CO_SITU_PROC              As Integer = 3
Private Const COL_NO_PROC_OPER_ATIV         As Integer = 4
Private Const COL_TP_LIQU_OPER_ATIV         As Integer = 5
Private Const COL_CO_LIQU_OPER_ATIV         As Integer = 6
Private Const COL_IN_ENVI_PREV_PJ           As Integer = 7
Private Const COL_IN_ENVI_PREV_A6           As Integer = 8
Private Const COL_IN_ENVI_RELZ_PJ           As Integer = 9
Private Const COL_IN_ENVI_RELZ_SOLI_A6      As Integer = 10
Private Const COL_IN_ENVI_RELZ_CONF_A6      As Integer = 11
Private Const COL_IN_VERI_REGR_CONF         As Integer = 12
Private Const COL_IN_VERI_REGR_CNCL         As Integer = 13
Private Const COL_IN_VERI_REGR_LIBE         As Integer = 14
Private Const COL_IN_ENVI_MESG_RETN         As Integer = 15
Private Const COL_IN_ENVI_MESG_SPB          As Integer = 16
Private Const COL_IN_DISP_LANC_CNTA_CRRT    As Integer = 17
Private Const COL_IN_ESTO_PJ_A6             As Integer = 18
Private Const COL_IN_ENVI_ALER              As Integer = 19

Private lngIndexClassifList                 As Long

' Formatas os títulos das colunas do grid
Private Sub flPreencherHeadersLvw()

On Error GoTo ErrorHandler

    With lvwProcOper.ColumnHeaders
        .Clear
        .Add 1, , "Tipo da Operação", 2500
        .Add 2, , "Código Tipo Operação", 0
        .Add 3, , "Situação", 2500
        .Add 4, , "Código Tipo Situação", 0
        .Add 5, , "Tipo do Processamento", 2500
        .Add 6, , "Tipo de Liquidação", 2500
        .Add 7, , "Código Tipo de Liquidação", 0
        .Add 8, , "Previsão PJ", 1500
        .Add 9, , "Previsão A6", 1500
        .Add 10, , "Realizado PJ", 1500
        .Add 11, , "Realizado Solicitado A6", 1500
        .Add 12, , "Realizado Conferido A6", 1500
        .Add 13, , "Verifica Regra Conferido", 1500
        .Add 14, , "Verifica Regra Conciliação", 1500
        .Add 15, , "Verifica Regra Liberação", 1500
        .Add 16, , "Mensagem Retorno", 1500
        .Add 17, , "Mensagem SPB", 1500
        .Add 18, , "Lançamento C/C e Contabilidade", 1500
        .Add 19, , "Estorno PJ", 1500
        .Add 20, , "Envia Alerta", 1500
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPreencherHeadersLvw", 0

End Sub

' Carrega todos os registros de controle de processamento na lista
Private Sub flCarregarLista()

#If EnableSoap = 1 Then
    Dim objCtrlProcessamento     As MSSOAPLib30.SoapClient30
#Else
    Dim objCtrlProcessamento     As A7Miu.clsCtrlProcessamento
#End If

Dim xmlRetorno              As MSXML2.DOMDocument40
Dim strRetorno              As String
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim objListItem             As ListItem
Dim dtTmp                   As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
Dim ColumnHeader            As MSComctlLib.ColumnHeader

    On Error GoTo ErrorHandler

    fgCursor True
    fgLockWindow Me.hwnd

    Set objCtrlProcessamento = fgCriarObjetoMIU("A7MIU.clsCtrlProcessamento")

    strRetorno = objCtrlProcessamento.LerTodos(vntCodErro, vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    lvwProcOper.ListItems.Clear

    If strRetorno = vbNullString Then
        fgCursor
        fgLockWindow
        Exit Sub
    End If

    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    xmlRetorno.loadXML strRetorno

    For Each objDomNode In xmlRetorno.documentElement.childNodes

        Set objListItem = lvwProcOper.ListItems.Add(, "|" & objDomNode.selectSingleNode("TP_OPER").Text & objDomNode.selectSingleNode("CO_SITU_PROC").Text & "|" & objDomNode.selectSingleNode("NO_PROC_OPER_ATIV").Text & "|" & objDomNode.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
        objListItem.Text = fgSelectSingleNode(objDomNode, "TP_OPER").Text & " - " & fgSelectSingleNode(objDomNode, "NO_TIPO_OPER").Text
        objListItem.SubItems(COL_CO_TP_OPER) = fgSelectSingleNode(objDomNode, "TP_OPER").Text
        objListItem.SubItems(COL_SITU_PROC) = fgSelectSingleNode(objDomNode, "CO_SITU_PROC").Text & " - " & fgSelectSingleNode(objDomNode, "DE_SITU_PROC").Text
        objListItem.SubItems(COL_CO_SITU_PROC) = fgSelectSingleNode(objDomNode, "CO_SITU_PROC").Text
        objListItem.SubItems(COL_NO_PROC_OPER_ATIV) = fgSelectSingleNode(objDomNode, "NO_PROC_OPER_ATIV").Text
        objListItem.SubItems(COL_TP_LIQU_OPER_ATIV) = fgSelectSingleNode(objDomNode, "TP_LIQU_OPER_ATIV").Text & " - " & fgSelectSingleNode(objDomNode, "NO_TIPO_LIQU_OPER_ATIV").Text
        objListItem.SubItems(COL_CO_LIQU_OPER_ATIV) = fgSelectSingleNode(objDomNode, "TP_LIQU_OPER_ATIV").Text
        objListItem.SubItems(COL_IN_ENVI_PREV_PJ) = fgSelectSingleNode(objDomNode, "IN_ENVI_PREV_PJ").Text
        objListItem.SubItems(COL_IN_ENVI_PREV_A6) = fgSelectSingleNode(objDomNode, "IN_ENVI_PREV_A6").Text
        objListItem.SubItems(COL_IN_ENVI_RELZ_PJ) = fgSelectSingleNode(objDomNode, "IN_ENVI_RELZ_PJ").Text
        objListItem.SubItems(COL_IN_ENVI_RELZ_SOLI_A6) = fgSelectSingleNode(objDomNode, "IN_ENVI_RELZ_SOLI_A6").Text
        objListItem.SubItems(COL_IN_ENVI_RELZ_CONF_A6) = fgSelectSingleNode(objDomNode, "IN_ENVI_RELZ_CONF_A6").Text
        objListItem.SubItems(COL_IN_VERI_REGR_CONF) = fgSelectSingleNode(objDomNode, "IN_VERI_REGR_CONF").Text
        objListItem.SubItems(COL_IN_VERI_REGR_CNCL) = fgSelectSingleNode(objDomNode, "IN_VERI_REGR_CNCL").Text
        objListItem.SubItems(COL_IN_VERI_REGR_LIBE) = fgSelectSingleNode(objDomNode, "IN_VERI_REGR_LIBE").Text
        objListItem.SubItems(COL_IN_ENVI_MESG_RETN) = fgSelectSingleNode(objDomNode, "IN_ENVI_MESG_RETN").Text
        objListItem.SubItems(COL_IN_ENVI_MESG_SPB) = fgSelectSingleNode(objDomNode, "IN_ENVI_MESG_SPB").Text
        objListItem.SubItems(COL_IN_DISP_LANC_CNTA_CRRT) = fgSelectSingleNode(objDomNode, "IN_DISP_LANC_CNTA_CRRT").Text
        objListItem.SubItems(COL_IN_ESTO_PJ_A6) = fgSelectSingleNode(objDomNode, "IN_ESTO_PJ_A6").Text
        objListItem.SubItems(COL_IN_ENVI_ALER) = fgSelectSingleNode(objDomNode, "IN_ENVI_ALER").Text

    Next objDomNode

    Set objCtrlProcessamento = Nothing
    Set xmlRetorno = Nothing

    fgLockWindow 0
    fgCursor

    Exit Sub

ErrorHandler:
    fgLockWindow 0
    fgCursor
    Set objCtrlProcessamento = Nothing
    Set xmlRetorno = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0
End Sub

Private Sub cboTpOperacao_click()

    txtTpOperacao.Text = Trim(Replace(Mid(cboTpOperacao.Text, 1, 3), "-", ""))

End Sub

Private Sub cboTpSituacao_click()

    txtTpSituacao.Text = Trim(Replace(Mid(cboTpSituacao.Text, 1, 3), "-", ""))

End Sub

Private Sub cboTpLiquidacao_click()

    txtTpLiquidacao.Text = Trim(Replace(Mid(cboTpLiquidacao.Text, 1, 2), "-", ""))

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call flCarregarLista
        Call fgCursor(False)
    End If

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyDown"

End Sub

Private Sub Form_Load()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A7Miu.clsMIU
#End If

Dim xmlDomSistema                           As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
    
    On Error GoTo ErrorHandler
    
    tlbCadastro.Buttons("Limpar").Enabled = True
    tlbCadastro.Buttons("Salvar").Enabled = True
    tlbCadastro.Buttons("Excluir").Enabled = False
    tlbCadastro.Buttons("Sair").Enabled = True
    txtTpSituacao.Enabled = False
    txtTpLiquidacao.Enabled = False
    txtTpOperacao.Enabled = False
    txtTpSituacao.Visible = False
    txtTpLiquidacao.Visible = False
    txtTpOperacao.Visible = False
    
    fgCursor True
    fgCenterMe Me
    
    Set Me.Icon = mdiBUS.Icon
    Me.Caption = "Cadastro de Controle de Processamento de Operações"
    Me.Show
    DoEvents
        
    Call flPreencherHeadersLvw
    Call flLimpaCampos
    
    Set objMIU = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmFiltro", "Form_Load")
    End If

    If xmlLer Is Nothing Then
        Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
        xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades").xml
    End If

    Call flCarregarLista
    Call flCarregaComboTpOperacao
    Call flCarregaComboTpSituacao
    Call flCarregaComboTpLiquidacao
        
    fgCursor False

Exit Sub
ErrorHandler:
    fgCursor False
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next

    Dim lngAlturaLista As Long

    tlbCadastro.Top = Me.ScaleHeight - tlbCadastro.Height
    fraCadastro.Top = tlbCadastro.Top - fraCadastro.Height - 60
    fraCadastro.Width = Me.ScaleWidth - fraCadastro.Left
    lngAlturaLista = fraCadastro.Top

    With lvwProcOper
        .Top = 0
        .Left = 0
        .Width = Me.Width - 100
        .Height = lngAlturaLista - 120
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub lvwProcOper_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwProcOper, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

    Exit Sub

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - lvwProcOper_ColumnClick"

End Sub

Private Sub lvwProcOper_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error Resume Next

    With Item
        cboTpOperacao.Text = .Text
        txtTpOperacao.Text = .ListSubItems(COL_CO_TP_OPER)
        '
        cboTpSituacao.Text = .ListSubItems(COL_SITU_PROC)
        txtTpSituacao.Text = .ListSubItems(COL_CO_SITU_PROC)
        '
        txtTpProcessamento.Text = .ListSubItems(COL_NO_PROC_OPER_ATIV)
        '
        cboTpLiquidacao.Text = .ListSubItems(COL_TP_LIQU_OPER_ATIV)
        txtTpLiquidacao.Text = .ListSubItems(COL_CO_LIQU_OPER_ATIV)
        '
        If .ListSubItems(COL_IN_ENVI_PREV_PJ) = "SIM" Then
            chkPrevPJ.Value = vbChecked
        Else
            chkPrevPJ.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_PREV_A6) = "SIM" Then
            chkPrevA6.Value = vbChecked
        Else
            chkPrevA6.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_RELZ_PJ) = "SIM" Then
            chkRelPJ.Value = vbChecked
        Else
            chkRelPJ.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_RELZ_SOLI_A6) = "SIM" Then
            chkRelSolA6.Value = vbChecked
        Else
            chkRelSolA6.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_RELZ_CONF_A6) = "SIM" Then
            chkRelConfA6.Value = vbChecked
        Else
            chkRelConfA6.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_VERI_REGR_CONF) = "SIM" Then
            chkVerRegrConf.Value = vbChecked
        Else
            chkVerRegrConf.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_VERI_REGR_CNCL) = "SIM" Then
            chkVerRegrConc.Value = vbChecked
        Else
            chkVerRegrConc.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_VERI_REGR_LIBE) = "SIM" Then
            chkVerRegrLib.Value = vbChecked
        Else
            chkVerRegrLib.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_MESG_RETN) = "SIM" Then
            chkMesgRetorno.Value = vbChecked
        Else
            chkMesgRetorno.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_MESG_SPB) = "SIM" Then
            chkMesgSPB.Value = vbChecked
        Else
            chkMesgSPB.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_DISP_LANC_CNTA_CRRT) = "SIM" Then
            chkLancCCContab.Value = vbChecked
        Else
            chkLancCCContab.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ESTO_PJ_A6) = "SIM" Then
            chkEstornoPJ.Value = vbChecked
        Else
            chkEstornoPJ.Value = vbUnchecked
        End If
        '
        If .ListSubItems(COL_IN_ENVI_ALER) = "SIM" Then
            chkEnviaAlerta.Value = vbChecked
        Else
            chkEnviaAlerta.Value = vbUnchecked
        End If
        '
    End With
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = True

End Sub

Private Sub lvwProcOper_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)

        Call flCarregarLista

        Call fgCursor(False)
    End If

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - lvwProcOper_KeyDown"

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strSelecaoFiltro                        As String
Dim strResultadoConfirmacao                 As String

On Error GoTo ErrorHandler

    fgCursor True

    Call flCarregarLista
    Call flLimpaCampos

    fgCursor False

Exit Sub

ErrorHandler:

    fgCursor False
    mdiBUS.uctLogErros.MostrarErros Err, "frmCadastroProcOperAtiv - tlbFiltro_ButtonClick", Me.Caption

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case gstrSalvar
            strOperacao = gstrOperIncluir
            Call flSalvar
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
    mdiBUS.uctLogErros.MostrarErros Err, "frmCadastroProcOperAtiv - tlbCadastro_ButtonClick", Me.Caption
    Call flCarregarLista
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    End If

End Sub

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()

On Error GoTo ErrorHandler

    cboTpLiquidacao.ListIndex = -1
    txtTpLiquidacao.Text = ""
    cboTpOperacao.ListIndex = -1
    txtTpOperacao.Text = ""
    cboTpSituacao.ListIndex = -1
    txtTpSituacao.Text = ""
    txtTpProcessamento.Text = ""
    chkPrevPJ.Value = vbUnchecked
    chkPrevA6.Value = vbUnchecked
    chkRelPJ.Value = vbUnchecked
    chkRelSolA6.Value = vbUnchecked
    chkRelConfA6.Value = vbUnchecked
    chkVerRegrConf.Value = vbUnchecked
    chkVerRegrConc.Value = vbUnchecked
    chkVerRegrLib.Value = vbUnchecked
    chkMesgRetorno.Value = vbUnchecked
    chkMesgSPB.Value = vbUnchecked
    chkLancCCContab.Value = vbUnchecked
    chkEstornoPJ.Value = vbUnchecked
    chkEnviaAlerta.Value = vbUnchecked
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

'' É acionado através no botão 'Salvar' da barra de ferramentas.
''
'' Tem como função, encaminhar a solicitação (Atualização dos dados na tabela) à
'' camada controladora de caso de uso (componente / classe / metodo ) :
''
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A7Miu.clsCtrlProcessamento
#End If

Dim strRetorno              As String
Dim strPropriedades         As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()

    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If

    fgCursor True

    Call flInterfaceToXml

    Set objMIU = fgCriarObjetoMIU("A7MIU.clsCtrlProcessamento")
    If objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Set objMIU = Nothing

        Call flLimpaCampos
        Call flCarregarLista

        MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    End If

    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor False
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'Valida o preenchimento dos campos
Private Function flValidarCampos() As String

On Error GoTo ErrorHandler

    If cboTpOperacao.ListIndex = -1 Then
        flValidarCampos = "Escolha o Tipo de Operação."
        cboTpOperacao.SetFocus
        Exit Function
    End If

    If cboTpSituacao.ListIndex = -1 Then
        flValidarCampos = "Escolha o Tipo de Situação."
        cboTpSituacao.SetFocus
        Exit Function
    End If

    If Trim(txtTpProcessamento.Text) = "" Then
        flValidarCampos = "Informe o Tipo de Processamento."
        txtTpProcessamento.SetFocus
        Exit Function
    Else
        If Len(txtTpProcessamento.Text) > 30 Then
            flValidarCampos = "Tipo de Processamento deve possuir menos de 30 caracteres."
            txtTpProcessamento.SetFocus
            Exit Function
        End If
    End If

    If cboTpLiquidacao.ListIndex = -1 Then
        flValidarCampos = "Escolha o Tipo de Liquidação."
        cboTpLiquidacao.SetFocus
        Exit Function
    End If

    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Preenche o conteúdo do XML com o conteúdo dos campos apresentados em tela
Private Function flInterfaceToXml() As String

On Error GoTo ErrorHandler

    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = strOperacao
        '
        If .selectSingleNode("TP_OPER") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "TP_OPER", txtTpOperacao.Text)
        Else
            .selectSingleNode("TP_OPER").Text = txtTpOperacao.Text
        End If
        '
        If .selectSingleNode("CO_SITU_PROC") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "CO_SITU_PROC", txtTpSituacao.Text)
        Else
            .selectSingleNode("CO_SITU_PROC").Text = txtTpSituacao.Text
        End If
        '
        If .selectSingleNode("NO_PROC_OPER_ATIV") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "NO_PROC_OPER_ATIV", fgLimpaCaracterInvalido(Trim(txtTpProcessamento.Text)))
        Else
            .selectSingleNode("NO_PROC_OPER_ATIV").Text = fgLimpaCaracterInvalido(Trim(txtTpProcessamento.Text))
        End If
        '
        If .selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "TP_LIQU_OPER_ATIV", txtTpLiquidacao.Text)
        Else
            .selectSingleNode("TP_LIQU_OPER_ATIV").Text = txtTpLiquidacao.Text
        End If
        '
        If .selectSingleNode("IN_ENVI_PREV_PJ") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_PREV_PJ", chkPrevPJ.Value)
        Else
            .selectSingleNode("IN_ENVI_PREV_PJ").Text = chkPrevPJ.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_PREV_A6") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_PREV_A6", chkPrevPJ.Value)
        Else
            .selectSingleNode("IN_ENVI_PREV_A6").Text = chkPrevA6.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_RELZ_PJ") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_RELZ_PJ", chkRelPJ.Value)
        Else
            .selectSingleNode("IN_ENVI_RELZ_PJ").Text = chkRelPJ.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_RELZ_SOLI_A6") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_RELZ_SOLI_A6", chkRelSolA6.Value)
        Else
            .selectSingleNode("IN_ENVI_RELZ_SOLI_A6").Text = chkRelSolA6.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_RELZ_CONF_A6") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_RELZ_CONF_A6", chkRelConfA6.Value)
        Else
            .selectSingleNode("IN_ENVI_RELZ_CONF_A6").Text = chkRelConfA6.Value
        End If
        '
        If .selectSingleNode("IN_VERI_REGR_CONF") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_VERI_REGR_CONF", chkVerRegrConf.Value)
        Else
            .selectSingleNode("IN_VERI_REGR_CONF").Text = chkVerRegrConf.Value
        End If
        '
        If .selectSingleNode("IN_VERI_REGR_CNCL") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_VERI_REGR_CNCL", chkVerRegrConc.Value)
        Else
            .selectSingleNode("IN_VERI_REGR_CNCL").Text = chkVerRegrConc.Value
        End If
        '
        If .selectSingleNode("IN_VERI_REGR_LIBE") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_VERI_REGR_LIBE", chkVerRegrLib.Value)
        Else
            .selectSingleNode("IN_VERI_REGR_LIBE").Text = chkVerRegrLib.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_MESG_RETN") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_MESG_RETN", chkMesgRetorno.Value)
        Else
            .selectSingleNode("IN_ENVI_MESG_RETN").Text = chkMesgRetorno.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_MESG_SPB") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_MESG_SPB", chkMesgSPB.Value)
        Else
            .selectSingleNode("IN_ENVI_MESG_SPB").Text = chkMesgSPB.Value
        End If
        '
        If .selectSingleNode("IN_DISP_LANC_CNTA_CRRT") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_DISP_LANC_CNTA_CRRT", chkLancCCContab.Value)
        Else
            .selectSingleNode("IN_DISP_LANC_CNTA_CRRT").Text = chkLancCCContab.Value
        End If
        '
        If .selectSingleNode("IN_ESTO_PJ_A6") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ESTO_PJ_A6", chkEstornoPJ.Value)
        Else
            .selectSingleNode("IN_ESTO_PJ_A6").Text = chkEstornoPJ.Value
        End If
        '
        If .selectSingleNode("IN_ENVI_ALER") Is Nothing Then
            Call fgAppendNode(xmlLer, "Grupo_Propriedades", "IN_ENVI_ALER", chkEnviaAlerta.Value)
        Else
            .selectSingleNode("IN_ENVI_ALER").Text = chkEnviaAlerta.Value
        End If
    End With

Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0

End Function

'Carrega o conteúdo dos combos
Private Sub flCarregaComboTpOperacao()

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsTipoOperacao", xmlLeitura))
    Call fgCarregarCombos(cboTpOperacao, xmlLeitura, "TipoOperacao", "TP_OPER", "NO_TIPO_OPER")
    Set xmlLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flCarregaComboTpOperacao", 0

End Sub

Private Sub flCarregaComboTpSituacao()

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsTipoSituacao", xmlLeitura))
    Call fgCarregarCombos(cboTpSituacao, xmlLeitura, "TipoSituacao", "CO_SITU_PROC", "DE_SITU_PROC")
    Set xmlLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing

    fgRaiseError App.EXEName, "frmConsultaVeiculoLegal", "flCarregaComboTpSituacao", 0

End Sub

Private Sub flCarregaComboTpLiquidacao()

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsTipoLiquidacao", xmlLeitura))
    Call fgCarregarCombos(cboTpLiquidacao, xmlLeitura, "TipoLiquidacao", "TP_LIQU_OPER_ATIV", "NO_TIPO_LIQU_OPER_ATIV")
    Set xmlLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set xmlLeitura = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaComboTpLiquidacao", 0

End Sub
