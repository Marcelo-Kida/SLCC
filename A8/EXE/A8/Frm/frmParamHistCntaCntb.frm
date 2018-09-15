VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParamHistCntaCntb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Conta e Histórico Contábil"
   ClientHeight    =   7260
   ClientLeft      =   1935
   ClientTop       =   2325
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCadastro 
      Height          =   6660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9525
      Begin VB.Frame fraTipoConta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3870
         Left            =   60
         TabIndex        =   10
         Top             =   2700
         Width           =   9360
         Begin VB.TextBox txtTipoOperacao 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   2280
            Width           =   5595
         End
         Begin VB.ComboBox cboTipoOperacao 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2280
            Width           =   5595
         End
         Begin VB.TextBox txtCentroDestino 
            Height          =   315
            Left            =   3540
            MaxLength       =   4
            TabIndex        =   25
            Top             =   2880
            Width           =   1845
         End
         Begin VB.ComboBox cboLocalLiquidacao 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1680
            Width           =   5595
         End
         Begin VB.TextBox txtLocalLiquidacao 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   5595
         End
         Begin VB.ComboBox cboSistema 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1080
            Width           =   5595
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   1830
            MaxLength       =   14
            TabIndex        =   7
            Top             =   3450
            Width           =   7245
         End
         Begin VB.TextBox txtHistorico 
            Height          =   315
            Left            =   150
            MaxLength       =   3
            TabIndex        =   6
            Top             =   3450
            Width           =   1605
         End
         Begin VB.TextBox txtContaCredito 
            Height          =   315
            Left            =   1830
            MaxLength       =   5
            TabIndex        =   3
            Top             =   2880
            Width           =   1605
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   5595
         End
         Begin VB.Frame fraTipoDebitoCredito 
            Caption         =   "Tipo Lançamento Conta Corrente"
            Height          =   855
            Left            =   5820
            TabIndex        =   12
            Top             =   270
            Width           =   3375
            Begin VB.OptionButton optEstornoCredito 
               Caption         =   "Estorno Crédito"
               Height          =   195
               Left            =   1800
               TabIndex        =   27
               Top             =   570
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.OptionButton optEstornoDebito 
               Caption         =   "Estorno Débito"
               Height          =   195
               Left            =   180
               TabIndex        =   26
               Top             =   570
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.OptionButton optDebito 
               Caption         =   "Débito"
               Height          =   195
               Left            =   180
               TabIndex        =   4
               Top             =   420
               Width           =   795
            End
            Begin VB.OptionButton optCredito 
               Caption         =   "Crédito"
               Height          =   195
               Left            =   1800
               TabIndex        =   5
               Top             =   420
               Width           =   855
            End
         End
         Begin VB.TextBox txtContaDebito 
            Height          =   315
            Left            =   150
            MaxLength       =   5
            TabIndex        =   2
            Top             =   2880
            Width           =   1605
         End
         Begin VB.TextBox txtEmpresa 
            Enabled         =   0   'False
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Top             =   420
            Width           =   5595
         End
         Begin VB.TextBox txtSistema 
            Enabled         =   0   'False
            Height          =   285
            Left            =   135
            TabIndex        =   20
            Top             =   1080
            Width           =   5595
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operação"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   1065
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Código Centro de Destino"
            Height          =   195
            Left            =   3540
            TabIndex        =   24
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Local de Liquidação"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Sistema"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   810
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Descrição do Histórico Contábil"
            Height          =   195
            Left            =   1830
            TabIndex        =   18
            Top             =   3210
            Width           =   2220
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código do Histórico"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   3210
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Conta Contábil Crédito"
            Height          =   195
            Left            =   1830
            TabIndex        =   16
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Conta Contábil Débito"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   2640
            Width           =   1545
         End
      End
      Begin MSComctlLib.ListView lvwConta 
         Height          =   2505
         Left            =   45
         TabIndex        =   8
         Top             =   210
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   4419
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   60
      Top             =   6210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamHistCntaCntb.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   5565
      TabIndex        =   15
      Top             =   6720
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParamHistCntaCntb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cadastro de Contas por Veículo Legal

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private strOperacao                         As String

Private strKeyItemSelected                  As String

Public lngBackOffice                        As Long

Private Const COL_EMPRESA                   As Integer = 0
Private Const COL_SISTEMA                   As Integer = 1
Private Const COL_LOCALLIQUIDACAO           As Integer = 2
Private Const COL_TIPO_MOVTO                As Integer = 3
Private Const COL_CONTA_DEBITO              As Integer = 4
Private Const COL_CONTA_CREDITO             As Integer = 5
Private Const COL_CENTRO_DESTINO            As Integer = 6
Private Const COL_HISTORICO                 As Integer = 7
Private Const COL_DESCRICAO                 As Integer = 8
Private Const COL_TIPOOPERACAO              As Integer = 9

Private Const strFuncionalidade             As String = "frmParamHistCntaCntb"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

'Seleciona o item do listview de acordo com a seleção atual
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lvwConta.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub

    blnEncontrou = False
    For Each objListItem In lvwConta.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lvwConta_ItemClick objListItem
           lvwConta.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing

    If Not blnEncontrou Then
       flLimpaCampos
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

'Formata as colunas do listview
Private Sub flFormataListView()

    lvwConta.ColumnHeaders.Add 1, , "Empresa", 1800, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 2, , "Sistema", 2000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 3, , "Local de Liquidação", 2000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 4, , "Tipo Movto.", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 5, , "Conta Débito", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 6, , "Conta Crédito", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 7, , "Centro Destino", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 8, , "Código Histórico", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 9, , "Descrição Histórico", 1000, lvwColumnLeft
    lvwConta.ColumnHeaders.Add 10, , "Tipo Operação", 1000, lvwColumnLeft

End Sub

'Salva as alterações efetuadas
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strRetorno                              As String
Dim strPropriedades                         As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()

    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If

    Call fgCursor(True)

    Call flInterfaceToXml

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    If strOperacao = "Incluir" Then
        With xmlLer.documentElement
            strKeyItemSelected = "|" & .selectSingleNode("SG_SIST").Text & _
                                 "|" & .selectSingleNode("CO_EMPR").Text & _
                                 "|" & .selectSingleNode("TP_BKOF").Text & _
                                 "|" & .selectSingleNode("CO_LOCA_LIQU").Text & _
                                 "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                 "|" & .selectSingleNode("TP_OPER").Text
        End With
    End If

    Call objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        strOperacao = gstrOperAlterar
        flXmlToInterface
    Else
        flLimpaCampos
    End If
    Set objMIU = Nothing

    Call flCarregaListView
    Call fgCursor(False)

    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
Exit Sub
ErrorHandler:

    Call fgCursor(False)

    Set objMIU = Nothing

    If strOperacao <> gstrOperExcluir Then
       With xmlLer.documentElement
            strKeyItemSelected = "|" & .selectSingleNode("SG_SIST").Text & _
                                 "|" & .selectSingleNode("CO_EMPR").Text & _
                                 "|" & .selectSingleNode("TP_BKOF").Text & _
                                 "|" & .selectSingleNode("CO_LOCA_LIQU").Text & _
                                 "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                 "|" & .selectSingleNode("TP_OPER").Text
       End With
    End If
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    fgRaiseError App.EXEName, "frmParamHistCntaCntb", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Valida os campos preenchidos
Private Function flValidarCampos() As String

On Error GoTo ErrorHandler
    
    If strOperacao = gstrOperIncluir Then
        If cboEmpresa.ListIndex = -1 Then
            flValidarCampos = "Selecione a Empresa."
            cboEmpresa.SetFocus
            Exit Function
        End If
        
        If cboSistema.ListIndex = -1 Then
            flValidarCampos = "Selecione um Sistema."
            cboSistema.SetFocus
            Exit Function
        End If
        
        If Not (optDebito.value Or optCredito.value Or optEstornoDebito.value Or optEstornoCredito.value) Then
            flValidarCampos = "Selecione o Tipo de Lançamento."
            Exit Function
        End If
        
        If cboLocalLiquidacao.ListIndex = -1 Then
            flValidarCampos = "Selecione o Local de Liquidação."
            cboLocalLiquidacao.SetFocus
            Exit Function
        End If
    
        If cboTipoOperacao.ListIndex = -1 Then
            flValidarCampos = "Selecione o Tipo de Operação."
            cboTipoOperacao.SetFocus
            Exit Function
        End If
    End If
    
    With txtContaDebito
        If Val(.Text) = 0 Then
            flValidarCampos = "Informe a Conta Débito."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With
    
    With txtContaCredito
        If Val(.Text) = 0 Then
            flValidarCampos = "Informe a Conta Crédito."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
        
        If Trim$(txtContaDebito.Text) = Trim$(.Text) Then
            flValidarCampos = "Conta Crédito deve ser diferente da Conta Débito."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With
    
    With txtCentroDestino
        If Trim$(.Text) = vbNullString Or Val(.Text) = 0 Then
            flValidarCampos = "Informe o Centro de Destino"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With
    
    With txtHistorico
        If Trim$(.Text) = vbNullString Then
            flValidarCampos = "Informe o Código do Histórico."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With

    With txtDescricao
        If Trim$(.Text) = vbNullString Then
            flValidarCampos = "Informe a Descrição do Histórico."
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()

On Error GoTo ErrorHandler

    strOperacao = "Incluir"

    cboEmpresa.ListIndex = -1
    cboEmpresa.Visible = True
    txtEmpresa.Visible = False
    
    cboSistema.ListIndex = -1
    cboSistema.Visible = True
    txtSistema.Visible = False
        
    txtContaDebito.Text = vbNullString
    txtContaCredito.Text = vbNullString
    txtHistorico.Text = vbNullString
    txtCentroDestino.Text = vbNullString
    txtDescricao.Text = vbNullString
    
    cboLocalLiquidacao.ListIndex = -1
    cboLocalLiquidacao.Visible = True
    txtLocalLiquidacao.Visible = False
    
    cboTipoOperacao.ListIndex = -1
    cboTipoOperacao.Visible = True
    txtTipoOperacao.Visible = False
    
    fraTipoDebitoCredito.Enabled = True
    optCredito.value = False
    optDebito.value = False
    optEstornoCredito.value = False
    optEstornoDebito.value = False

    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    
    cboEmpresa.SetFocus
        
Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, "frmParamHistCntaCntb", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Preenche os campos em tela com o conteúdo do documento XML
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
          Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
          Dim objMIU                              As A8MIU.clsMIU
#End If

      Dim strChaveRegistro                        As String
      Dim vntCodErro                              As Variant
      Dim vntMensagemErro                         As Variant

10    On Error GoTo ErrorHandler

20        If lvwConta.SelectedItem Is Nothing Then
30            flLimpaCampos
40            Exit Sub
50        End If

60        strChaveRegistro = lvwConta.SelectedItem.Key

70        With xmlLer.documentElement
80            .selectSingleNode("//@Operacao").Text = "Ler"
90            .selectSingleNode("//SG_SIST").Text = Split(strChaveRegistro, "|")(1)
100           .selectSingleNode("//CO_EMPR").Text = Split(strChaveRegistro, "|")(2)
110           .selectSingleNode("//TP_BKOF").Text = Split(strChaveRegistro, "|")(3)
120           .selectSingleNode("//CO_LOCA_LIQU").Text = Split(strChaveRegistro, "|")(4)
130           .selectSingleNode("//IN_LANC_DEBT_CRED").Text = Split(strChaveRegistro, "|")(5)
140           .selectSingleNode("//TP_OPER").Text = Split(strChaveRegistro, "|")(6)
              
150       End With

160       Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
170       Call xmlLer.loadXML(objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
          
180       If vntCodErro <> 0 Then
190           GoTo ErrorHandler
200       End If
          
210       Set objMIU = Nothing

220       With xmlLer.documentElement
              
230           txtEmpresa.Visible = True
240           cboEmpresa.Visible = False
              
250           If Not xmlMapaNavegacao.selectSingleNode("//Grupo_Empresa[CO_EMPR=" & .selectSingleNode("CO_EMPR").Text & "]/NO_REDU_EMPR") Is Nothing Then
260               txtEmpresa.Text = .selectSingleNode("CO_EMPR").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_Empresa[CO_EMPR=" & .selectSingleNode("CO_EMPR").Text & "]/NO_REDU_EMPR").Text
270           Else
280               txtEmpresa.Text = .selectSingleNode("CO_EMPR").Text
290           End If
              
300           txtLocalLiquidacao.Visible = True
310           cboLocalLiquidacao.Visible = False
320           txtLocalLiquidacao.Text = .selectSingleNode("CO_LOCA_LIQU").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & .selectSingleNode("CO_LOCA_LIQU").Text & "']/DE_LOCA_LIQU").Text
              
330           txtSistema.Visible = True
340           cboSistema.Visible = False
350           txtSistema.Text = lvwConta.SelectedItem.SubItems(1)
                      
360           txtContaDebito.Text = .selectSingleNode("CO_CNTA_DEBT").Text
370           txtContaCredito.Text = .selectSingleNode("CO_CNTA_CRED").Text
380           txtCentroDestino.Text = .selectSingleNode("CO_CENT_DEST").Text
390           txtHistorico.Text = .selectSingleNode("CO_HIST_CNTA_CNTB").Text
400           txtDescricao.Text = .selectSingleNode("DE_HIST_CNTA_CNTB").Text
                      
410           txtTipoOperacao.Visible = True
420           cboTipoOperacao.Visible = False
              
430           If .selectSingleNode("TP_OPER").Text > 0 Then
440               txtTipoOperacao.Text = .selectSingleNode("TP_OPER").Text & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_TipoOperacao[TP_OPER='" & .selectSingleNode("TP_OPER").Text & "']/NO_TIPO_OPER").Text
450           Else
460               txtTipoOperacao.Text = "<Padrão>"
470           End If
              
480           fraTipoDebitoCredito.Enabled = False
              
490           Select Case Val(.selectSingleNode("IN_LANC_DEBT_CRED").Text)
                  Case enumTipoDebitoCredito.Credito
500                   optCredito.value = True
510               Case enumTipoDebitoCredito.Debito
520                   optDebito.value = True
530               Case enumTipoDebitoCreditoEstorno.EstornoCredito
540                   optEstornoCredito.value = True
550               Case enumTipoDebitoCreditoEstorno.EstornoDebito
560                   optEstornoDebito.value = True
570           End Select
              
580           strUltimaAtualizacao = .selectSingleNode("DH_ULTI_ATLZ").Text
              
590           tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
600       End With

610   Exit Sub
ErrorHandler:

620       Set objMIU = Nothing
          
630       If vntCodErro <> 0 Then
640           Err.Number = vntCodErro
650           Err.Description = vntMensagemErro
660       End If

670       fgRaiseError App.EXEName, "frmParamHistCntaCntb", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro, "Linha: " & Erl

End Sub

'Preenche as tags do documento XML com o conteúdo dos controles da tela
Private Function flInterfaceToXml() As String

Dim lngTipoDebitoCredito                    As Long

On Error GoTo ErrorHandler

    With xmlLer.documentElement

         .selectSingleNode("@Operacao").Text = strOperacao
         
         If strOperacao <> gstrOperExcluir Then
         
            If strOperacao = gstrOperIncluir Then
               .selectSingleNode("SG_SIST").Text = fgObterCodigoCombo(cboSistema.Text)
               .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa.Text)
               .selectSingleNode("TP_BKOF").Text = lngBackOffice
               .selectSingleNode("CO_LOCA_LIQU").Text = fgObterCodigoCombo(cboLocalLiquidacao.Text)
               
                Select Case True
                    Case optCredito.value
                        lngTipoDebitoCredito = enumTipoDebitoCredito.Credito
                    Case optDebito.value
                        lngTipoDebitoCredito = enumTipoDebitoCredito.Debito
                    Case optEstornoCredito.value
                        lngTipoDebitoCredito = enumTipoDebitoCreditoEstorno.EstornoCredito
                    Case optEstornoDebito.value
                        lngTipoDebitoCredito = enumTipoDebitoCreditoEstorno.EstornoDebito
                End Select
               
               .selectSingleNode("IN_LANC_DEBT_CRED").Text = lngTipoDebitoCredito
               .selectSingleNode("TP_OPER").Text = fgObterCodigoCombo(cboTipoOperacao.Text)
            End If
            
            .selectSingleNode("CO_CNTA_DEBT").Text = txtContaDebito.Text
            .selectSingleNode("CO_CNTA_CRED").Text = txtContaCredito.Text
            .selectSingleNode("CO_CENT_DEST").Text = txtCentroDestino.Text
            .selectSingleNode("CO_HIST_CNTA_CNTB").Text = txtHistorico.Text
            .selectSingleNode("DE_HIST_CNTA_CNTB").Text = txtDescricao.Text
         
         End If

    End With

Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0

End Function

'Inicializa os controles e variáveis
Private Sub flInicializar()

#If EnableSoap = 1 Then
          Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
          Dim objMIU              As A8MIU.clsMIU
#End If

      Dim vntCodErro              As Variant
      Dim vntMensagemErro         As Variant

10    On Error GoTo ErrorHandler

20        Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
30        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")

40        If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
              
50            If vntCodErro <> 0 Then
60                GoTo ErrorHandler
70            End If
              
80            Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
90        End If

100       lngBackOffice = CLng("0" & xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_TipoBackOffice/TP_BKOF").Text)
110       Call fgCarregarCombos(cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
120       Call fgCarregarCombos(cboLocalLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU")
            
130       Call flCarregarComboTipoOperacao
          
140       If xmlLer Is Nothing Then
150          Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
160          xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_ParmHistCntaCntb").xml
170       End If

180       Set objMIU = Nothing
190       Exit Sub

ErrorHandler:

200       Set objMIU = Nothing
210       Set xmlMapaNavegacao = Nothing
          
220       If vntCodErro <> 0 Then
230           Err.Number = vntCodErro
240           Err.Description = vntMensagemErro
250       End If

260       fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0, 0, "Linha: " & Erl

End Sub

'Define o número máximo de caracteres permitidos nos controles
Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0
End Sub

'Carrega o conteúdo do ListView
Private Sub flCarregaListView()

#If EnableSoap = 1 Then
          Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
          Dim objMIU                              As A8MIU.clsMIU
#End If

      Dim xmlLerTodos                             As MSXML2.DOMDocument40
      Dim xmlDomNode                              As MSXML2.IXMLDOMNode
      Dim xmlTipoOperacao                         As MSXML2.IXMLDOMNode
      Dim objListItem                             As MSComctlLib.ListItem
      Dim strNomeTipoLiquidacao                   As String
      Dim strDescTipoMovimento                    As String

      Dim strTempChave                            As String
      Dim vntCodErro                              As Variant
      Dim vntMensagemErro                         As Variant

10    On Error GoTo ErrorHandler

20        fgCursor True

30        lvwConta.ListItems.Clear
40        lvwConta.HideSelection = False

50        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
60        Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")

70        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParmHistCntaCntb/@Operacao").Text = "LerTodos"
80        Call xmlLerTodos.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParmHistCntaCntb").xml, _
                                                   vntCodErro, _
                                                   vntMensagemErro))

90        If vntCodErro <> 0 Then
100           GoTo ErrorHandler
110       End If

120       Set objMIU = Nothing

130       For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_ParmHistCntaCntb/*")
140           With xmlDomNode
                  
150               If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & .selectSingleNode("CO_LOCA_LIQU").Text & "']/DE_LOCA_LIQU") Is Nothing Then
                  
160                   strTempChave = "|" & .selectSingleNode("SG_SIST").Text & _
                                     "|" & .selectSingleNode("CO_EMPR").Text & _
                                     "|" & .selectSingleNode("TP_BKOF").Text & _
                                     "|" & .selectSingleNode("CO_LOCA_LIQU").Text & _
                                     "|" & .selectSingleNode("IN_LANC_DEBT_CRED").Text & _
                                     "|" & .selectSingleNode("TP_OPER").Text
          
170                   Set objListItem = lvwConta.ListItems.Add(, strTempChave)
                      
180                   If xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & .selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR") Is Nothing Then
190                       objListItem.Text = .selectSingleNode("CO_EMPR").Text
200                   Else
210                       objListItem.Text = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & .selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
220                   End If
                      
                      
230                   objListItem.SubItems(COL_SISTEMA) = .selectSingleNode("SG_SIST").Text & " - " & .selectSingleNode("NO_SIST").Text
                      
240                   Select Case .selectSingleNode("IN_LANC_DEBT_CRED").Text
                          Case enumTipoDebitoCredito.Credito
250                           strDescTipoMovimento = "Crédito"
260                       Case enumTipoDebitoCredito.Debito
270                           strDescTipoMovimento = "Débito"
280                       Case enumTipoDebitoCreditoEstorno.EstornoCredito
290                           strDescTipoMovimento = "Estorno Crédito"
300                       Case enumTipoDebitoCreditoEstorno.EstornoDebito
310                           strDescTipoMovimento = "Estorno Débito"
320                   End Select
                      
330                   objListItem.SubItems(COL_TIPO_MOVTO) = strDescTipoMovimento
340                   objListItem.SubItems(COL_LOCALLIQUIDACAO) = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & .selectSingleNode("CO_LOCA_LIQU").Text & "']/DE_LOCA_LIQU").Text
350                   objListItem.SubItems(COL_CONTA_DEBITO) = .selectSingleNode("CO_CNTA_DEBT").Text
360                   objListItem.SubItems(COL_CONTA_CREDITO) = .selectSingleNode("CO_CNTA_CRED").Text
370                   objListItem.SubItems(COL_CENTRO_DESTINO) = .selectSingleNode("CO_CENT_DEST").Text
380                   objListItem.SubItems(COL_HISTORICO) = .selectSingleNode("CO_HIST_CNTA_CNTB").Text
390                   objListItem.SubItems(COL_DESCRICAO) = .selectSingleNode("DE_HIST_CNTA_CNTB").Text
                      
400                   If .selectSingleNode("TP_OPER").Text > 0 Then
410                       For Each xmlTipoOperacao In xmlMapaNavegacao.selectNodes("//Repeat_TipoOperacao/Grupo_TipoOperacao[TP_MESG_RECB_INTE='154']")
420                           If Not xmlTipoOperacao Is Nothing Then
430                               If CLng(xmlTipoOperacao.selectSingleNode("TP_OPER").Text) = .selectSingleNode("TP_OPER").Text Then
440                                    objListItem.SubItems(COL_TIPOOPERACAO) = xmlTipoOperacao.selectSingleNode("TP_OPER").Text & _
                                       " - " & xmlTipoOperacao.selectSingleNode("NO_TIPO_OPER").Text
450                                   Exit For
460                               End If
470                           Else
480                               objListItem.SubItems(COL_TIPOOPERACAO) = ""
490                               Exit For
500                           End If
510                       Next
                                   
520                   Else
530                       objListItem.SubItems(COL_TIPOOPERACAO) = ""
540                   End If
550               End If
560           End With
570       Next

580       Set xmlLerTodos = Nothing
590       fgCursor

600   Exit Sub
ErrorHandler:
610       Set objMIU = Nothing
620       Set xmlLerTodos = Nothing
630       fgCursor
          
640       If vntCodErro <> 0 Then
650           Err.Number = vntCodErro
660           Err.Description = vntMensagemErro
670       End If

680       fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0, 0, "Linha: " & Erl

End Sub

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    If cboEmpresa.ListIndex <> -1 Then
        
        If fgObterCodigoCombo(cboEmpresa.Text) = "701" Then
            optEstornoCredito.Enabled = False
            optEstornoCredito.value = False
            optEstornoDebito.Enabled = False
            optEstornoDebito.value = False
        Else
            optEstornoCredito.Enabled = True
            optEstornoDebito.Enabled = True
        End If
        
        fgCursor True
        flCarregaComboSistema
        fgCursor False
        
    End If
       
Exit Sub
ErrorHandler:
   fgCursor False
    
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboEmpresa_Click"
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

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao

    Call flLimpaCampos

    Call fgCursor(True)
    Call flFormataListView
    Call flInicializar
    Call flDefinirTamanhoMaximoCampos
    Call flCarregaListView

    Call fgCursor(False)

Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmParamHistCntaCntb - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    Set xmlLer = Nothing

End Sub

Private Sub lvwConta_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lvwConta.Sorted = True
    lvwConta.SortKey = ColumnHeader.Index - 1

    If lvwConta.SortOrder = lvwAscending Then
        lvwConta.SortOrder = lvwDescending
    Else
        lvwConta.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwConta_ColumnClick"

End Sub

Private Sub lvwConta_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface

    strKeyItemSelected = Item.Key

    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmParamHistCntaCntb - lvwConta_ItemClick", Me.Caption
    flRecarregar

End Sub

'Atualiza o conteúdo da tela
Private Sub flRecarregar()

On Error GoTo ErrorHandler

    fgCursor True

    flLimpaCampos
    Call flCarregaListView

    fgCursor

Exit Sub
ErrorHandler:

    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flRecarregar"
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case gstrSalvar
            Call flSalvar
            If strOperacao = gstrOperAlterar Then
               flPosicionaItemListView
            End If
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

    mdiLQS.uctlogErros.MostrarErros Err, "frmParamHistCntaCntb - tlbCadastro_ButtonClick", Me.Caption

    Call flCarregaListView

    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub

Private Sub txtCentroDestino_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtCentroDestino_KeyPress"
End Sub

Private Sub txtContaCredito_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtContaCredito_KeyPress"
End Sub

Private Sub txtContaDebito_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtContaDebito_KeyPress"
End Sub

Private Sub flCarregarComboTipoOperacao()
Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

       
    cboTipoOperacao.Clear

    
    'Grupo_TipoOperacao[TP_OPER='" & pstrTipoOperacao & "']"
    For Each objNode In xmlMapaNavegacao.selectNodes("//Repeat_TipoOperacao/Grupo_TipoOperacao[TP_MESG_RECB_INTE='154']")
        If Not objNode Is Nothing Then
            If CLng(objNode.selectSingleNode("TP_MESG_RECB_INTE").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
                cboTipoOperacao.AddItem objNode.selectSingleNode("TP_OPER").Text & " - " & _
                                    objNode.selectSingleNode("NO_TIPO_OPER").Text
           End If
        End If
    Next
    
    cboTipoOperacao.AddItem "< Padrão >", 0
    cboTipoOperacao.ListIndex = 0
    cboTipoOperacao.Enabled = cboTipoOperacao.ListCount > 1
    
    Exit Sub



ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flCarregarComboTipoOperacao"
End Sub



'Carrega o conteúdo dos combos
Private Sub flCarregaComboSistema()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    cboSistema.Clear
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa)
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/@Operacao").Text = "LerTodos"
    
    Call xmlLerTodos.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema").xml, _
                                             vntCodErro, _
                                             vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If Not xmlLerTodos.xml = Empty Then
    
        For Each xmlDomNode In xmlLerTodos.selectSingleNode("//Repeat_Sistema").childNodes
            
            With xmlDomNode
                cboSistema.AddItem Trim(.selectSingleNode("SG_SIST").Text) & " - " & .selectSingleNode("NO_SIST").Text
            End With
        Next
    
    End If

    Set xmlLerTodos = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0
    
End Sub

