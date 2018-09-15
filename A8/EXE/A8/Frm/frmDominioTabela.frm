VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDominioTabela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleção Domínio Entrada Manual"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSelecao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3960
      Left            =   90
      TabIndex        =   4
      Top             =   75
      Width           =   6885
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   150
         TabIndex        =   0
         Top             =   495
         Width           =   6570
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   1095
         Width           =   6570
      End
      Begin MSComctlLib.ListView lstDominio 
         Height          =   2400
         Left            =   120
         TabIndex        =   2
         Top             =   1470
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   4233
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
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
            Text            =   "Sistema"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Veículo Legal"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblCod 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   840
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3525
      TabIndex        =   3
      Top             =   4065
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   582
      ButtonWidth     =   2143
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesquisar"
            Key             =   "Atualizar"
            Object.ToolTipText     =   "Pesquisar"
            ImageKey        =   "Atualizar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Selecionar"
            Key             =   "Selecionar"
            Object.ToolTipText     =   "Selecionar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   60
      Top             =   3915
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
            Picture         =   "frmDominioTabela.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":5922
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":5F56
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDominioTabela.frx":6270
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDominioTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Objeto responsavel pela exibição dos dominios para entrada manual,
' através da camada de controle de caso de uso MIU.
'
Public lngCodigoEmpresa                     As Long
Public strNomeTabela                        As String
Public strMensagem                          As String
Public blnCancel                            As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
    
    If KeyAscii = 27 Then
    
        tlbCadastro_ButtonClick tlbCadastro.Buttons("Sair")
        
    End If
    
Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Select Case UCase(strNomeTabela)
        
        Case "PJ.TB_PRODUTO"
            fraSelecao.Caption = "Produto"
            
        Case "PJ.TB_SEGMENTO"
            fraSelecao.Caption = "Segmento"
            flAtualizar
            
        Case "PJ.TB_EVENTO_FINANCEIRO"
            fraSelecao.Caption = "Evento Financeiro"
            flAtualizar
            
        Case "PJ.TB_INDEXADOR"
            fraSelecao.Caption = "Indexador"
            flAtualizar
            
        Case "PJ.TB_LOCAL_LIQUIDACAO"
            fraSelecao.Caption = "Local de Liquidação"
            flAtualizar
            
        Case "A8.TB_VEIC_LEGA"
            fraSelecao.Caption = "Veículo Legal"
            
        Case "PJ.TB_TIPO_CONTA"
            fraSelecao.Caption = "Tipo Conta"
            flAtualizar
            
        Case "A8.TB_MESG_RECB_ENVI_SPB"
            fraSelecao.Caption = "Mensagem BMC0112 - Demonstrativo de custos "
            
            lstDominio.ColumnHeaders(1).Width = 2000
            lstDominio.ColumnHeaders(2).Width = 2000
            lstDominio.ColumnHeaders(3).Width = 1000
            lstDominio.ColumnHeaders(4).Width = 1500
                        
            lstDominio.ColumnHeaders(1).Text = "Num Controle BMC"
            lstDominio.ColumnHeaders(2).Text = "Veículo Legal"
            lstDominio.ColumnHeaders(3).Text = "Valor"
            lstDominio.ColumnHeaders(4).Text = "Data Hora Mensagem"
            
            lblCod.Visible = False
            lblDesc.Visible = False
            txtCodigo.Visible = False
            txtDescricao.Visible = False
            
            lstDominio.Top = 210
            lstDominio.Height = lstDominio.Height + 950
            
            flAtualizar
            
        Case "CHACAM"
            fraSelecao.Caption = "Mensagem " & strMensagem
            
            lstDominio.ColumnHeaders(1).Width = 1500
            lstDominio.ColumnHeaders(2).Width = 1500
            lstDominio.ColumnHeaders(3).Width = 2000
            lstDominio.ColumnHeaders(4).Width = 1500
            
            lstDominio.ColumnHeaders(1).Text = "Chave Assoc Câmbio"
            lstDominio.ColumnHeaders(2).Text = "Valor Moeda Nacional"
            lstDominio.ColumnHeaders(3).Text = "Número Controle IF"
            lstDominio.ColumnHeaders(4).Text = "Data Hora Mensagem"
            
            lblCod.Visible = False
            lblDesc.Visible = False
            txtCodigo.Visible = False
            txtDescricao.Visible = False
            
            lstDominio.Top = 210
            lstDominio.Height = lstDominio.Height + 950
            
            flAtualizar
            
        Case "CO_REG_OPER_CAMB"
            fraSelecao.Caption = "Mensagem " & strMensagem
            
            lstDominio.ColumnHeaders(1).Width = 1500
            lstDominio.ColumnHeaders(2).Width = 1500
            lstDominio.ColumnHeaders(3).Width = 2000
            lstDominio.ColumnHeaders(4).Width = 1500
            
            lstDominio.ColumnHeaders(1).Text = "Cód. Reg. Oper. Câmbio"
            lstDominio.ColumnHeaders(2).Text = "Código Mensagem"
            lstDominio.ColumnHeaders(3).Text = "Número Controle IF"
            lstDominio.ColumnHeaders(4).Text = "Data Hora Mensagem"
            
            lblCod.Visible = False
            lblDesc.Visible = False
            txtCodigo.Visible = False
            txtDescricao.Visible = False
            
            lstDominio.Top = 210
            lstDominio.Height = lstDominio.Height + 950
            
            flAtualizar
            
        Case "CO_REG_OPER_CAMB2"
            fraSelecao.Caption = "Mensagem " & strMensagem
            
            lstDominio.ColumnHeaders(1).Width = 1500
            lstDominio.ColumnHeaders(2).Width = 1500
            lstDominio.ColumnHeaders(3).Width = 2000
            lstDominio.ColumnHeaders(4).Width = 1500
            
            lstDominio.ColumnHeaders(1).Text = "Cód. Reg. Oper. Câmbio"
            lstDominio.ColumnHeaders(2).Text = "Cód. Reg. Oper. Câmbio2"
            lstDominio.ColumnHeaders(3).Text = "Número Controle IF"
            lstDominio.ColumnHeaders(4).Text = "Data Hora Mensagem"
            
            lblCod.Visible = False
            lblDesc.Visible = False
            txtCodigo.Visible = False
            txtDescricao.Visible = False
            
            lstDominio.Top = 210
            lstDominio.Height = lstDominio.Height + 950
            
            flAtualizar
            
    End Select

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnCancel = True
End Sub

Private Sub lstDominio_DblClick()
    
    Me.Hide
    
    blnCancel = False
    
End Sub

Private Sub lstDominio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            
        tlbCadastro_ButtonClick tlbCadastro.Buttons("Selecionar")
    
    End If

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        
        Case gstrAtualizar
            flAtualizar
            blnCancel = False
        Case "Selecionar"
            Me.Hide
            blnCancel = False
        Case gstrSair
            Me.Hide
            blnCancel = True
    End Select
    fgCursor
    
Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmDominioTabela - tlbCadastro_ButtonClick", Me.Caption
    
End Sub

' Carrega os domínios já existente e preenche a interface com os mesmos,
' através da classe controladora de caso de uso MIU, método A8MIU.clsMensagem.LerTodosDominioTabela

Private Sub flAtualizar()

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim xmlMensagem             As MSXML2.DOMDocument40
Dim xmlDominio              As MSXML2.DOMDocument40
Dim xmlNode                 As MSXML2.IXMLDOMNode
Dim objListItem             As ListItem
Dim strDominio              As String
Dim strMensagemXML          As String
Dim strKey                  As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")

    lstDominio.ListItems.Clear

    strDominio = objMensagem.LerTodosDominioTabela(lngCodigoEmpresa, _
                                                   strNomeTabela, _
                                                   txtCodigo.Text, _
                                                   Trim(txtDescricao.Text), _
                                                   strMensagem, _
                                                   vntCodErro, _
                                                   vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If Trim(strDominio) = vbNullString Then Exit Sub
    
    If strNomeTabela <> "A8.TB_VEIC_LEGA" _
    And strNomeTabela <> "A8.TB_MESG_RECB_ENVI_SPB" _
    And strNomeTabela <> "ChACAM" _
    And strNomeTabela <> "CO_REG_OPER_CAMB" _
    And strNomeTabela <> "CO_REG_OPER_CAMB2" Then
        lstDominio.ColumnHeaders(3).Text = ""
    End If
    
    Set xmlDominio = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlDominio.loadXML(strDominio) Then
        fgErroLoadXML xmlDominio, App.EXEName, "frmDominioTabela", "flAtualizar"
    End If

    For Each xmlNode In xmlDominio.documentElement.childNodes
        
        If strNomeTabela = "A8.TB_VEIC_LEGA" Then
            
            Set objListItem = lstDominio.ListItems.Add(, "K" & xmlNode.selectSingleNode("CODIGO").Text & xmlNode.selectSingleNode("SG_SIST").Text, xmlNode.selectSingleNode("CODIGO").Text)
            
            objListItem.SubItems(1) = xmlNode.selectSingleNode("DESCRICAO").Text
        
        ElseIf strNomeTabela = "A8.TB_MESG_RECB_ENVI_SPB" Then
            
            strKey = "K" & xmlNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                           xmlNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & "|" & _
                           xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text

            Set objListItem = lstDominio.ListItems.Add(, strKey, xmlNode.selectSingleNode("CONTROLE").Text)
            objListItem.SubItems(1) = xmlNode.selectSingleNode("CO_VEIC_LEGA").Text & " - " & xmlNode.selectSingleNode("NO_VEIC_LEGA").Text
            objListItem.SubItems(2) = xmlNode.selectSingleNode("VALOR").Text
            objListItem.SubItems(3) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
            
        ElseIf strNomeTabela = "ChACAM" Then
            
            strKey = "K" & xmlNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                           xmlNode.selectSingleNode("CD_ASSO_CAMB").Text & "|" & _
                           xmlNode.selectSingleNode("DT_OPER").Text & "|" & _
                           xmlNode.selectSingleNode("IN_OPER_DEBT_CRED").Text & "|" & _
                           xmlNode.selectSingleNode("PE_TAXA_NEGO").Text & "|" & _
                           xmlNode.selectSingleNode("VA_FINC").Text & "|" & _
                           xmlNode.selectSingleNode("VA_MOED_ESTR").Text & "|" & _
                           xmlNode.selectSingleNode("DT_LIQU").Text

            Set objListItem = lstDominio.ListItems.Add(, strKey, xmlNode.selectSingleNode("CD_ASSO_CAMB").Text)
            objListItem.SubItems(1) = FormatCurrency(CDbl(xmlNode.selectSingleNode("VA_FINC").Text), 2)
            objListItem.SubItems(2) = xmlNode.selectSingleNode("NU_CTRL_IF").Text
            objListItem.SubItems(3) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
            
        ElseIf strNomeTabela = "CO_REG_OPER_CAMB" Then
        
            strMensagemXML = objMensagem.ObterXMLMensagem(CLng(0 + xmlNode.selectSingleNode("CO_TEXT_XML").Text), vntCodErro, vntMensagemErro)
            
            If strMensagemXML <> "" Then
                Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
                xmlMensagem.loadXML strMensagemXML
            End If
            
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If
            
            strKey = "K" & xmlNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                           xmlNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & "|" & _
                           xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text & "|" & _
                           xmlMensagem.selectSingleNode("//CNPJBaseIF").Text & "|" & _
                           xmlMensagem.selectSingleNode("//CodMoedaISO").Text & "|" & _
                           xmlMensagem.selectSingleNode("//VlrME").Text & "|" & _
                           xmlMensagem.selectSingleNode("//TaxCam").Text & "|" & _
                           xmlMensagem.selectSingleNode("//VlrMN").Text & "|" & _
                           xmlMensagem.selectSingleNode("//DtEntrMN").Text & "|" & _
                           xmlMensagem.selectSingleNode("//DtEntrME").Text & "|" & _
                           xmlMensagem.selectSingleNode("//DtLiquid").Text

            Set objListItem = lstDominio.ListItems.Add(, strKey, xmlNode.selectSingleNode("NU_COMD_OPER").Text)
            objListItem.SubItems(1) = xmlNode.selectSingleNode("CO_MESG_SPB").Text
            objListItem.SubItems(2) = xmlNode.selectSingleNode("NU_CTRL_IF").Text
            objListItem.SubItems(3) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
            
        ElseIf strNomeTabela = "CO_REG_OPER_CAMB2" Then
            
            strKey = "K" & xmlNode.selectSingleNode("NU_CTRL_IF").Text & "|" & _
                           xmlNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & "|" & _
                           xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text

            Set objListItem = lstDominio.ListItems.Add(, strKey, xmlNode.selectSingleNode("NU_COMD_OPER").Text)
            objListItem.SubItems(1) = xmlNode.selectSingleNode("NR_OPER_CAMB_2").Text
            objListItem.SubItems(2) = xmlNode.selectSingleNode("NU_CTRL_IF").Text
            objListItem.SubItems(3) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
            
        Else
            Set objListItem = lstDominio.ListItems.Add(, "K" & xmlNode.selectSingleNode("CODIGO").Text, xmlNode.selectSingleNode("CODIGO").Text)
            objListItem.SubItems(1) = xmlNode.selectSingleNode("DESCRICAO").Text
        End If
        
        If Not xmlNode.selectSingleNode("SG_SIST") Is Nothing Then
            objListItem.Tag = Format(xmlNode.selectSingleNode("SG_SIST").Text, "@@@") & xmlNode.selectSingleNode("IN_SIST_SITU_CNTG").Text
            objListItem.SubItems(2) = Format(xmlNode.selectSingleNode("SG_SIST").Text, "@@@") & " - " & xmlNode.selectSingleNode("NO_SIST").Text
        End If
    Next

    Set objMensagem = Nothing
    Set xmlDominio = Nothing

Exit Sub
ErrorHandler:
    Set xmlDominio = Nothing
    Set objMensagem = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmDominioTabela", "flAtualizar", 0
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        tlbCadastro_ButtonClick tlbCadastro.Buttons("Atualizar")
    
    End If

End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        tlbCadastro_ButtonClick tlbCadastro.Buttons("Atualizar")
    
    End If

End Sub
