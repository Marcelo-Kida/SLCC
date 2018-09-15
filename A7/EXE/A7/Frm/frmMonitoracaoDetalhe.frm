VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMonitoracaoDetalhe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhe da Mensagem"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11685
   Begin VB.Frame fraDetalhe 
      Caption         =   "Histórico da Mensagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   4530
      Width           =   11370
      Begin RichTextLib.RichTextBox txtMotivo 
         Height          =   1860
         Left            =   5700
         TabIndex        =   23
         Top             =   390
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3281
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMonitoracaoDetalhe.frx":0000
      End
      Begin MSComctlLib.ListView lstHistorico 
         Height          =   1860
         Left            =   75
         TabIndex        =   13
         Top             =   405
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3281
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
            Text            =   "Data"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ocorrência"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Motivo do Cancelamento"
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
         Index           =   8
         Left            =   5715
         TabIndex        =   24
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.Frame fraDetalhe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5490
      Index           =   1
      Left            =   75
      TabIndex        =   9
      Top             =   1470
      Width           =   11565
      Begin RichTextLib.RichTextBox rtbMensagemEntrada 
         Height          =   2625
         Left            =   105
         TabIndex        =   15
         Top             =   390
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   4630
         _Version        =   393217
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMonitoracaoDetalhe.frx":0082
      End
      Begin RichTextLib.RichTextBox rtbMensagemSaida 
         Height          =   2625
         Left            =   5835
         TabIndex        =   16
         Top             =   375
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   4630
         _Version        =   393217
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMonitoracaoDetalhe.frx":0104
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Mensagem Saida"
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
         Index           =   5
         Left            =   5835
         TabIndex        =   11
         Top             =   150
         Width           =   1455
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Mensagem Entrada"
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
         Index           =   9
         Left            =   105
         TabIndex        =   10
         Top             =   150
         Width           =   1635
      End
   End
   Begin VB.Frame fraDetalhe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   -30
      Width           =   11580
      Begin VB.TextBox txtSistemaOrigem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   8580
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1050
         Width           =   2895
      End
      Begin VB.TextBox txtEmpresaOrigem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1050
         Width           =   2700
      End
      Begin VB.TextBox txtIdentificadorMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   450
         Width           =   4830
      End
      Begin VB.TextBox txtPrioridade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   8580
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   450
         Width           =   2865
      End
      Begin VB.TextBox txtNaturezaMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   450
         Width           =   2700
      End
      Begin VB.TextBox txtTipoMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1050
         Width           =   5685
      End
      Begin VB.TextBox txtCodMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   450
         Width           =   795
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Sistema Origem"
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
         Index           =   7
         Left            =   8580
         TabIndex        =   22
         Top             =   810
         Width           =   1320
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Empresa Origem"
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
         Index           =   2
         Left            =   5850
         TabIndex        =   20
         Top             =   795
         Width           =   1380
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Identificador da Mensagem"
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
         Left            =   975
         TabIndex        =   18
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Prioridade Fila Saída"
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
         Index           =   4
         Left            =   8580
         TabIndex        =   7
         Top             =   210
         Width           =   1800
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Mensagem"
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
         Index           =   3
         Left            =   5850
         TabIndex        =   5
         Top             =   225
         Width           =   2010
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Mensagem"
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
         Left            =   120
         TabIndex        =   3
         Top             =   810
         Width           =   1620
      End
      Begin VB.Label lblMonitoracao 
         AutoSize        =   -1  'True
         Caption         =   "Seq."
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   90
      Top             =   6930
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
            Picture         =   "frmMonitoracaoDetalhe.frx":0186
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":0298
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":05B2
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":0904
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":0A16
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":0D30
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":104A
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracaoDetalhe.frx":1364
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandosForm 
      Height          =   330
      Left            =   10890
      TabIndex        =   14
      Top             =   7035
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
Attribute VB_Name = "frmMonitoracaoDetalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFB2ED50251"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Form"
'Objeto responsável pela exibição das informações da mensagem selecionada na tela de monitoração de mensagens do sistema A7.
Option Explicit

Public lngCodigoMensagem                    As Long

'Carregar informações da mensagem selecionada tais como sistemas origem, empresa e mensagem de entrada e saída.
Private Sub flCarregarDetalheMensagem()

#If EnableSoap = 1 Then
    Dim objMonitoracao      As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao      As A7Miu.clsMonitoracao
#End If

Dim objDetalheMensagem      As MSXML2.DOMDocument40
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim objListItem             As ListItem
Dim strDetalhe              As String
Dim strNarurezaMensagem     As String
Dim strQueryXpath           As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    Set objDetalheMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        
    strDetalhe = objMonitoracao.ObterDetalheMensagem(lngCodigoMensagem, _
                                                     vntCodErro, _
                                                     vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strDetalhe <> "" Then
        
        If Not objDetalheMensagem.loadXML(strDetalhe) Then
            Call fgErroLoadXML(objDetalheMensagem, App.EXEName, "frmMonitoracaoDetalhe", "flCarregarDetalheMensagem")
        End If
        
        strQueryXpath = objDetalheMensagem.selectSingleNode("//TP_MESG").Text
        strQueryXpath = "Grupo_TipoMensagem[TP_MESG='" & strQueryXpath & "']/TP_NATZ_MESG"
        strQueryXpath = gxmlTipoMensagem.documentElement.selectSingleNode(strQueryXpath).Text
        
        Select Case strQueryXpath
            Case Is = enumNaturezaMensagem.MensagemConsulta
                strNarurezaMensagem = "Consulta"
            Case Is = enumNaturezaMensagem.MensagemECO
                strNarurezaMensagem = "Eco"
            Case Is = enumNaturezaMensagem.MensagemEnvio
                strNarurezaMensagem = "Envio de Dados"
        End Select
                
        Me.txtCodMensagem.Text = objDetalheMensagem.selectSingleNode("//CO_MESG").Text
        
        strQueryXpath = objDetalheMensagem.selectSingleNode("//TP_MESG").Text
        strQueryXpath = "Grupo_TipoMensagem[TP_MESG='" & strQueryXpath & "']/NO_TIPO_MESG"
        strQueryXpath = gxmlTipoMensagem.documentElement.selectSingleNode(strQueryXpath).Text
        Me.txtTipoMensagem.Text = objDetalheMensagem.selectSingleNode("//TP_MESG").Text & " - " & strQueryXpath
        
        strQueryXpath = objDetalheMensagem.selectSingleNode("//TP_MESG").Text
        strQueryXpath = "Grupo_TipoMensagem[TP_MESG='" & strQueryXpath & "']/CO_PRIO_FILA_SAID_MESG"
        strQueryXpath = gxmlTipoMensagem.documentElement.selectSingleNode(strQueryXpath).Text
        Me.txtPrioridade.Text = strQueryXpath
        
        strQueryXpath = objDetalheMensagem.selectSingleNode("//CO_EMPR_ORIG").Text
        strQueryXpath = "Grupo_Empresa[CO_EMPR='" & strQueryXpath & "']/NO_REDU_EMPR"
        strQueryXpath = gxmlEmpresa.documentElement.selectSingleNode(strQueryXpath).Text
        Me.txtEmpresaOrigem.Text = objDetalheMensagem.selectSingleNode("//CO_EMPR_ORIG").Text & "-" & strQueryXpath
        
        strQueryXpath = fgCompletaString(objDetalheMensagem.selectSingleNode("//SG_SIST_ORIG").Text, " ", 3, False)
        strQueryXpath = "Grupo_Sistema[SG_SIST='" & strQueryXpath & "']/NO_SIST"
        strQueryXpath = gxmlSistema.documentElement.selectSingleNode(strQueryXpath).Text
        Me.txtSistemaOrigem.Text = objDetalheMensagem.selectSingleNode("//SG_SIST_ORIG").Text & "-" & strQueryXpath
        
        Me.txtIdentificadorMensagem.Text = objDetalheMensagem.selectSingleNode("//CO_CMPO_ATRB_IDEF_MESG").Text
        Me.txtNaturezaMensagem.Text = strNarurezaMensagem
        
        Me.rtbMensagemEntrada.Text = objDetalheMensagem.selectSingleNode("//TX_CNTD_ENTR").Text
        Me.rtbMensagemSaida.Text = objDetalheMensagem.selectSingleNode("//TX_CNTD_SAID").Text
    End If
        
    lstHistorico.ListItems.Clear
    
    For Each objDomNode In objDetalheMensagem.selectNodes("//Repeat_SituacaoMensagem/*")
        
        Set objListItem = Me.lstHistorico.ListItems.Add(, , fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_OCOR_MESG").Text))
        
        objListItem.SubItems(1) = objDomNode.selectSingleNode("DE_OCOR_MESG").Text
        objListItem.SubItems(2) = objDomNode.selectSingleNode("DE_ABRV_OCOR_MESG").Text
        objListItem.Tag = objDomNode.selectSingleNode("TX_DTLH_OCOR_ERRO").Text
        
    Next
    
    Set objMonitoracao = Nothing
    Set objDetalheMensagem = Nothing
    
    Exit Sub

ErrorHandler:

    Set objMonitoracao = Nothing
    Set objDetalheMensagem = Nothing

    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmMonitoracao - flInit")
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Call tlbComandosForm_ButtonClick(tlbComandosForm.Buttons(1))
    End If

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    Call fgCenterMe(Me)
    
    Me.Icon = mdiBUS.Icon
    
    Call fgCursor(True)
    
    DoEvents
    
    Call flCarregarDetalheMensagem
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    mdiBUS.uctLogErros.MostrarErros Err, "frmMonitoracaoDetalhe - Form_Load"

End Sub

Private Sub lstHistorico_ItemClick(ByVal Item As MSComctlLib.ListItem)

    txtMotivo.Text = Item.Tag

End Sub

Private Sub tlbComandosForm_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Unload Me

End Sub

