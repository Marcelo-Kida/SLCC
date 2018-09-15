VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlteracaoTipoCompromisso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de Tipo de Compromisso"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboTipoCompromisso 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   5595
   End
   Begin VB.TextBox txtComando 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   5595
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   5220
      TabIndex        =   4
      Top             =   840
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoTipoCompromisso.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoTipoCompromisso.frx":031A
            Key             =   "Salvar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Compromisso"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Número do comando"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "frmAlteracaoTipoCompromisso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Permite que o tipo de compromisso da operação seja alterado através de
'' interação com a camada controladora de caso de uso A8MIU
''
'' São consideradas especificamente classes de destino:
''   A8MIU.clsOperacao
''

Option Explicit

Private strDataUltimaAtualizacao            As String
Private lngNumeroSequencia                  As Long
Private strComando                          As String
Private strTipoCompromisso                  As String
Private strXmlTipoCompromisso               As String
Private lngTipoOperacao                     As Long
Private strAcaoAnterior                     As String
Private strStatusOperacao                   As String

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"
End Sub

Private Sub Form_Load()

Dim xmlDomTipoCompromisso                   As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCenterMe Me
    fgCursor True
    Set Me.Icon = mdiLQS.Icon
    txtComando.Text = strComando
    
    Set xmlDomTipoCompromisso = CreateObject("MSXML2.DOMDocument.4.0")
    xmlDomTipoCompromisso.loadXML strXmlTipoCompromisso
    fgCarregarCombos cboTipoCompromisso, xmlDomTipoCompromisso, "DominioAtributo", "CO_DOMI", "DE_DOMI"
    flPosicionaComboCompromisso
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    Set xmlDomTipoCompromisso = Nothing
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoTipoCompromisso - Form_Load", Me.Caption

End Sub

'' Posiciona o combo de tipo de compromisso de acordo com o tipo de compromisso da
'' operação
Private Sub flPosicionaComboCompromisso()
    
Dim intCont                                 As Integer
Dim blnEncontrou                            As Boolean
    
On Error GoTo ErrorHandler

    For intCont = 0 To cboTipoCompromisso.ListCount - 1
        cboTipoCompromisso.ListIndex = intCont
        If fgObterCodigoCombo(cboTipoCompromisso.Text) = strTipoCompromisso Then
            blnEncontrou = True
            Exit For
        End If
    Next intCont
    
    If Not blnEncontrou Then
        cboTipoCompromisso.ListIndex = -1
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaComboCompromisso", 0

End Sub

'' Altera o tipo de compromisso através do método flSalvar ou fecha o objeto
Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
    
        Case gstrSalvar
            flSalvar
        
        Case gstrSair
            Unload Me
    
    End Select
    fgCursor
Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoTipoCompromisso - tlbCadastro_ButtonClick", Me.Caption
End Sub

'' Altera o tipo de compromisso da operação através da camada de controle de caso
'' de uso MIU, método:   A8MIU.clsOperacao.AlterarTipoCompromisso
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objOperacao             As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao             As A8MIU.clsOperacao
#End If

Dim xmlFiltro                   As MSXML2.DOMDocument40
Dim strNovoTipoCompromisso      As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    If cboTipoCompromisso.ListIndex < 0 Then
        Exit Sub
    End If

    strNovoTipoCompromisso = fgObterCodigoCombo(cboTipoCompromisso.Text)
    
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    fgAppendNode xmlFiltro, "", "Filtro", ""
    fgAppendNode xmlFiltro, "Filtro", "NU_SEQU_OPER_ATIV", lngNumeroSequencia
    fgAppendNode xmlFiltro, "Filtro", "DH_ULTI_ATLZ", strDataUltimaAtualizacao
    fgAppendNode xmlFiltro, "Filtro", "TP_CPRO_OPER_ATIV", IIf(lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaIda, strNovoTipoCompromisso, "")
    fgAppendNode xmlFiltro, "Filtro", "TP_CPRO_RETN_OPER_ATIV", IIf(lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaVolta, strNovoTipoCompromisso, "")
    fgAppendNode xmlFiltro, "Filtro", "TX_CNTD_ANTE_ACAO", strAcaoAnterior
    fgAppendNode xmlFiltro, "Filtro", "CO_ULTI_SITU_PROC", strStatusOperacao
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    objOperacao.AlterarTipoCompromisso xmlFiltro.xml, vntCodErro, vntMensagemErro
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objOperacao = Nothing
    Set xmlFiltro = Nothing
    
    Unload Me

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

Public Property Get DataUltimaAtualizacao() As String
    DataUltimaAtualizacao = strDataUltimaAtualizacao
End Property

Public Property Let DataUltimaAtualizacao(ByVal NewValue As String)
    strDataUltimaAtualizacao = NewValue
End Property

Public Property Get NumeroSequencia() As Long
    NumeroSequencia = lngNumeroSequencia
End Property

Public Property Let NumeroSequencia(ByVal NewValue As Long)
    lngNumeroSequencia = NewValue
End Property

Public Property Get Comando() As String
    Comando = strComando
End Property

Public Property Let Comando(ByVal NewValue As String)
    strComando = NewValue
End Property

Public Property Get TipoCompromisso() As String
    TipoCompromisso = strTipoCompromisso
End Property

Public Property Let TipoCompromisso(ByVal NewValue As String)
    strTipoCompromisso = NewValue
End Property

Public Property Get XmlTipoCompromisso() As String
    XmlTipoCompromisso = strXmlTipoCompromisso
End Property

Public Property Let XmlTipoCompromisso(ByVal NewValue As String)
    strXmlTipoCompromisso = NewValue
End Property

Public Property Get TipoOperacao() As Long
    TipoOperacao = lngTipoOperacao
End Property

Public Property Let TipoOperacao(ByVal NewValue As Long)
    lngTipoOperacao = NewValue
End Property

Public Property Get AcaoAnterior() As String
    AcaoAnterior = strAcaoAnterior
End Property

Public Property Let AcaoAnterior(ByVal NewValue As String)
    strAcaoAnterior = NewValue
End Property

Public Property Get StatusOperacao() As String
    StatusOperacao = strStatusOperacao
End Property

Public Property Let StatusOperacao(ByVal NewValue As String)
    strStatusOperacao = NewValue
End Property
