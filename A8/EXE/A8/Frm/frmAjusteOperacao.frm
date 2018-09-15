VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjusteOperacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Valor da Operação"
   ClientHeight    =   3165
   ClientLeft      =   2745
   ClientTop       =   2430
   ClientWidth     =   4305
   Icon            =   "frmAjusteOperacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      Begin VB.TextBox txtJustificativa 
         Height          =   315
         Left            =   180
         MaxLength       =   70
         TabIndex        =   10
         Text            =   "txtJustificativa"
         Top             =   2100
         Width           =   3675
      End
      Begin NumBox.Number numValor 
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         Top             =   1380
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Decimais        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelecao     =   0   'False
      End
      Begin NumBox.Number numValorOriginal 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1380
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Decimais        =   2
         Locked          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483637
         AutoSelecao     =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Justificativa:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1860
         Width           =   870
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   1020
         TabIndex        =   4
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   1020
         TabIndex        =   2
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Operação:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor Original:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor &Ajustado:"
         Height          =   195
         Left            =   2100
         TabIndex        =   7
         Top             =   1140
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Height          =   330
      Left            =   1140
      TabIndex        =   9
      Top             =   2820
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      ButtonWidth     =   2461
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar            "
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   540
      Top             =   2700
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
            Picture         =   "frmAjusteOperacao.frx":0442
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":0554
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":0666
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":09B8
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":0D0A
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":105C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":13AE
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":16C8
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjusteOperacao.frx":1B1A
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAjusteOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
' Este componente tem como objetivo possibilitar ao usuário o ajuste do valor de uma operação.

#If EnableSoap = 1 Then
    Private objOperacao                 As MSSOAPLib30.SoapClient30
#Else
    Private objOperacao                 As A8MIU.clsOperacao
#End If

Private vntNU_SEQU_OPER_ATIV            As Variant
Private vntDH_ULTI_ATLZ                 As Variant

Public blnSalvou                        As Boolean  'se salvou a alteraçao

Private Sub Form_Activate()

    numValor.SetFocus
    txtJustificativa.Text = ""
    fgCursor

End Sub

Private Sub Form_Initialize()

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objOperacao = Nothing
    Set frmAjusteOperacao = Nothing
End Sub

'' Salva o valor alterado da operação no banco de dados.
Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strRetorno              As String
Dim strErro                 As String
Dim strMensagemConfirmacao  As String
Dim strJustificativa        As String

Dim intAcao                 As enumAcaoConciliacao
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    fgCursor True
    
    blnSalvou = False

    Select Case Button.Key
        Case "Salvar"
            
            strJustificativa = Trim(txtJustificativa.Text)
            If strJustificativa = vbNullString Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Informe uma Justificativa para a alteração do valor da operação."
                frmMural.Show vbModal
            Else
                If objOperacao.AjustarValor(vntNU_SEQU_OPER_ATIV, vntDH_ULTI_ATLZ, numValor.Valor, strJustificativa, vntCodErro, vntMensagemErro) Then
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                    
                    blnSalvou = True
                    Me.Hide
                End If
            End If
            
        Case gstrSair
            Me.Hide
    
    End Select
    fgCursor
    Exit Sub

ErrorHandler:
    fgCursor
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmAjusteOperacao - tlbFiltro_ButtonClick", Me.Caption

End Sub

'' Armazena o número da operação de que está sendo ajustada
''
Public Property Let SequenciaOperacao(ByVal pSeqOperacao As Variant)
    
Dim strRet              As String
Dim xmlOp               As MSXML2.DOMDocument40
Dim xmlFiltro           As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    vntNU_SEQU_OPER_ATIV = pSeqOperacao
    
    'busca detalhes da operação
    Set xmlOp = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlFiltro.loadXML ""
    Call fgAppendNode(xmlFiltro, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_NumeroOperacao", "")
    Call fgAppendNode(xmlFiltro, "Grupo_NumeroOperacao", "Numero", vntNU_SEQU_OPER_ATIV, "Repeat_Filtros")
    
    strRet = objOperacao.ObterDetalheOperacao(xmlFiltro.xml, _
                                              vntCodErro, _
                                              vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strRet = vbNullString Then
        'Operação não encontrada
        tlbFiltro.Buttons("Salvar").Enabled = False
    
    Else
        xmlOp.loadXML strRet
    
        vntDH_ULTI_ATLZ = xmlOp.selectSingleNode("//DH_ULTI_ATLZ").Text
        numValor.Valor = Val(Replace(xmlOp.selectSingleNode("//VA_OPER_ATIV").Text, ",", "."))
        
        If Val(Replace(xmlOp.selectSingleNode("//VA_OPER_ATIV_REAJ").Text, ",", ".")) = 0 Then
            numValorOriginal.Valor = Val(Replace(xmlOp.selectSingleNode("//VA_OPER_ATIV").Text, ",", "."))
        Else
            numValorOriginal.Valor = Val(Replace(xmlOp.selectSingleNode("//VA_OPER_ATIV_REAJ").Text, ",", "."))
        End If
        
        lblCodigo.Caption = xmlOp.selectSingleNode("//CO_OPER_ATIV").Text
        lblData.Caption = fgDtXML_To_Interface(xmlOp.selectSingleNode("//DT_OPER_ATIV").Text)
    End If
    
    Set xmlOp = Nothing
    Set xmlFiltro = Nothing
   
ErrorHandler:
   
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Property

