VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlteracaoAgendamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Agendamento"
   ClientHeight    =   2985
   ClientLeft      =   5340
   ClientTop       =   3495
   ClientWidth     =   3855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3855
   Begin VB.Frame fraTipoAgendamento 
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
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtComando 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtData 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtGradeInicio 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtGradeFim 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   59047939
         UpDown          =   -1  'True
         CurrentDate     =   37886
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número Comando"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   405
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   1845
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Término"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1485
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Grade"
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
         Left            =   150
         TabIndex        =   7
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Início"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1125
         Width           =   1275
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
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
      Left            =   120
      Top             =   2400
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
            Picture         =   "frmAlteracaoAgendamento.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlteracaoAgendamento.frx":031A
            Key             =   "Salvar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAlteracaoAgendamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
' Este componente possibilita ao usuário a alteração do agendamento de mensagens.

Private intTipoAgendamento                  As enumTipoAgendamento
Private intStatusOperacao                   As enumStatusOperacao
Private intStatusMensagem                   As enumStatusMensagem
Private vntSequenciaOperacao                As Variant
Private strNumeroControleIF                 As String
Private strDTRegistroMensagemSPB            As String
Private strDHUltimaAtualizacao              As String
Private strComando                          As String
Private strDataOperacaoMensagem             As String
Private strHoraAgendamento                  As String
Private strCodigoMensagem                   As String
Private lngCodigoMensagemXML                As Long
Private lngLocalLiquidacao                  As Long

'' Atribui/retorna o Tipo de Agendamento
Public Property Let TipoAgendamento(ByVal pintTipoAgendamento As enumTipoAgendamento)
    intTipoAgendamento = pintTipoAgendamento
End Property

'' Atribui/retorna o Status da operação
Public Property Let StatusOperacao(ByVal pintStatusOperacao As enumStatusOperacao)
    intStatusOperacao = pintStatusOperacao
End Property

'' Atribui/retorna o Status da Mensagem
Public Property Let StatusMensagem(ByVal pintStatusMensagem As enumStatusMensagem)
    intStatusMensagem = pintStatusMensagem
End Property

'' Atribui/retorna a sequência da operação
Public Property Let SequenciaOperacao(ByVal pvntSequenciaOperacao As Variant)
    vntSequenciaOperacao = pvntSequenciaOperacao
End Property

'' Atribui/retorna o Número de Controle IF
Public Property Let NumeroControleIF(ByVal pstrNumeroControleIF As String)
    strNumeroControleIF = pstrNumeroControleIF
End Property

'' Atribui/retorna a data do registro da mensagem SPB
Public Property Let DTRegistroMensagemSPB(ByVal pstrDTRegistroMensagemSPB As String)
    strDTRegistroMensagemSPB = pstrDTRegistroMensagemSPB
End Property

'' Atribui/retorna a data da última atualização da operação/mensagem
Public Property Let DHUltimaAtualizacao(ByVal pstrDHUltimaAtualizacao As String)
    strDHUltimaAtualizacao = pstrDHUltimaAtualizacao
End Property

'' Atribui/retorna o número do comando
Public Property Let Comando(ByVal pstrComando As String)
    strComando = pstrComando
End Property

'' Atribui/retorna a data da Operação/Mensagem
Public Property Let DataOperacaoMensagem(ByVal pstrDataOperacaoMensagem As String)
    strDataOperacaoMensagem = pstrDataOperacaoMensagem
End Property

'' Atribui/retorna a Hora do agendamento
Public Property Let HoraAgendamento(ByVal pstrHoraAgendamento As String)
    strHoraAgendamento = pstrHoraAgendamento
End Property

'' Atribui/retorna o Código da mensagem
Public Property Let CodigoMensagem(ByVal pstrCodigoMensagem As String)
    strCodigoMensagem = pstrCodigoMensagem
End Property

'' Atribui/retorna o XML do Código da mensagem
Public Property Let CodigoMensagemXML(ByVal plngCodigoMensagemXML As Long)
    lngCodigoMensagemXML = plngCodigoMensagemXML
End Property

'' Atribui/retorna o Local de liquidação
Public Property Let LocalLiquidacao(ByVal plngLocalLiquidacao As Long)
    lngLocalLiquidacao = plngLocalLiquidacao
End Property

Private Sub dtpHora_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Activate()
    dtpHora.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"
End Sub

Private Sub Form_Load()

#If EnableSoap = 1 Then
    Dim objGradeHorario    As MSSOAPLib30.SoapClient30
#Else
    Dim objGradeHorario    As A8MIU.clsGradeHorario
#End If

Dim xmlDomGradeHorario     As MSXML2.DOMDocument40
Dim xmlDomLeitura          As MSXML2.DOMDocument40
Dim strRetLeitura          As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

    On Error GoTo ErrorHandler
    
    fgCursor True
    
    fraTipoAgendamento.Caption = IIf(intTipoAgendamento = enumTipoAgendamento.Operacao, "Operação", "Mensagem")
    
    txtComando.Text = strComando
    txtData.Text = strDataOperacaoMensagem
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomGradeHorario = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomGradeHorario, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlDomGradeHorario, "Repeat_Filtros", "Grupo_GradeHorario", "")
    Call fgAppendNode(xmlDomGradeHorario, "Grupo_GradeHorario", "CodigoMensagem", strCodigoMensagem)
    Call fgAppendNode(xmlDomGradeHorario, "Grupo_GradeHorario", "CodigoMensagemXML", lngCodigoMensagemXML)
    Call fgAppendNode(xmlDomGradeHorario, "Grupo_GradeHorario", "LocalLiquidacao", lngLocalLiquidacao)
    '>>> -------------------------------------------------------------------------------------------

    Set objGradeHorario = fgCriarObjetoMIU("A8MIU.clsGradeHorario")
    strRetLeitura = objGradeHorario.ObterGradeHorario(xmlDomGradeHorario.xml, vntCodErro, vntMensagemErro)
        
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objGradeHorario = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call xmlDomLeitura.loadXML(strRetLeitura)
        
        txtGradeInicio.Text = Format(fgDtHrStr_To_DateTime(xmlDomLeitura.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioAbertura").Text), "HH:MM")
        txtGradeFim.Text = Format(fgDtHrStr_To_DateTime(xmlDomLeitura.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioEncerramento").Text), "HH:MM")
        
        Set xmlDomLeitura = Nothing
    End If
    
    dtpHora.value = IIf(strHoraAgendamento = vbNullString, gstrDataVazia, strHoraAgendamento)
    
    Set xmlDomGradeHorario = Nothing
    
    Call fgCursor(False)
    
    Exit Sub
    
ErrorHandler:
    Call fgCursor(False)
    
    Set objGradeHorario = Nothing
    Set xmlDomGradeHorario = Nothing
    Set xmlDomLeitura = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoAgendamento - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmAlteracaoAgendamento = Nothing

End Sub

'' Altera o horário de agendamento da mensagem/operação através de interação com a
'' camada de controle de caso de uso MIU, utilizando os métodos conforme
'' necessário:   A8MIU.clsOperacao.AlterarAgendamento   A8MIU.clsMensagem.
'' AlterarAgendamento
Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

#If EnableSoap = 1 Then
    Dim objOperacao        As MSSOAPLib30.SoapClient30
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao        As A8MIU.clsOperacao
    Dim objMensagem        As A8MIU.clsMensagem
#End If

Dim strxmlAgendamento      As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case gstrSalvar
            strxmlAgendamento = flMontaXMLAgendamento(intTipoAgendamento)
            
            If intTipoAgendamento = enumTipoAgendamento.Operacao Then
                Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
                Call objOperacao.AlterarAgendamento(strxmlAgendamento, vntCodErro, vntMensagemErro)
                Set objOperacao = Nothing
            ElseIf intTipoAgendamento = enumTipoAgendamento.MENSAGEM Then
                Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
                Call objMensagem.AlterarAgendamento(strxmlAgendamento, vntCodErro, vntMensagemErro)
                Set objMensagem = Nothing
            End If
            
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If
            
            MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    End Select
    
    fgCursor
    
    Unload Me
    
    Exit Sub

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    fgCursor
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmAlteracaoAgendamento - tlbCadastro_ButtonClick", Me.Caption
    
    Unload Me

End Sub

'' Retorna uma String contendo o XML de entrada para a camada de controle de caso
'' de uso MIU
Private Function flMontaXMLAgendamento(ByVal pintTipoAgendamento As enumTipoAgendamento) As String

Dim xmlDomAgendamento                       As MSXML2.DOMDocument40
Dim strHorarioAgendamento                   As String

On Error GoTo ErrorHandler

    If dtpHora.value = gstrDataVazia Then
        strHorarioAgendamento = ""
    Else
        strHorarioAgendamento = Format(fgDataHoraServidor(Data), "YYYYMMDD") & _
                                Format(dtpHora.Hour, "0#") & _
                                Format(dtpHora.Minute, "0#") & _
                                "00"
    End If
    
    '>>> Formata XML Filtro padrão ---------------------------------------------------------------------------
    Set xmlDomAgendamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomAgendamento, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlDomAgendamento, "Repeat_Filtros", "Grupo_Agendamento", "")
    
    If pintTipoAgendamento = enumTipoAgendamento.Operacao Then
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "Operacao", vntSequenciaOperacao)
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "LocalLiquidacao", lngLocalLiquidacao)
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "StatusOperacao", intStatusOperacao)
    ElseIf pintTipoAgendamento = enumTipoAgendamento.MENSAGEM Then
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "NumeroControleIF", strNumeroControleIF)
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "DTRegistroMensagemSPB", strDTRegistroMensagemSPB)
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "CodigoMensagemXML", lngCodigoMensagemXML)
        Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "StatusMensagem", intStatusMensagem)
    End If
    
    Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "CodigoMensagem", strCodigoMensagem)
    Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "HorarioAgendamento", strHorarioAgendamento)
    Call fgAppendNode(xmlDomAgendamento, "Grupo_Agendamento", "DHUltimaAtualizacao", strDHUltimaAtualizacao)
    '>>> -----------------------------------------------------------------------------------------------------

    flMontaXMLAgendamento = xmlDomAgendamento.xml
    
    Set xmlDomAgendamento = Nothing

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMontaXMLAgendamento", 0

End Function

