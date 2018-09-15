VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComplementoMensagensConciliacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada Manual - Dados Complementares Mensagem"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8685
   Begin VB.Frame fraMensagem 
      Caption         =   "Dados da Mensagem "
      Height          =   3495
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   8505
      Begin VB.TextBox txtISPBIF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1785
         Width           =   8265
      End
      Begin VB.TextBox txtNumeroComando 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3045
         Width           =   8265
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2415
         Width           =   8265
      End
      Begin VB.TextBox txtContraparte 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1155
         Width           =   8265
      End
      Begin VB.TextBox txtVeiculoLegal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   525
         Width           =   8265
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "ISPB IF Creditada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1530
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Número de Comando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2820
         Width           =   1770
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2190
         Width           =   450
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Banco Liquidante Contraparte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   930
         Width           =   2550
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.Frame fraSpread 
      Height          =   2310
      Left            =   75
      TabIndex        =   5
      Top             =   3660
      Width           =   8520
      Begin VB.OptionButton optPagamento 
         Caption         =   "Pagamento via STR0007"
         Height          =   195
         Index           =   1
         Left            =   2670
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optPagamento 
         Caption         =   "Pagamento via STR0004"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2355
      End
      Begin FPSpread.vaSpread vasDadosComplementares 
         Height          =   1875
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   8280
         _Version        =   196608
         _ExtentX        =   14605
         _ExtentY        =   3307
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         MaxCols         =   1
         MaxRows         =   1
         NoBorder        =   -1  'True
         ProcessTab      =   -1  'True
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmComplementoMensagensConciliacao.frx":0000
         UnitType        =   2
         ScrollBarTrack  =   3
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   90
      Top             =   5700
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
            Picture         =   "frmComplementoMensagensConciliacao.frx":020E
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":0528
            Key             =   "Padrao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":097A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":0C94
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":0FAE
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":12C8
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":171A
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComplementoMensagensConciliacao.frx":1B6C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   5670
      TabIndex        =   7
      Top             =   6015
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar Parametrização Padrão"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lblComplemento 
      AutoSize        =   -1  'True
      Caption         =   "Atributos em negrito são obrigatórios"
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
      Left            =   1380
      TabIndex        =   8
      Top             =   6090
      Width           =   3120
   End
End
Attribute VB_Name = "frmComplementoMensagensConciliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pelo complemento de mensagens para Concilição,
'' através da camada de controle de caso de uso MIU.
''
Option Explicit

Public objSelectedItem                      As ListItem
Public xmlComplemento                       As MSXML2.DOMDocument40
Public strComboFinalidade                   As String

Private blnContaCorrente                    As Boolean

Private Enum enumTipoPagamento
    ViaSTR0004 = 0
    ViaSTR0007 = 1
End Enum

Private Sub flFormatarSpread()
    
    On Error GoTo ErrorHandler

    'Identifica se o item atual possui ou não dados de Conta Corrente
    blnContaCorrente = IIf(optPagamento(enumTipoPagamento.ViaSTR0004).value, False, True)
    
    With vasDadosComplementares
        .ReDraw = False
        
        .MaxCols = 2
        .MaxRows = 1
        
        .ColWidth(1) = 5000
        .ColWidth(2) = 2900
        
        .SetText 1, 0, "Nome Campos a Complementar"
        .SetText 2, 0, "Complemento"
        
        .EditEnterAction = EditEnterActionDown
        .CursorStyle = CursorStyleArrow
        .TypeMaxEditLen = 200
        
        'Possui dados Conta Corrente << STR0007 >>
        If blnContaCorrente Then
            .MaxRows = 4
            
            .SetText 1, 1, "Agência Creditada"
            .SetText 1, 2, "Conta Creditada"
            .SetText 1, 3, "CNPJ ou CPF Titular"
            .SetText 1, 4, "Nome Titular"
        
        'Não possui dados Conta Corrente << STR0004 >>
        Else
            .MaxRows = 3
            
            .SetText 1, 1, "Número do Documento"
            .SetText 1, 2, "Finalidade"
            .SetText 1, 3, "Histórico da Mensagem"
            
        End If
        
        .BlockMode = True
        
        .Col = 1
        .Row = 1
        .Col2 = 1
        .Row2 = .MaxRows
        .FontBold = True
        .Lock = True
        .Protect = True
        
        .BlockMode = False
    
        .Col = 1
        .Row = 1
        .Action = ActionActiveCell
        
        .ReDraw = True
    End With
        
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flFormatarSpread", 0

End Sub

'Limpar as informações para uma nova inclusão

Private Sub flLimpaCampos()

    On Error GoTo ErrorHandler

    With vasDadosComplementares
        .BlockMode = True
        
        .Col = 2
        .Row = 1
        .Col2 = 2
        .Row2 = .MaxRows
        .Action = ActionClearText
        
        .BlockMode = False
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

'' Salva as alterações efetuadas através da camada controladora de casos de uso
'' MIU, método A8MIU.clsMIU.Executar

Private Function flSalvar() As Boolean

Dim vntTexto                                As Variant
Dim strValidacao                            As String
    
    On Error GoTo ErrorHandler
    
    strValidacao = flValidarCampos
    
    If strValidacao <> vbNullString Then
        frmMural.Display = strValidacao
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        
        flSalvar = False
        Exit Function
    End If
    
    With vasDadosComplementares
        Call fgAppendNode(xmlComplemento, "", "Repeat_Conciliacao", "")
        Call fgAppendNode(xmlComplemento, "Repeat_Conciliacao", "Grupo_Mensagem", "")
        Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "DT_SIST", Format$(fgDataHoraServidor(DataAux), "YYYYMMDD"))
        Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "CO_ISPB_IF_CRED", txtISPBIF.Text)
        Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "VA_OPER_ATIV", fgVlr_To_Xml(txtValor.Text))
        
        If blnContaCorrente Then
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "TP_CNTA_CRED", "CC")
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "TP_PESS_CRED", "J")
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "NU_DOCT", vbNullString)
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "CO_FIND_COBE", vbNullString)
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "DE_HIST_MESG", vbNullString)
            
            .GetText 2, 1, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "CO_AGEN_CRED", vntTexto)
            
            .GetText 2, 2, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "NU_CNTA_CRED", vntTexto)
            
            .GetText 2, 3, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "CO_CNPJ_CPF_CRED_1", vntTexto)
            
            .GetText 2, 4, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "NO_TITU_1", vntTexto)
            
        Else
            .GetText 2, 1, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "NU_DOCT", vntTexto)
            
            .GetText 2, 2, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "CO_FIND_COBE", fgObterCodigoCombo(vntTexto))
            
            .GetText 2, 3, vntTexto
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "DE_HIST_MESG", vntTexto)
            
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "DT_AGND", String$(8, 0))
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "TX_FILLER_STR0004", String$(969, " "))
            Call fgAppendNode(xmlComplemento, "Grupo_Mensagem", "CO_IDEF_TRAF", String$(25, " "))
            
        End If
    End With

    flSalvar = True
    
    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Function

'' Retorna uma String referente a um preenchimento incorreto na interface. Se
'' todos os campos estiverem preenchidos corretamente, retorna vbNullString

Private Function flValidarCampos() As String

Dim vntTexto                                As Variant
Dim lngLinhasSpread                         As Long
Dim arrMensagens()                          As Variant
    
    On Error GoTo ErrorHandler

    With vasDadosComplementares
        If blnContaCorrente Then
            ReDim arrMensagens(1 To 4, 1 To 2)
            
            arrMensagens(1, 1) = 1
            arrMensagens(2, 1) = 2
            arrMensagens(3, 1) = 3
            arrMensagens(4, 1) = 4
            
            arrMensagens(1, 2) = "Favor informar a Agência Creditada."
            arrMensagens(2, 2) = "Favor informar a Conta Creditada."
            arrMensagens(3, 2) = "Favor informar o CNPJ ou CPF Titular."
            arrMensagens(4, 2) = "Favor informar o Nome Titular."
            
        Else
            ReDim arrMensagens(1 To 3, 1 To 2)
            
            arrMensagens(1, 1) = 1
            arrMensagens(2, 1) = 2
            arrMensagens(3, 1) = 3
            
            arrMensagens(1, 2) = "Favor informar o Número do Documento."
            arrMensagens(2, 2) = "Favor selecionar a Finalidade."
            arrMensagens(3, 2) = "Favor informar o Histórico da Mensagem."
            
        End If
            
        For lngLinhasSpread = LBound(arrMensagens()) To UBound(arrMensagens())
        
            .GetText 2, arrMensagens(lngLinhasSpread, 1), vntTexto
            If vntTexto = vbNullString Then
                flValidarCampos = arrMensagens(lngLinhasSpread, 2)
                Exit Function
            End If
        
        Next
    
    End With

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo ErrorHandler
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Me.Icon = mdiLQS.Icon
    fgCenterMe Me
    
    Set xmlComplemento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call flLimpaCampos
    
    If Not objSelectedItem Is Nothing Then
        txtVeiculoLegal.Text = objSelectedItem.Text
        txtContraparte.Text = objSelectedItem.SubItems(2)
        txtISPBIF.Text = Split(objSelectedItem.Tag, "|")(6)
        txtValor.Text = Replace$(objSelectedItem.SubItems(4), "-", vbNullString)
        txtNumeroComando.Text = objSelectedItem.SubItems(3)
    End If
    
    Call flFormatarSpread
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSelectedItem = Nothing
End Sub

Private Sub optPagamento_Click(Index As Integer)
    Call flFormatarSpread
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Salvar"
            If flSalvar Then Unload Me
        
        Case "Limpar"
            Call flLimpaCampos
        
        Case "Sair"
            Unload Me
            
    End Select
        
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbCadastro_ButtonClick", Me.Caption

End Sub

Private Sub vasDadosComplementares_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    On Error GoTo ErrorHandler
    
    With vasDadosComplementares
        .BlockMode = False
        
        .Col = Col
        .Row = Row
        .CellType = CellTypeEdit
        
        .Col = NewCol
        .Row = NewRow
        
        If NewCol = 1 Then
            .Lock = True
            .Protect = True
        
        Else
            .Lock = False
            .Protect = False
            
            If blnContaCorrente Then
                .CellType = CellTypeEdit
                
                Select Case NewRow
                    Case 1
                        .TypeMaxEditLen = 8
                    Case 2, 3, 4
                        .TypeMaxEditLen = 15
                End Select
            
            Else
                Select Case NewRow
                    Case 1
                        .CellType = CellTypeEdit
                        .TypeMaxEditLen = 6
                    Case 2
                        .CellType = CellTypeComboBox
                        .TypeComboBoxMaxDrop = 11
                        .TypeComboBoxList = strComboFinalidade
                    Case 3
                        .CellType = CellTypeEdit
                        .TypeMaxEditLen = 200
                End Select
            
            End If
            
        End If
            
    End With
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - vasDadosComplementares_LeaveCell", Me.Caption

End Sub
