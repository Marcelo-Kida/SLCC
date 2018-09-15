VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacaoArquivoCBLC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ferramentas - Importação Arquivo CBLC"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7980
   Begin MSComDlg.CommonDialog cdlgArquivoCBLC 
      Left            =   6360
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   90
      TabIndex        =   3
      Top             =   4320
      Width           =   7815
      Begin VB.CommandButton cmdSelecaoArquivo 
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
         Height          =   255
         Left            =   6930
         TabIndex        =   5
         Top             =   390
         Width           =   480
      End
      Begin VB.TextBox txtNomeArquivo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivos Importados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   7815
      Begin MSComctlLib.ListView lstLogImportacao 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome Arquivo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data Importação"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuário"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      ButtonWidth     =   2884
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar Arquivo"
            Key             =   "importar"
            ImageKey        =   "reprocessar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh            "
            Key             =   "refresh"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r                   "
            Key             =   "sair"
            ImageKey        =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   7080
      Top             =   5160
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
            Picture         =   "frmImportacaoArquivoCBLC.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":1286
            Key             =   "reprocessar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivoCBLC.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmImportacaoArquivoCBLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngCodigoErroNegocio                As Long
Private strComplemento                      As String

Private Sub cmdSelecaoArquivo_Click()

On Error GoTo ErrorHandler
    
    cdlgArquivoCBLC.DialogTitle = "Selecione o arquivo CBLC"
    cdlgArquivoCBLC.FileName = ""
    cdlgArquivoCBLC.Filter = "*.txt"
    cdlgArquivoCBLC.Action = 1
        
    If cdlgArquivoCBLC.FileTitle <> "" Then
        txtNomeArquivo.Text = cdlgArquivoCBLC.FileName
    End If
      
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmImportacaoArquivoCBLC - cmdSelecaoArquivo_Click", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    Call flCarregaHistoricoImpoertacao
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmImportacaoArquivoCBLC - Form_Load", Me.Caption

End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Select Case Button.Key
        Case "importar"
            Call flImportarArquivo
            
        Case "refresh"
            Call flCarregaHistoricoImpoertacao
        Case "sair"
            Unload Me
    End Select
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub

Private Function flImportarArquivo()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsRemessaFinanceiraCBLC
#End If

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim blnReprocessar                          As Boolean
Dim xmlRemessa                              As String
Dim xmlValidacaoRemessa                     As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
        
    If Trim$(txtNomeArquivo) = vbNullString Then
        MsgBox "Selecione um arquivo", vbCritical, Me.Caption
        cmdSelecaoArquivo.SetFocus
        Exit Function
    End If
        
    xmlRemessa = ProcessaRemessaFinanceiraCBLC(txtNomeArquivo.Text)
    
    If xmlRemessa = vbNullString Then Exit Function
    
    Set xmlValidacaoRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    xmlValidacaoRemessa.loadXML xmlRemessa
    
    If Val(xmlValidacaoRemessa.selectSingleNode("//TOTAL_LANCAMENTOS").Text) = 0 Then
        If MsgBox("O arquivo selecionado não possui registros de lançamentos a serem importados." & vbNewLine & "Deseja prosseguir com a importação mesmo assim ?", vbYesNo + vbQuestion, "Importação Arquivo CBLC") = vbNo Then
            Exit Function
        End If
    End If
    
    Call fgCursor(True)
    DoEvents
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraCBLC")
    
    vntCodErro = 0
    blnReprocessar = objMIU.VerificaArquivoProcessado(txtNomeArquivo, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then GoTo ErrorHandler
    
    If blnReprocessar Then
        
        If MsgBox("O arquivo selecionado já foi importado em " & fgDataHoraServidor(DataAux) & "." & vbNewLine & "Deseja sobrepor importação anterior?", vbYesNo + vbQuestion, "Importação Arquivo CBLC") = vbYes Then
            
            vntCodErro = 0
            
            Set objMIU = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraCBLC")
            Call objMIU.ProcessaRemessaFinanceiraCBLC(xmlRemessa, blnReprocessar, vntCodErro, vntMensagemErro)
            Set objMIU = Nothing
            
            If vntCodErro <> 0 Then GoTo ErrorHandler
            
        End If
    
    Else
        
        vntCodErro = 0
        
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraCBLC")
        Call objMIU.ProcessaRemessaFinanceiraCBLC(xmlRemessa, blnReprocessar, vntCodErro, vntMensagemErro)
        Set objMIU = Nothing
    
        If vntCodErro <> 0 Then GoTo ErrorHandler
    
    End If
    
    Call flCarregaHistoricoImpoertacao
    
    Set objMIU = Nothing
    Call fgCursor(False)
    
    MsgBox "Arquivo importado com sucesso !", vbInformation, Me.Caption
    
    Exit Function

ErrorHandler:
    Call fgCursor(False)
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Function

' Carregar grid com histórico de importação arquivo CBLC - D0
Private Sub flCarregaHistoricoImpoertacao()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsRemessaFinanceiraCBLC
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim strLeitura                              As String
Dim objListItem                             As MSComctlLib.ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

    On Error GoTo ErrorHandler
        
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsRemessaFinanceiraCBLC")
    strLeitura = objMIU.LerTodosLogRemessa(enumLocalLiquidacao.CLBCAcoes, _
                                           vntCodErro, _
                                           vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    lstLogImportacao.ListItems.Clear
    
    If xmlLerTodos.loadXML(strLeitura) Then
        For Each objDomNode In xmlLerTodos.documentElement.childNodes
            Set objListItem = lstLogImportacao.ListItems.Add(, , objDomNode.selectSingleNode("NO_ARQU_CAMR").Text)
            objListItem.SubItems(1) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
            objListItem.SubItems(2) = objDomNode.selectSingleNode("CO_USUA_ULTI_ATLZ").Text
        Next objDomNode
    End If

    lstLogImportacao.Refresh
    Set xmlLerTodos = Nothing
    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flPreencherHistorico", 0

End Sub

Private Function ProcessaRemessaFinanceiraCBLC(ByVal pstrNomeArquivo As String) As String

Dim udtAMDF_ALCOHeader                      As udtAMDF_ALCOHeader
Dim udtAMDF_ALCOLancamento                  As udtAMDF_ALCOLancamento
Dim udtAMDF_ALCOTrailer                     As udtAMDF_ALCOTrailer

Dim udtAMDF_ALCOHeaderAux                   As udtAMDF_ALCOHeaderAux
Dim udtAMDF_ALCOLancamentoAux               As udtAMDF_ALCOLancamentoAux
Dim udtAMDF_ALCOTrailerAux                  As udtAMDF_ALCOTrailerAux

Dim xmlRemessaCBLC                          As MSXML2.DOMDocument40
Dim xmlLancamentoCBLC                       As MSXML2.DOMDocument40

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strCaminho                              As String
Dim lngArquivo                              As Integer
Dim strLinha                                As String
Dim lngContaLinha                           As Long
Dim strValoLancamento                       As String
Dim strValorInteiro                         As String
Dim strValorDecimal                         As String
Dim lngCodigoEmpresa                        As Long
Dim strDataServidor                         As String
Dim strIDArquivo                            As String
Dim strComplementoErro                      As String
Dim strNomeArquivo                          As String

Dim blnIncluir                              As Boolean

    On Error GoTo ErrorHandler
    
    lngContaLinha = 0
    
    If InStr(1, pstrNomeArquivo, "|") > 0 Then
        strNomeArquivo = Split(pstrNomeArquivo, "|")(0)
    Else
        strNomeArquivo = pstrNomeArquivo
    End If
    
    strNomeArquivo = StrReverse$(strNomeArquivo)
    strNomeArquivo = Mid(strNomeArquivo, 1, InStr(1, strNomeArquivo, "\") - 1)
    strNomeArquivo = StrReverse$(strNomeArquivo)
    
    If InStr(1, UCase(strNomeArquivo), ".TXT") > 0 Then
        strNomeArquivo = Replace(UCase(strNomeArquivo), ".TXT", vbNullString)
    ElseIf InStr(1, UCase(strNomeArquivo), ".DAT") > 0 Then
        strNomeArquivo = Replace(UCase(strNomeArquivo), ".DAT", vbNullString)
    Else
        strNomeArquivo = Replace(UCase(strNomeArquivo), ".TXT", vbNullString)
    End If
    
    strCaminho = Split(pstrNomeArquivo, "|")(0)
    
    lngArquivo = FreeFile()
              
    If Dir(strCaminho) = vbNullString Then
        'Arquivo remessa CBLC não existe
        'lngCodigoErroNegocio = 3124
        strComplementoErro = "Arquivo remessa CBLC não existe" & vbCrLf & strComplementoErro
        frmMural.Display = strComplementoErro
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Function
    End If
            
    strDataServidor = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    
    Set xmlRemessaCBLC = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlRemessaCBLC, "", "REME_CBLC", "")
        
    'Leitura do Arquivo
    Open strCaminho For Input As #lngArquivo
    While Not EOF(lngArquivo)
            
        blnIncluir = False
        
        Line Input #lngArquivo, strLinha
                        
        udtAMDF_ALCOHeaderAux.String = strLinha
        LSet udtAMDF_ALCOHeader = udtAMDF_ALCOHeaderAux
                        
        Select Case udtAMDF_ALCOHeader.TipoRegistro
            Case "00"
                
                'Header
                udtAMDF_ALCOHeaderAux.String = strLinha
                LSet udtAMDF_ALCOHeader = udtAMDF_ALCOHeaderAux
                strIDArquivo = strLinha
                
                If udtAMDF_ALCOHeader.DataGeracao <> strDataServidor Then
                    'Data Geração arquivo incompatível
                    'lngCodigoErroNegocio = 3126
                    strComplementoErro = "Data Geração arquivo incompatível" & vbCrLf & vbCrLf
                    strComplementoErro = strComplementoErro & "Data Sistema : " & fgDtXML_To_Interface(strDataServidor) & vbCrLf
                    strComplementoErro = strComplementoErro & "Data Geração arquivo : " & fgDtXML_To_Interface(udtAMDF_ALCOHeader.DataGeracao) & vbCrLf
                    
                    ProcessaRemessaFinanceiraCBLC = vbNullString
                    Close #lngArquivo
                    frmMural.Display = strComplementoErro
                    frmMural.IconeExibicao = IconExclamation
                    frmMural.Show vbModal
                    Exit Function
                End If
            
                Set xmlLancamentoCBLC = CreateObject("MSXML2.DOMDocument.4.0")

                Call fgAppendNode(xmlLancamentoCBLC, "", "LANC_CBLC", "")

                Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NU_SEQU_ARQU_CAMR", 0)
                Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LOCA_LIQU", enumLocalLiquidacao.CLBCAcoes)
                Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NO_ARQU_CAMR", pstrNomeArquivo)
            
            Case "01"
                'Lancamento
                
                udtAMDF_ALCOLancamentoAux.String = strLinha
                LSet udtAMDF_ALCOLancamento = udtAMDF_ALCOLancamentoAux
                 
                'Sempre Data D0
                If udtAMDF_ALCOLancamento.DataEfetivacao = strDataServidor Then
                    
                    'Somente Forma Pagamento 3 (STR) ou Branco
                    If Val(udtAMDF_ALCOLancamento.FormaPagamento) = 3 Or _
                       Val(udtAMDF_ALCOLancamento.FormaPagamento) = 0 Then
                       blnIncluir = True
                    Else
                        blnIncluir = False
                    End If
                       
                Else
                    blnIncluir = False
                End If

                If blnIncluir Then

                    lngContaLinha = lngContaLinha + 1
                    
                    If lngContaLinha > 1 Then

                        Set xmlLancamentoCBLC = CreateObject("MSXML2.DOMDocument.4.0")
    
                        Call fgAppendNode(xmlLancamentoCBLC, "", "LANC_CBLC", "")
    
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NU_SEQU_ARQU_CAMR", 0)
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LOCA_LIQU", enumLocalLiquidacao.CLBCAcoes)
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NO_ARQU_CAMR", pstrNomeArquivo)
                        
                    End If
                    
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "TP_REGT", udtAMDF_ALCOLancamento.TipoRegistro)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "DT_EFET_LANC", udtAMDF_ALCOLancamento.DataEfetivacao)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_GRUP_LANC_FINC", Val(udtAMDF_ALCOLancamento.CodigoGrupo))
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LANC_FINC", Val(udtAMDF_ALCOLancamento.CodigoLancamentoFinanceriro))
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_IDEF_LANC", udtAMDF_ALCOLancamento.IdentificacaoLancamento)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LANC_FINC", udtAMDF_ALCOLancamento.CodigoLancamentoFinanceriro)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_AGET_CPEN", 0)

                    If udtAMDF_ALCOLancamento.LancamentoParaQualificado = "Q" Then
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_CLIE_QULF", enumIndicadorSimNao.Sim)
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_CLIE_QULF", udtAMDF_ALCOLancamento.CodigoClienteQualificado)
                    Else
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_CLIE_QULF", enumIndicadorSimNao.Nao)
                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_CLIE_QULF", vbNullString)
                    End If

                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_COTR", Val(udtAMDF_ALCOLancamento.CodigoCorretora))
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NO_COTR", udtAMDF_ALCOLancamento.NomeCorretora)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "SG_SIST", vbNullString)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_VEIC_LEGA", vbNullString)
                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "SITU_LANC", udtAMDF_ALCOLancamento.SituacaoLancamento)

                    'A - MOVIMENTO DO DIA
                    'P - PREVISTO
                    'H - ECLUIDO
                    If UCase(udtAMDF_ALCOLancamento.SituacaoLancamento) = "H" Then
                        If UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "D" Then
                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito)
                        ElseIf UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "C" Then
                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito)
                        End If
                    Else
                        If UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "D" Then
                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito)
                        ElseIf UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "C" Then
                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito)
                        End If
                    End If

                    strValoLancamento = udtAMDF_ALCOLancamento.ValorLancamento
                    strValorInteiro = Val(Left$(strValoLancamento, 16))
                    strValorDecimal = Right$(strValoLancamento, 2)
                    strValoLancamento = strValorInteiro & "," & strValorDecimal

                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "VA_LANC", strValoLancamento)

                    Select Case udtAMDF_ALCOLancamento.BancoLiquidante
                        Case enumISPB.IspbBANESPA
                            lngCodigoEmpresa = enumCodigoEmpresa.Banespa
                        Case enumISPB.IspbBOZZANO
                            lngCodigoEmpresa = enumCodigoEmpresa.Bozano
                        Case enumISPB.IspbMERIDIONAL
                            lngCodigoEmpresa = enumCodigoEmpresa.Meridional
                        Case enumISPB.IspbSANTANDER
                            lngCodigoEmpresa = enumCodigoEmpresa.Santander
                        Case Else
                           lngCodigoEmpresa = enumCodigoEmpresa.Santander
                    End Select

                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_EMPR", lngCodigoEmpresa)

                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "TX_REME", strLinha)

                    Call fgAppendXML(xmlRemessaCBLC, "REME_CBLC", xmlLancamentoCBLC.xml)

                    Set xmlLancamentoCBLC = Nothing
                
                End If
            
            Case "99"
                udtAMDF_ALCOTrailerAux.String = strLinha
                LSet udtAMDF_ALCOTrailer = udtAMDF_ALCOTrailerAux

                If Val(udtAMDF_ALCOTrailer.TotalRegistros) <> lngContaLinha + 2 Then
                    'Erro processamento arquivo CBLC
                    lngCodigoErroNegocio = 3125
                    
                    strComplementoErro = "Total Arquivo:" & udtAMDF_ALCOTrailer.TotalRegistros & vbCrLf & "Total Processados:" & lngContaLinha
                    
                    strComplementoErro = "Erro processamento arquivo CBLC " & vbCrLf & strComplementoErro
                    ProcessaRemessaFinanceiraCBLC = vbNullString
                    Close #lngArquivo
                    frmMural.Display = strComplementoErro
                    frmMural.IconeExibicao = IconExclamation
                    frmMural.Show vbModal
                    Exit Function
                End If

                If lngContaLinha = 0 Then
                    Call fgAppendXML(xmlRemessaCBLC, "REME_CBLC", xmlLancamentoCBLC.xml)
                    Set xmlLancamentoCBLC = Nothing
                End If
                
                Call fgAppendNode(xmlRemessaCBLC, "REME_CBLC", "TOTAL_LANCAMENTOS", lngContaLinha)
        
        End Select
    Wend

    Close #lngArquivo

    ProcessaRemessaFinanceiraCBLC = xmlRemessaCBLC.xml

    Set xmlRemessaCBLC = Nothing

    Exit Function

ErrorHandler:
    Close #lngArquivo
        
    Set xmlLancamentoCBLC = Nothing
    Set xmlRemessaCBLC = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "ProcessaRemessaFinanceiraCBLC", 0
    
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessaRemessaFinanceiraCBLC Sub", lngCodigoErroNegocio, intNumeroSequencialErro, strComplementoErro)

End Function
