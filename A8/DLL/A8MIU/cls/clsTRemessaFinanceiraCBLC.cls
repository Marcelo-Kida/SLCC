VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTRemessaFinanceiraCBLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes a remessa CBLC

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                          As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Public Function ProcessaRemessaFinanceiraCBLC(ByVal pstrXMLRemessa As String, _
                                              ByVal pblnReprocessar As Boolean) As Boolean

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
    
'    lngContaLinha = 0
'
'    'Para debugar
'    'pstrNomeArquivo = "AMDF20050202100001.txt"
'    'strCaminho = "D:\SLCC\CBLC\Arq\" & pstrNomeArquivo
'    '--------------------------------------------------------------------
'
'    If InStr(1, pstrNomeArquivo, "|") > 0 Then
'        strNomeArquivo = Split(pstrNomeArquivo, "|")(0)
'    Else
'        strNomeArquivo = pstrNomeArquivo
'    End If
'
'    strNomeArquivo = StrReverse$(strNomeArquivo)
'    strNomeArquivo = Mid(strNomeArquivo, 1, InStr(1, strNomeArquivo, "\") - 1)
'    strNomeArquivo = StrReverse$(strNomeArquivo)
'
'    If InStr(1, UCase(strNomeArquivo), ".TXT") > 0 Then
'        strNomeArquivo = Replace(UCase(strNomeArquivo), ".TXT", vbNullString)
'    ElseIf InStr(1, UCase(strNomeArquivo), ".DAT") > 0 Then
'        strNomeArquivo = Replace(UCase(strNomeArquivo), ".DAT", vbNullString)
'    Else
'        strNomeArquivo = Replace(UCase(strNomeArquivo), ".TXT", vbNullString)
'    End If
'
'    strCaminho = Split(pstrNomeArquivo, "|")(0)
'
'    lngArquivo = FreeFile()
'
''    If Dir(strCaminho) = vbNullString Then
''        'Arquivo remessa CBLC não existe
''        lngCodigoErroNegocio = 3124
''        strComplementoErro = "Arquivo:" & strCaminho
''        GoTo ErrorHandler
''    End If
'
'    strDataServidor = fgDt_To_Xml(fgDataHoraServidor(DataAux))
'
'    Set xmlRemessaCBLC = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(xmlRemessaCBLC, "", "REME_CBLC", "")
'
'    'Leitura do Arquivo
'    Open strCaminho For Input As #lngArquivo
'    While Not EOF(lngArquivo)
'
'        blnIncluir = False
'
'        Line Input #lngArquivo, strLinha
'
'        udtAMDF_ALCOHeaderAux.String = strLinha
'        LSet udtAMDF_ALCOHeader = udtAMDF_ALCOHeaderAux
'
'        Select Case udtAMDF_ALCOHeader.TipoRegistro
'            Case "00"
'                'Header
'                udtAMDF_ALCOHeaderAux.String = strLinha
'                LSet udtAMDF_ALCOHeader = udtAMDF_ALCOHeaderAux
'                strIDArquivo = strLinha
'
'                If udtAMDF_ALCOHeader.DataGeracao <> strDataServidor Then
'                    'Data Geração arquivo incompatível
'                    lngCodigoErroNegocio = 3126
'                    strComplementoErro = "Data Arquivo:" & udtAMDF_ALCOHeader.DataGeracao & "<BR> Data Servidor:" & strDataServidor
'                    GoTo ErrorHandler
'                End If
'
'            Case "01"
'                'Lancamento
'
'
'
'                udtAMDF_ALCOLancamentoAux.String = strLinha
'                LSet udtAMDF_ALCOLancamento = udtAMDF_ALCOLancamentoAux
'
'                'Sempre Data D0
'                If udtAMDF_ALCOLancamento.DataEfetivacao = strDataServidor Then
'
'                    'Somente Forma Pagamento 3 (STR) ou Branco
'                    If Val(udtAMDF_ALCOLancamento.FormaPagamento) = 3 Or _
'                       Val(udtAMDF_ALCOLancamento.FormaPagamento) = 0 Then
'                       blnIncluir = True
'                    Else
'                        blnIncluir = False
'                    End If
'
'                Else
'                    blnIncluir = False
'                End If
'
'                If blnIncluir Then
'
'                    lngContaLinha = lngContaLinha + 1
'
'                    Set xmlLancamentoCBLC = CreateObject("MSXML2.DOMDocument.4.0")
'
'                    Call fgAppendNode(xmlLancamentoCBLC, "", "LANC_CBLC", "")
'
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NU_SEQU_ARQU_CAMR", 0)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LOCA_LIQU", enumLocalLiquidacao.CLBCAcoes)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NO_ARQU_CAMR", pstrNomeArquivo)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "TP_REGT", udtAMDF_ALCOLancamento.TipoRegistro)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "DT_EFET_LANC", udtAMDF_ALCOLancamento.DataEfetivacao)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_GRUP_LANC_FINC", Val(udtAMDF_ALCOLancamento.CodigoGrupo))
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LANC_FINC", Val(udtAMDF_ALCOLancamento.CodigoLancamentoFinanceriro))
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_IDEF_LANC", udtAMDF_ALCOLancamento.IdentificacaoLancamento)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_LANC_FINC", udtAMDF_ALCOLancamento.CodigoLancamentoFinanceriro)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_AGET_CPEN", 0)
'
'                    If udtAMDF_ALCOLancamento.LancamentoParaQualificado = "Q" Then
'                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_CLIE_QULF", enumIndicadorSimNao.Sim)
'                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_CLIE_QULF", udtAMDF_ALCOLancamento.CodigoClienteQualificado)
'                    Else
'                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_CLIE_QULF", enumIndicadorSimNao.Nao)
'                        Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_CLIE_QULF", vbNullString)
'                    End If
'
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_COTR", Val(udtAMDF_ALCOLancamento.CodigoCorretora))
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "NO_COTR", udtAMDF_ALCOLancamento.NomeCorretora)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "SG_SIST", vbNullString)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_VEIC_LEGA", vbNullString)
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "SITU_LANC", udtAMDF_ALCOLancamento.SituacaoLancamento)
'
'                    'A - MOVIMENTO DO DIA
'                    'P - PREVISTO
'                    'H - ECLUIDO
'                    If UCase(udtAMDF_ALCOLancamento.SituacaoLancamento) = "H" Then
'                        If UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "D" Then
'                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito)
'                        ElseIf UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "C" Then
'                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito)
'                        End If
'                    Else
'                        If UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "D" Then
'                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito)
'                        ElseIf UCase(udtAMDF_ALCOLancamento.TipoLancamento) = "C" Then
'                            Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito)
'                        End If
'                    End If
'
'                    strValoLancamento = udtAMDF_ALCOLancamento.ValorLancamento
'                    strValorInteiro = Val(Left$(strValoLancamento, 16))
'                    strValorDecimal = Right$(strValoLancamento, 2)
'                    strValoLancamento = strValorInteiro & "," & strValorDecimal
'
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "VA_LANC", strValoLancamento)
'
'                    Select Case udtAMDF_ALCOLancamento.BancoLiquidante
'                        Case enumISPB.IspbBANESPA
'                            lngCodigoEmpresa = enumCodigoEmpresa.Banespa
'                        Case enumISPB.IspbBOZZANO
'                            lngCodigoEmpresa = enumCodigoEmpresa.Bozano
'                        Case enumISPB.IspbMERIDIONAL
'                            lngCodigoEmpresa = enumCodigoEmpresa.Meridional
'                        Case enumISPB.IspbSANTANDER
'                            lngCodigoEmpresa = enumCodigoEmpresa.Santander
'                        Case Else
'                           lngCodigoEmpresa = enumCodigoEmpresa.Santander
'                    End Select
'
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "CO_EMPR", lngCodigoEmpresa)
'
'                    Call fgAppendNode(xmlLancamentoCBLC, "LANC_CBLC", "TX_REME", strLinha)
'
'                    Call fgAppendXML(xmlRemessaCBLC, "REME_CBLC", xmlLancamentoCBLC.xml)
'
'                    Set xmlLancamentoCBLC = Nothing
'                End If
'            Case "99"
'                udtAMDF_ALCOTrailerAux.String = strLinha
'                LSet udtAMDF_ALCOTrailer = udtAMDF_ALCOTrailerAux
'
'                If Val(udtAMDF_ALCOTrailer.TotalRegistros) <> lngContaLinha + 2 Then
'                    'Erro processamento arquivo CBLC
'                    lngCodigoErroNegocio = 3125
'                    strComplementoErro = "Total Arquivo:" & udtAMDF_ALCOTrailer.TotalRegistros & "<BR>  Total Processados:" & lngContaLinha
'                    GoTo ErrorHandler
'                End If
'
'            End Select
'    Wend
'
'    Close #lngArquivo

    Set xmlRemessaCBLC = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlRemessaCBLC.loadXML(pstrXMLRemessa)
    Call flGravarProcessaRemessaFinanceiraCBLC(xmlRemessaCBLC, pblnReprocessar)
    Set xmlRemessaCBLC = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Close #lngArquivo
        
    Set xmlLancamentoCBLC = Nothing
    Set xmlRemessaCBLC = Nothing
    
    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If
    
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessaRemessaFinanceiraCBLC Sub", lngCodigoErroNegocio, intNumeroSequencialErro, strComplementoErro)


End Function

Private Function flGravarProcessaRemessaFinanceiraCBLC(ByVal pxmlRemessaCBLC As MSXML2.DOMDocument40, _
                                                       ByVal pblnReprocessar As Boolean) As Boolean

Dim objRemessaFinanceiraCamara              As A8LQS.clsRemessaFinanceiraCamara
Dim objIdentVeivLega                        As A6A7A8.clsIdPartCamaraVeicLegal

Dim xmlLancamento                           As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strCodVeicLega                          As String
Dim strSgSist                               As String
Dim vntIdentPartCBLC                        As Variant
Dim vntSequenciaArquivo                     As Variant

On Error GoTo ErrorHandler

    Set xmlLancamento = CreateObject("MSXML2.DOMDocument.4.0")
    Set objRemessaFinanceiraCamara = CreateObject("A8LQS.clsRemessaFinanceiraCamara")

    Call objRemessaFinanceiraCamara.IncluirLogRemessa(pxmlRemessaCBLC.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                                      pxmlRemessaCBLC.selectSingleNode("//NO_ARQU_CAMR").Text, _
                                                      vntSequenciaArquivo, _
                                                      pblnReprocessar)

    If pxmlRemessaCBLC.selectSingleNode("//CO_COTR") Is Nothing Then
        GoTo ExitFunction
    End If
    
    'Obter Veiculo Legal
    For Each xmlNode In pxmlRemessaCBLC.selectNodes("//REME_CBLC/LANC_CBLC")
        
        xmlLancamento.loadXML xmlNode.xml

        If xmlNode.selectSingleNode("CO_COTR").Text = enumCodigoCorretora.SantanderCCVM Then
            vntIdentPartCBLC = xmlNode.selectSingleNode("CO_COTR").Text
        Else
            If Trim$(xmlNode.selectSingleNode("CO_CLIE_QULF").Text) = vbNullString Then
                vntIdentPartCBLC = enumCodigoCorretora.SantanderBrasilSA
            Else
                vntIdentPartCBLC = Val(xmlNode.selectSingleNode("CO_CLIE_QULF").Text)
            End If

        End If

        Set objIdentVeivLega = CreateObject("A6A7A8.clsIdPartCamaraVeicLegal")
        Call objIdentVeivLega.ObterVeiculoLegal(vntIdentPartCBLC, _
                                                strCodVeicLega, _
                                                strSgSist, _
                                                enumLocalLiquidacao.CLBCAcoes)

        Set objIdentVeivLega = Nothing

        xmlNode.selectSingleNode("CO_VEIC_LEGA").Text = strCodVeicLega
        xmlNode.selectSingleNode("SG_SIST").Text = strSgSist
        xmlNode.selectSingleNode("NU_SEQU_ARQU_CAMR").Text = vntSequenciaArquivo

    Next

    For Each xmlNode In pxmlRemessaCBLC.selectNodes("//REME_CBLC/LANC_CBLC")
        xmlLancamento.loadXML xmlNode.xml
        Call objRemessaFinanceiraCamara.Incluir(xmlLancamento)
    Next

ExitFunction:
    Call objRemessaFinanceiraCamara.AlterarLogRemessa(pxmlRemessaCBLC.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                                      vntSequenciaArquivo)

    Exit Function

ErrorHandler:
    Set objIdentVeivLega = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGravarProcessaRemessaFinanceiraCBLC Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


