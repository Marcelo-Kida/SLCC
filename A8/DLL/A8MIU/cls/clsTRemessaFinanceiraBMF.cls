VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTRemessaFinanceiraBMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes a remessa BMF

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

'Executar o processamento de gravação do arquivo da BMF
Public Function ProcessaRemessaFinanceiraBMF(ByVal pstrNomeArquivo As String) As Boolean

Dim udtT025_Header                          As udtT025_Header
Dim udtT025_Lancamento                      As udtT025_Lancamento

Dim udtT025_HeaderAux                       As udtT025_HeaderAux
Dim udtT025_LancamentoAux                   As udtT025_LancamentoAux

Dim xmlRemessaBMF                           As MSXML2.DOMDocument40
Dim xmlLancamentoBMF                        As MSXML2.DOMDocument40

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

    'Para debugar
    'pstrNomeArquivo = "AMDF20050202100001.txt"
    'strCaminho = "D:\SLCC\BMF\Arq\" & pstrNomeArquivo
    '--------------------------------------------------------------------

    strNomeArquivo = StrReverse$(pstrNomeArquivo)
    strNomeArquivo = Mid(strNomeArquivo, 1, InStr(1, strNomeArquivo, "\") - 1)
    strNomeArquivo = StrReverse$(strNomeArquivo)
    strNomeArquivo = Replace(UCase(strNomeArquivo), ".TXT", vbNullString)

    strCaminho = pstrNomeArquivo

    lngArquivo = FreeFile()

    If Dir(strCaminho) = vbNullString Then
        'Arquivo remessa BMF não existe.
        lngCodigoErroNegocio = 3129
        strComplementoErro = "Arquivo:" & strCaminho
        GoTo ErrorHandler
    End If

    strDataServidor = fgDt_To_Xml(fgDataHoraServidor(DataAux))

    Set xmlRemessaBMF = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlRemessaBMF, "", "REME_BMF", "")

    'Leitura do Arquivo
    Open strCaminho For Input As #lngArquivo
    While Not EOF(lngArquivo)

        blnIncluir = False

        Line Input #lngArquivo, strLinha

        udtT025_HeaderAux.String = strLinha
        LSet udtT025_Header = udtT025_HeaderAux

        Select Case udtT025_Header.TipoRegistro
            Case "01"
                'Header
                udtT025_HeaderAux.String = strLinha
                LSet udtT025_Header = udtT025_HeaderAux

                If udtT025_Header.DataNetting <> strDataServidor Then
                    'Data Geração arquivo incompatível
                    lngCodigoErroNegocio = 3126
                    strComplementoErro = "Data Arquivo:" & udtT025_Header.DataGeracao & "<BR> Data Servidor:" & strDataServidor
                    GoTo ErrorHandler
                End If

            Case "02"
                'Lancamento

                udtT025_LancamentoAux.String = strLinha
                LSet udtT025_Lancamento = udtT025_LancamentoAux

                'Somente Código de Lançamento 006, 007, 018, 038, 097
                If Val(udtT025_Lancamento.CodigoLancamento) = 6 Or _
                   Val(udtT025_Lancamento.CodigoLancamento) = 7 Or _
                   Val(udtT025_Lancamento.CodigoLancamento) = 18 Or _
                   Val(udtT025_Lancamento.CodigoLancamento) = 38 Or _
                   Val(udtT025_Lancamento.CodigoLancamento) = 97 Then
                   blnIncluir = True
                Else
                    blnIncluir = False
                End If

                If blnIncluir Then

                    lngContaLinha = lngContaLinha + 1

                    Set xmlLancamentoBMF = CreateObject("MSXML2.DOMDocument.4.0")

                    Call fgAppendNode(xmlLancamentoBMF, "", "LANC_BMF", "")

                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "NU_SEQU_ARQU_CAMR", 0)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_LOCA_LIQU", enumLocalLiquidacao.BMD)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "NO_ARQU_CAMR", strNomeArquivo)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "TP_REGT", udtT025_Lancamento.TipoRegistro)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "DT_EFET_LANC", udtT025_Header.DataNetting)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_GRUP_LANC_FINC", Val(udtT025_Lancamento.CodigoLancamento))
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_AGET_CPEN", 0)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_IDEF_LANC", udtT025_Lancamento.IdentificaoTransacao & udtT025_Lancamento.ComplementoTransacao)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "IN_CLIE_QULF", enumIndicadorSimNao.sim)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_CLIE_QULF", udtT025_Header.CodigoCliente)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_COTR", Val(udtT025_Header.CodigoCorretora))
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "NO_COTR", "")
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "SG_SIST", vbNullString)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_VEIC_LEGA", vbNullString)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_LANC_FINC", 0)
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_TEXT_XML", 0)

                    If udtT025_Lancamento.DebitoCredito = "C" Then
                        Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito)
                    Else
                        Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito)
                    End If

                    strValoLancamento = udtT025_Lancamento.ValorLancamento
                    strValorInteiro = Val(Left$(strValoLancamento, 13))
                    strValorDecimal = Right$(strValoLancamento, 2)
                    strValoLancamento = strValorInteiro & "," & strValorDecimal

                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "VA_LANC", strValoLancamento)

                    'Carlos - Ver com Andrea o codigo da empresa
                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "CO_EMPR", 523)

                    Call fgAppendNode(xmlLancamentoBMF, "LANC_BMF", "TX_REME", strLinha)

                    Call fgAppendXML(xmlRemessaBMF, "REME_BMF", xmlLancamentoBMF.xml)

                    Set xmlLancamentoBMF = Nothing
                End If

            End Select
    Wend

    Close #lngArquivo

    Call flGravarProcessaRemessaFinanceirabmf(xmlRemessaBMF)

    Set xmlRemessaBMF = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Close #lngArquivo

    Set xmlLancamentoBMF = Nothing
    Set xmlRemessaBMF = Nothing

    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessaRemessaFinanceiraBMF Sub", lngCodigoErroNegocio, intNumeroSequencialErro, strComplementoErro)

End Function

'Executar a chamada da gravação do arquivo da BMF
Private Function flGravarProcessaRemessaFinanceirabmf(ByVal pxmlRemessaBMF As MSXML2.DOMDocument40) As Boolean

Dim objRemessaFinanceiraCamara              As A8LQS.clsRemessaFinanceiraCamara
Dim objIdentVeivLega                        As A6A7A8.clsIdPartCamaraVeicLegal

Dim xmlLancamento                           As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strCodVeicLega                          As String
Dim strSgSist                               As String
Dim vntIdentPartCamr                        As Variant
Dim vntSequenciaArquivo                     As Variant

On Error GoTo ErrorHandler

    Set xmlLancamento = CreateObject("MSXML2.DOMDocument.4.0")
    Set objRemessaFinanceiraCamara = CreateObject("A8LQS.clsRemessaFinanceiraCamara")

    Call objRemessaFinanceiraCamara.IncluirLogRemessa(pxmlRemessaBMF.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                                      pxmlRemessaBMF.selectSingleNode("//NO_ARQU_CAMR").Text, _
                                                      vntSequenciaArquivo)

    'Obter Veiculo Legal
    For Each xmlNode In pxmlRemessaBMF.selectNodes("//REME_BMF/LANC_BMF")
        xmlLancamento.loadXML xmlNode.xml

        If xmlNode.selectSingleNode("CO_COTR").Text = enumCodigoCorretora.SantanderCCVM Then
            vntIdentPartCamr = xmlNode.selectSingleNode("CO_COTR").Text
        Else
            If Trim$(xmlNode.selectSingleNode("CO_CLIE_QULF").Text) = vbNullString Then
                vntIdentPartCamr = enumCodigoCorretora.SantanderBrasilSA
            Else
                vntIdentPartCamr = Val(xmlNode.selectSingleNode("CO_CLIE_QULF").Text)
            End If
        End If

        Set objIdentVeivLega = CreateObject("A6A7A8.clsIdPartCamaraVeicLegal")
        Call objIdentVeivLega.ObterVeiculoLegal(vntIdentPartCamr, _
                                                strCodVeicLega, _
                                                strSgSist, _
                                                enumLocalLiquidacao.CLBCAcoes)

        Set objIdentVeivLega = Nothing

        xmlNode.selectSingleNode("CO_VEIC_LEGA").Text = strCodVeicLega
        xmlNode.selectSingleNode("SG_SIST").Text = strSgSist
        xmlNode.selectSingleNode("NU_SEQU_ARQU_CAMR").Text = vntSequenciaArquivo

    Next

    For Each xmlNode In pxmlRemessaBMF.selectNodes("//REME_BMF/LANC_BMF")
        xmlLancamento.loadXML xmlNode.xml
        Call objRemessaFinanceiraCamara.Incluir(xmlLancamento)
    Next

    Call objRemessaFinanceiraCamara.AlterarLogRemessa(pxmlRemessaBMF.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                                      vntSequenciaArquivo)

    Exit Function
ErrorHandler:

    Set objIdentVeivLega = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGravarProcessaRemessaFinanceirabmf Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
