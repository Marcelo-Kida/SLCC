VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo gerenciamento do envio de informações para outros sistemas, como PJ, A6, Legados, etc.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private ObjectContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

'Método genérico para a execução de outros métodos desta classe
Public Function Executar(ByRef pxmlExecucao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlExecucao.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodosFluxoAtividade"
            Executar = flLerTodosFluxoAtividade(pxmlExecucao)
        Case "ExcluirFluxoAtividade"
            Executar = flExcluirFluxoAtividade(pxmlExecucao)
        Case "IncluirFluxoAtividade"
            Executar = flIncluirFluxoAtividade(pxmlExecucao)
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Exclui registros da tabela de fluxo de processamento conforme filtro informado
Private Function flExcluirFluxoAtividade(ByRef pxmlExecucao As MSXML2.DOMDocument40) As String

Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereFiltro                          As String
Dim strValores                              As String

    On Error GoTo ErrorHandler

    strValores = vbNullString
    For Each objXMLNode In pxmlExecucao.selectNodes("//Grupo_LocalLiquidacao/*")
        strValores = strValores & objXMLNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND CO_LOCA_LIQU IN (" & strValores & ")" & vbNewLine
    End If
    
    strSQL = " DELETE                       " & vbNewLine & _
             " FROM    A8.TB_CTRL_FLUX_ATIV " & vbNewLine
    
    If strWhereFiltro <> vbNullString Then
        strWhereFiltro = " WHERE" & strWhereFiltro
        strWhereFiltro = Replace$(strWhereFiltro, "WHERE AND", "WHERE")
        strSQL = strSQL & strWhereFiltro
    End If
    
    Call fgExecuteSQL(strSQL)
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    flExcluirFluxoAtividade = vbNullString
    
    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluirFluxoAtividade Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Ler tabela TB_CTRL_FLUX_ATIV conforme filtro informado
Private Function flLerTodosFluxoAtividade(ByVal pxmlFiltro As MSXML2.DOMDocument40) As String

Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereFiltro                          As String
Dim strValores                              As String

    On Error GoTo ErrorHandler
    
    strValores = vbNullString
    For Each objXMLNode In pxmlFiltro.selectNodes("//Grupo_LocalLiquidacao/*")
        strValores = strValores & objXMLNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND A.CO_LOCA_LIQU IN (" & strValores & ")" & vbNewLine
    End If
    
    strSQL = " SELECT  A.CO_MESG_SPB,           " & vbNewLine & _
             "         A.TP_OPER                " & vbNewLine & _
             " FROM    A8.TB_CTRL_FLUX_ATIV A   " & vbNewLine

    If strWhereFiltro <> vbNullString Then
        strWhereFiltro = " WHERE" & strWhereFiltro
        strWhereFiltro = Replace$(strWhereFiltro, "WHERE AND", "WHERE")
        strSQL = strSQL & strWhereFiltro
    End If
    
    strSQL = strSQL & " ORDER BY A.CO_MESG_SPB, A.TP_OPER "
                 
    flLerTodosFluxoAtividade = fgQueryXMLLerTodos("FluxoAtividade", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLerTodosFluxoAtividade Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Inclui registros da tabela de fluxo de processamento conforme filtro informado
Private Function flIncluirFluxoAtividade(ByRef pxmlExecucao As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String
Dim strCodigoMensagem                       As String
Dim strTipoOperacao                         As String

    On Error GoTo ErrorHandler
    
    With pxmlExecucao
        If Not .selectSingleNode("//CO_MESG_SPB") Is Nothing Then
            strCodigoMensagem = "'" & .selectSingleNode("//CO_MESG_SPB").Text & "'"
        Else
            strCodigoMensagem = "NULL"
        End If
    
        If Not .selectSingleNode("//TP_OPER") Is Nothing Then
            strTipoOperacao = .selectSingleNode("//TP_OPER").Text
        Else
            strTipoOperacao = "NULL"
        End If
    
        strSQL = " INSERT INTO A8.TB_CTRL_FLUX_ATIV (                 " & vbNewLine & _
                 "      CO_MESG_SPB,                                  " & vbNewLine & _
                 "      CO_SITU_PROC,                                 " & vbNewLine & _
                 "      CO_EMPR_FUSI,                                 " & vbNewLine & _
                 "      CO_LOCA_LIQU,                                 " & vbNewLine & _
                 "      TP_LIQU_OPER_ATIV,                            " & vbNewLine & _
                 "      TP_OPER                                       " & vbNewLine & _
                 " ) VALUES (                                         " & vbNewLine & _
                 "  " & strCodigoMensagem & ",                        " & vbNewLine & _
                 "  " & enumStatusOperacao.Inativa & ",               " & vbNewLine & _
                 "  1,                                                " & vbNewLine & _
                 "  " & .selectSingleNode("//CO_LOCA_LIQU").Text & ", " & vbNewLine & _
                 "  " & enumTipoLiquidacao.SemModalidade & ",         " & vbNewLine & _
                 "  " & strTipoOperacao & ")"
    End With
    
    Call fgExecuteSQL(strSQL)
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    flIncluirFluxoAtividade = vbNullString
    
    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirFluxoAtividade Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar a chamada dos métodos de envio de informações para lançamento de conta corrente, PJ, A6, mensagem SPB e retorno para o Legado.
Public Function GerenciarChamadas(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                  ByVal pstrFuncionalidade As String, _
                         Optional ByRef xmlRepeticao As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objLegado                               As A8LQS.clsLegado
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim strMensagem                             As String
Dim lngTipoLiquidacao                       As Long
Dim lngTipoOperacao                         As Long
Dim lngTipoSolicitacao                      As Long

    On Error GoTo ErrorHandler

    If Not xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing And _
       Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        
        If CLng(xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.LiquidacaoFutura And _
           fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text) > fgDataHoraServidor(enumFormatoDataHora.Data) Then
            Exit Function
        End If
    
    End If
    
    'Carlos - Gerenciamento de chamada das rotinas
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
        'lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    End If
    strXML = objProcessoOperacao.LerTodos(pstrFuncionalidade, _
                                          CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                                          enumIndicadorSimNao.Nao, , _
                                          lngTipoLiquidacao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    'Entrada Manual
    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        
        flEntradaManual xmlRemessa, xmlControleProcessoOperacao, pstrFuncionalidade
        objProcessoOperacao.EnviarPrevisao xmlRemessa, xmlControleProcessoOperacao

        If xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim And pstrFuncionalidade = "RecebimentoOperacao" _
        And (xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.RegistroOperacaoInterbancaria Or _
        xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.RegistroOperacaoArbitragem Or _
        xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.IFInformaLiquidacaoInterbancaria) Then
            Exit Function
        End If

    End If

    If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
        lngTipoOperacao = Val(xmlRemessa.selectSingleNode("//TP_OPER").Text)
    End If
    
    If Not xmlRemessa.selectSingleNode("//TP_SOLI") Is Nothing Then
        lngTipoSolicitacao = Val(xmlRemessa.selectSingleNode("//TP_SOLI").Text)
    End If
    
    'Confirmação Automática
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CONF").Text = enumIndicadorSimNao.Sim Then
        If Not flConfirmacao(xmlRemessa, xmlControleProcessoOperacao, pstrFuncionalidade) Then
            Exit Function
        End If
    End If

    If lngTipoOperacao = enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais _
    Or lngTipoOperacao = enumTipoOperacaoLQS.LancamentoCCCashFlow _
    Or lngTipoOperacao = enumTipoOperacaoLQS.LancamentoCCCashFlowStrikeFixo _
    Or lngTipoOperacao = enumTipoOperacaoLQS.LancamentoCCSwapJuros Then
        objProcessoOperacao.EnviarPrevisao xmlRemessa, xmlControleProcessoOperacao
    End If

    If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
        xmlControleProcessoOperacao.selectSingleNode("//IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim
        xmlControleProcessoOperacao.selectSingleNode("//IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim
        
        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
            Call objProcessoOperacao.EstornarPrevisao(xmlRemessa, xmlControleProcessoOperacao)
        End If
    End If
    
    'Conciliação Automática
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CNCL").Text = enumIndicadorSimNao.Sim Then
        If Not Conciliacao(xmlRemessa, xmlControleProcessoOperacao, pstrFuncionalidade) Then
            Exit Function
        End If
    End If

    'Realizado PJ
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
        If Not xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
            flEnviarRealizadoPJ xmlRemessa
        End If
    End If

    'Lançamento Conta Corrente
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
        flContaCorrente xmlRemessa, False
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        flEnviarRealizadoSolicitadoA6 xmlRemessa
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
        flEnviarRealizadoConfirmadoA6 xmlRemessa
    End If

    'mensagem retorno paro Legado
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        Set objLegado = CreateObject("A8LQS.clsLegado")
        strMensagem = objLegado.EnviarRetorno(xmlRemessa)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        Set objLegado = Nothing
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_LIBE").Text = enumIndicadorSimNao.Sim Then
        'Alterado tipo de ação para liberação
        'xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text = enumTipoSolicitacao.Liberar
        objOperacao.OperacaoDisponivelConsulta xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               enumIndicadorSimNao.Nao
        LiberarOperacao xmlRemessa, False, , , xmlRepeticao
    End If

    Set objOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing
    Set objProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarChamadas Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controla o ponto de envio da previsão para os sistemas PJ e A6 da operações gerada pela entrada manual.

Private Function flEntradaManual(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                 ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40, _
                                 ByVal pstrFuncionalidade As String) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlControleProcessoOperacaoAux          As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim lngTipoLiquidacao                       As Long

On Error GoTo ErrorHandler

    'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
    'na entrada manual.
    If Not xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG") Is Nothing Then
        If xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.Sim Then
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Nao
            End If
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Nao
            End If
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
            End If
        End If
    Else
        If pstrFuncionalidade = "Confirmacao" Then
            Set xmlControleProcessoOperacaoAux = CreateObject("MSXML2.DOMDocument.4.0")
            Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
            If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
                'lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
            End If
            strXML = objProcessoOperacao.LerTodos("RecebimentoBoleto", _
                                                  CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                                                  enumIndicadorSimNao.Nao)

            Set objProcessoOperacao = Nothing
            If strXML = "" Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Nao
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Nao
            Else
                xmlControleProcessoOperacaoAux.loadXML strXML

                'se entrada manual envia a previsão na confirmação
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = xmlControleProcessoOperacaoAux.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = xmlControleProcessoOperacaoAux.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text
            End If
        End If
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verifica se a funcionalidade de Confirmação está automático ou não no workflow, controlando o status que a operação deve ficar.

Private Function flConfirmacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                               ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40, _
                               ByVal pstrFuncionalidade As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objContingencia                         As A8LQS.clsContingencia
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim objProcessoCCR                          As A8LQS.clsProcessoMensagemCCR
Dim dtmDataOperacao                         As Date
Dim lngStatusOperacao                       As Long
Dim lngCodigoRetornoVerificacao             As Long

On Error GoTo ErrorHandler

    flConfirmacao = False

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objContingencia = CreateObject("A8LQS.clsContingencia")
    Set objRegra = CreateObject("A8LQS.clsRegra")
    
    
    'Verifica regra para confirmação automática
    If objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Confirmar, lngCodigoRetornoVerificacao) Then
        
        'Verifica para Despesas e Eventos, caso seja confirmação automática não alterar o Status para Concordância.
        Set objContingencia = CreateObject("A8LQS.clsContingencia")
        
        If objContingencia.VerificarSistContingencia(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) Then
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         enumJustificativa.SistemaEmContingencia
            Set objHistSituacaoOperacao = Nothing
            Set objOperacao = Nothing
            Set objContingencia = Nothing
            Set objRegra = Nothing
            
            Exit Function
        
        Else
            'Qdo de entrada manual sempre parar na entrada manual
            If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
                If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                    If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
                        
                        'If xmlRemessa.selectSingleNode("//TP_OPER").Text <> enumTipoOperacaoLQS.NegociacaoOperacaoCCR Then
                    
                            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                         enumJustificativa.EntradaManual
                            Set objHistSituacaoOperacao = Nothing
                            Set objOperacao = Nothing
                            Set objContingencia = Nothing
                            Set objRegra = Nothing
                            Exit Function
                        'End If
                    End If
                End If
                
            End If

            'Para Operações Definitivas e Compromissadas a data pode ser retroativa, com isso para para a confirmação
            If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                dtmDataOperacao = fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text)
            Else
                dtmDataOperacao = fgDataHoraServidor(enumFormatoDataHora.Data)
            End If

            If dtmDataOperacao = fgDataHoraServidor(enumFormatoDataHora.Data) Then
                
                If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
                    
                    If xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoEventoCBLC Then
                        
                        lngStatusOperacao = enumStatusOperacao.ConcordanciaBackofficeAutomatico
                    
                    'BMF
                    ElseIf xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF Then
                        
                        lngStatusOperacao = enumStatusOperacao.ConcordanciaBackofficeAutomatico
                    
                    ElseIf xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Then
                        
                        lngStatusOperacao = enumStatusOperacao.ConcordanciaBalcaoAutomatica
                    
                    ElseIf xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.NegociacaoOperacaoCCR Then
                    
                        'CCR - KIDA
                        Set objProcessoCCR = CreateObject("A8LQS.clsProcessoMensagemCCR")
                        
                        If Not xmlRemessa.selectSingleNode("//TpManut") Is Nothing Then
                            If xmlRemessa.documentElement.selectSingleNode("//TpManut").Text = "I" Then
                                If Not objProcessoCCR.ConciliarOperacao(CLng(xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text)) Then
                                    
                                    Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                                    
                                    objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                                 xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                                 enumJustificativa.SemMensagemCCR0006
                                    Set objHistSituacaoOperacao = Nothing
                                    Set objOperacao = Nothing
                                    Set objContingencia = Nothing
                                    Set objRegra = Nothing
                                    Exit Function
                                End If
                            End If
                        End If
                        Set objProcessoCCR = Nothing
                        lngStatusOperacao = enumStatusOperacao.ConcordanciaAutomatica
                    Else
                        'RATS 946
                        If Val(xmlRemessa.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
                            lngStatusOperacao = enumStatusOperacao.ConcordanciaBackofficeAutomatico
                        Else
                            lngStatusOperacao = enumStatusOperacao.ConcordanciaAutomatica
                        End If
                    End If
                Else
                    lngStatusOperacao = enumStatusOperacao.ConcordanciaAutomatica
                End If
                
                objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                    lngStatusOperacao
                xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
            
            Else
                Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                             xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                             enumJustificativa.OperacaoRetroativa
                Set objHistSituacaoOperacao = Nothing
                Exit Function
            End If
        End If
        Set objContingencia = Nothing
    Else
    
        If xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.NegociacaoOperacaoCCR Then
            'CCR - KIDA
            If pstrFuncionalidade = "Confirmacao" Then
                Set objProcessoCCR = CreateObject("A8LQS.clsProcessoMensagemCCR")
                Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                
                If Not xmlRemessa.selectSingleNode("//TpManut") Is Nothing Then
                    If xmlRemessa.documentElement.selectSingleNode("//TpManut").Text = "I" Then
                        If Not objProcessoCCR.ConciliarOperacao(CLng(xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text)) Then
                            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                         enumJustificativa.SemMensagemCCR0006
                        End If
                    End If
                End If
                Set objProcessoCCR = Nothing
            End If
        Else
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         IIf(lngCodigoRetornoVerificacao <> 0, _
                                                                lngCodigoRetornoVerificacao, _
                                                                enumJustificativa.RegraWorkflow)
            Set objHistSituacaoOperacao = Nothing
        End If
        
        Set objHistSituacaoOperacao = Nothing
        Set objOperacao = Nothing
        Set objContingencia = Nothing
        Set objRegra = Nothing
        
        Exit Function
    End If
    Set objRegra = Nothing

    flConfirmacao = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConfirmacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Function Conciliacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                   Optional ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40, _
                   Optional ByVal pstrFuncionalidade As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objRegra                                As A8LQS.clsRegra
Dim objContingencia                         As A8LQS.clsContingencia
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objAlerta                               As A8LQS.clsAlerta
Dim lngJustificativa                        As enumJustificativa
Dim lngCodigoRetornoVerificacao             As Long
Dim lngStatusOperacao                       As Long

    On Error GoTo ErrorHandler

    Conciliacao = False

    Set objRegra = CreateObject("A8LQS.clsRegra")
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    
    lngStatusOperacao = enumStatusOperacao.AConciliar
    
    If (xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao _
    And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarRegistro) _
    Or (xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica _
    And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.EmSer) _
    Or (xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica _
    And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarRegistro) Then
        
        If xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica _
        And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.AConciliarRegistro Then
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliarRegistro
            Set objOperacao = Nothing
        End If
        
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        If Not objConciliacao.ConciliarOperacao(xmlRemessa, lngJustificativa) Then
            
            If Not xmlRemessa.selectSingleNode("//ENCONTROU_IDENT_NEG_BMC") Is Nothing Then
                Call flGeraAlertaCamara("BMC0015DOPER", _
                                        enumTipoBackOffice.Tesouraria, _
                                        xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                        xmlRemessa)
            
            Else
                Call flGeraAlertaCamara("BMC0015NLOCA", _
                                        enumTipoBackOffice.Tesouraria, _
                                        xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                        xmlRemessa)
            
            End If
            
            Set objRegra = Nothing
            Set objConciliacao = Nothing
            Set objAlerta = Nothing
            Exit Function
        End If
    
    ElseIf (xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao _
    Or xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica) _
    And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarAceite Then
        
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        If Not objConciliacao.ConciliarOperacao(xmlRemessa, lngJustificativa) Then
            Call flGeraAlertaCamara("BMC0005NLOCA", _
                                    enumTipoBackOffice.Tesouraria, _
                                    xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                    xmlRemessa)
            Set objRegra = Nothing
            Set objConciliacao = Nothing
            Set objAlerta = Nothing
            Exit Function
        End If
    
    ElseIf xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacoesBMC _
    And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.EmSer _
    Or xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacoesBMC _
    And xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarAceite Then
        
        If xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.AConciliarAceite Then
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliarAceite
            Set objOperacao = Nothing
        End If
        
        If objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Conciliar, lngCodigoRetornoVerificacao) Then
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            If Not objConciliacao.ConciliarOperacao(xmlRemessa, lngJustificativa) Then
                Call flGeraAlertaCamara("BMC0005NLOCA", _
                                        enumTipoBackOffice.Tesouraria, _
                                        xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                        xmlRemessa)
                Set objRegra = Nothing
                Set objConciliacao = Nothing
                Set objAlerta = Nothing
                Exit Function
            End If
        Else
            Exit Function
        End If
    
    Else
        'Verifica regra para conciliação automática
        If objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Conciliar, lngCodigoRetornoVerificacao) Then
            Set objContingencia = CreateObject("A8LQS.clsContingencia")
            If objContingencia.VerificarSistContingencia(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) Then
                Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                             lngStatusOperacao, _
                                                             enumJustificativa.SistemaEmContingencia
                Set objHistSituacaoOperacao = Nothing
                Exit Function
            Else
                Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                If Not objConciliacao.ConciliarOperacao(xmlRemessa, lngJustificativa) Then
                    Set objOperacao = CreateObject("A8LQS.clsOperacao")
                    
                    
                    If Not xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
                        If Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.Registrada And _
                           Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.RegistradaAutomatica Then
                            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                lngStatusOperacao, _
                                                                lngJustificativa
                        End If
                    End If
                    
                    'Geração de alerta
                    objAlerta.GerarAlerta xmlRemessa, , , _
                                          enumFatorGeradorAlerta.OperacaoNaoConciliadaAutomatica
                    Set objOperacao = Nothing
                    Set objConciliacao = Nothing
                    Exit Function
                End If
                Set objConciliacao = Nothing
            End If
        Else
            If Not xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
                If Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.Registrada And _
                   Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.RegistradaAutomatica Then
                     
                    Set objOperacao = CreateObject("A8LQS.clsOperacao")
                    objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                        lngStatusOperacao, _
                                                        IIf(lngCodigoRetornoVerificacao <> 0, _
                                                            lngCodigoRetornoVerificacao, _
                                                            enumJustificativa.RegraWorkflow)
                    Set objOperacao = Nothing
                    
                End If
            End If
            
            Exit Function
        End If
    End If
    
    Set objRegra = Nothing
    Set objConciliacao = Nothing
    Set objContingencia = Nothing
    Set objAlerta = Nothing

    Conciliacao = True

    Exit Function
ErrorHandler:

    Set objRegra = Nothing
    Set objConciliacao = Nothing
    Set objContingencia = Nothing
    Set objAlerta = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Conciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executa a chamada para o envio do realizado para o PJ verificando regras especificas por operação.

Private Function flEnviarRealizadoPJ(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEnviarRealizadoPJ = False

    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
'            If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
'                flEnviarRealizadoPJ = False
                Exit Function
'            End If
        End If
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing

    flEnviarRealizadoPJ = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function flEnviarPrivistoMEPJ(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEnviarPrivistoMEPJ = False

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EnviarPrevisaoME(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing

    flEnviarPrivistoMEPJ = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
'Executa a chamada para o estorno do realizado para o PJ verificando regras especificas por operação.

Private Function flEstornarRealizadoPJ(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEstornarRealizadoPJ = False

    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
'            If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
'                flEstornarRealizadoPJ = False
                Exit Function
'            End If
        End If
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EstornarRealizado(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing

    flEstornarRealizadoPJ = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEstornarRealizadoPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executa a chamada para o envio do realizado solicitado para o A6 verificando regras especificas por operação.

Private Function flEnviarRealizadoSolicitadoA6(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEnviarRealizadoSolicitadoA6 = False

    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
    '            flEnviarRealizadoSolicitadoA6 = False
            Exit Function
    '        End If
        End If
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
    Set objProcessoOperacao = Nothing
    Set objSubReserva = Nothing

    flEnviarRealizadoSolicitadoA6 = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoSolicitadoA6 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executa a chamada para o envio do estorno do realizado solicitado para o A6 verificando regras especificas por operação.

Private Function flEstornarRealizadoSolicitadoA6(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEstornarRealizadoSolicitadoA6 = False

    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
    '            flEstornarRealizadoSolicitadoA6 = False
            Exit Function
    '        End If
        End If
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
    Set objProcessoOperacao = Nothing
    Set objSubReserva = Nothing

    flEstornarRealizadoSolicitadoA6 = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEstornarRealizadoSolicitadoA6 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executa a chamada para o envio do realizado confirmado para o A6 verificando regras especificas por operação.

Private Function flEnviarRealizadoConfirmadoA6(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEnviarRealizadoConfirmadoA6 = False

    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
    '            flEnviarRealizadoConfirmadoA6 = False
            Exit Function
    '        End If
        End If
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
    Set objProcessoOperacao = Nothing
    Set objSubReserva = Nothing

    flEnviarRealizadoConfirmadoA6 = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoConfirmadoA6 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controla a liberação mensagem SPB gerada por uma operação.

Public Function LiberarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                ByVal pblnLiberarViaTela As Boolean, _
                       Optional ByVal pblnIgnoraGradeHorario As Boolean = False, _
                       Optional ByVal pblnOrigemConciliacao As Boolean = False, _
                       Optional ByRef xmlRepeticao As MSXML2.DOMDocument40, _
                       Optional ByVal pblnFluxoAutomatico As Boolean = False) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objMensagem                             As A8LQS.clsMensagem
Dim objLegado                               As A8LQS.clsLegado
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlCAM0002                              As MSXML2.DOMDocument40
Dim xmlCAM0005R2                            As MSXML2.DOMDocument40
Dim xmlCAM0006R2                            As MSXML2.DOMDocument40
Dim xmlCAM0009R2                            As MSXML2.DOMDocument40
Dim xmlCAM0013R2                            As MSXML2.DOMDocument40
Dim xmlBMC0005                              As MSXML2.DOMDocument40
Dim xmlBMC0015                              As MSXML2.DOMDocument40
Dim xmlCAM0055                              As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim strMensagem                             As String
Dim strDataRegistroMensagemSPB              As String
Dim blnEnviarMensagem                       As Boolean
Dim blnEnviarRetorno                        As Boolean
Dim blnExisteMensagem                       As Boolean
Dim lngTipoLiquidacao                       As Long
Dim lngCont                                 As Long
Dim lngSequencialConciliacaoOperacao        As Long 'Para guardar o sequencia gerado na tabela de conciliacao

    On Error GoTo ErrorHandler

    'Inclusão da repetição na mensagens
    If Not xmlRepeticao Is Nothing Then
        If xmlRepeticao.xml <> vbNullString Then
            fgAppendXML xmlRemessa, "MESG", xmlRepeticao.xml
        End If
    End If

    'Esta rotina atende a liberação automática e a liberação via Tela
    If Not pblnLiberarViaTela Then
        If Not flLiberarAutomatico(xmlRemessa) Then
            Exit Function
        End If
    Else
        If Not flLiberarViaTela(xmlRemessa, pblnIgnoraGradeHorario, pblnOrigemConciliacao, pblnFluxoAutomatico) Then
            Exit Function
        End If
    End If

    If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosJurosSWAP Or _
       xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosJurosTERMO Or _
       xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosCETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.SemModalidade Then
            flLiquidarEventosJurosCETIP xmlRemessa
            Exit Function
        End If
    End If

    'Carlos - Gerenciamento de chamada das rotinas
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
        'lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    End If
    
    strXML = objProcessoOperacao.LerTodos("Liberacao", _
                                          CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                                          enumIndicadorSimNao.Nao)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
    'na entrada manual.
    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        If Not xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.Sim Then
                If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim
                Else
                    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
                End If
            End If
        End If
    End If
            
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
            objProcessoOperacao.EnviarPrevisaoBMACetip xmlRemessa, xmlControleProcessoOperacao
        'Pagamento de Despesas
        'ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas And (xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.EnvioSTR0006PagDespesas Or xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.EnvioSTR0008PagDespesas) Then
        '    objProcessoOperacao.EnviarPrevisaoBMACetip xmlRemessa, xmlControleProcessoOperacao
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        flEnviarRealizadoSolicitadoA6 xmlRemessa
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
        flEnviarRealizadoConfirmadoA6 xmlRemessa
    End If
    
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        
        If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                flEnviarRealizadoPJ xmlRemessa
            End If
        
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                flEnviarRealizadoPJ xmlRemessa
            End If
            
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.LancamentoContaCorrenteBG Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                flEnviarRealizadoPJ xmlRemessa
            End If
        
        'CCR
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.NegociacaoOperacaoCCR Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
                flEnviarPrivistoMEPJ xmlRemessa
            End If
            
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.IFInformaLiquidacaoInterbancaria Then
            
            'PJ Moeda Estrangeira
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ_ME").Text = enumIndicadorSimNao.Sim Then
                flEnviarRealizadoMEPJ xmlRemessa
            End If
            
            'PJ Moeda Nacional só é enviado quando TipoNegociaçãoCambial = 1 (Interbancária)
            If Not xmlRemessa.selectSingleNode("//TP_NEGO_CAML") Is Nothing Then
                If xmlRemessa.selectSingleNode("//TP_NEGO_CAML").Text = enumTipoNegociacaoCambial.Interbancaria Then
                    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                        flEnviarRealizadoPJ xmlRemessa
                    End If
                End If
            End If
            
        End If
         
    End If
    
    'Lançamento Conta Corrente
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
        flContaCorrente xmlRemessa, False
    End If

    'Envio Mensagem SPB
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_SPB").Text = enumIndicadorSimNao.Sim Then
        'Se Operacao = Registro Operacao BMC Balcao
        If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Then
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.CancelamentoSolicitado _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaReativacao _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaReativacaoAutomatica Then
                Set objMensagem = CreateObject("A8LQS.clsMensagem")
                strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
                If blnEnviarMensagem Then
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                End If
                Set objMensagem = Nothing
            End If
        'Operacao 230 com status Liberada e Liberada Automatica
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega Then
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Then
                Set objMensagem = CreateObject("A8LQS.clsMensagem")
                strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
                If blnEnviarMensagem Then
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                End If
                Set objMensagem = Nothing
            End If
        'Operacao 231 com status Liberada e Liberada Automatica
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega Then
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica Then
                'Busca Mensagem CAM0006R2
                Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Set xmlCAM0006R2 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0006R2", , , , xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, _
                                                                     , , , , , , xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text)
                'Se Entrada Manual altera status da CAM0006R2 para Confirmada Contingencia e Associa a Operacao
                If xmlCAM0006R2.xml <> vbNullString Then
                    If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                        If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                            objMesgRecbEnviSPB.AlterarStatus xmlCAM0006R2.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                             xmlCAM0006R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                             enumStatusMensagem.ConfirmadaContingencia, _
                                                             xmlCAM0006R2.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, , , _
                                                             xmlCAM0006R2.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                                                             xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                        End If
                    End If
                End If
                
                Set xmlCAM0006R2 = Nothing
                
                'Busca Mensagem CAM0055
                strDataRegistroMensagemSPB = fgDtHr_To_Xml(DateTime.Now)
                blnExisteMensagem = True
                Do While blnExisteMensagem
                    Set xmlCAM0055 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0055", , strDataRegistroMensagemSPB, , _
                                                                       xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, , , , , , , _
                                                                       xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text)
                    'Associa Mensagem CAM0055 a Operacao
                    If xmlCAM0055.xml <> vbNullString Then
                        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                                objMesgRecbEnviSPB.Alterar xmlCAM0055.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                           xmlCAM0055.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           xmlCAM0055.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , _
                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                            End If
                        End If
                        strDataRegistroMensagemSPB = xmlCAM0055.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
                    Else
                        blnExisteMensagem = False
                    End If
                    Set xmlCAM0055 = Nothing
                Loop
                
                'Busca Mensagem CAM0005R2
                strDataRegistroMensagemSPB = fgDtHr_To_Xml(DateTime.Now)
                blnExisteMensagem = True
                Do While blnExisteMensagem
                    Set xmlCAM0005R2 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0005R2", , strDataRegistroMensagemSPB, , _
                                                                         xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, , , , , , , _
                                                                         xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text)
                    'Associa Mensagem CAM0005R2 a Operacao
                    If xmlCAM0005R2.xml <> vbNullString Then
                        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                                objMesgRecbEnviSPB.Alterar xmlCAM0005R2.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                           xmlCAM0005R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           xmlCAM0005R2.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , _
                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                            End If
                        End If
                        strDataRegistroMensagemSPB = xmlCAM0005R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
                    Else
                        blnExisteMensagem = False
                    End If
                    Set xmlCAM0005R2 = Nothing
                Loop
                
                Set objMesgRecbEnviSPB = Nothing
                
                'Monta a Mensagem SPB
                Set objMensagem = CreateObject("A8LQS.clsMensagem")
                strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
                If blnEnviarMensagem Then
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                End If
                Set objMensagem = Nothing
            End If
        
        'Operacao 233 com status Liberada e Liberada Automatica
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica Then
                'Busca Mensagem CAM0009R2
                Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Set xmlCAM0009R2 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0009R2", , , , xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, _
                                                                     , , , , , , xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text)
                'Se Entrada Manual altera status da CAM0009R2 para Confirmada Contingencia
                If xmlCAM0009R2.xml <> vbNullString Then
                    If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                        If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                            objMesgRecbEnviSPB.AlterarStatus xmlCAM0009R2.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                             xmlCAM0009R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                             enumStatusMensagem.ConfirmadaContingencia, _
                                                             xmlCAM0009R2.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, , , _
                                                             xmlCAM0009R2.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                                                             xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                        End If
                    End If
                End If
                
                Set xmlCAM0009R2 = Nothing
                
                'Busca Mensagem CAM0055
                strDataRegistroMensagemSPB = fgDtHr_To_Xml(DateTime.Now)
                blnExisteMensagem = True
                Do While blnExisteMensagem
                    Set xmlCAM0055 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0055", , strDataRegistroMensagemSPB, , _
                                                                       xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, , , , , , , _
                                                                       xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text)
                    'Associa Mensagem CAM0055 a Operacao
                    If xmlCAM0055.xml <> vbNullString Then
                        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                                objMesgRecbEnviSPB.Alterar xmlCAM0055.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                           xmlCAM0055.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           xmlCAM0055.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , _
                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                            End If
                        End If
                        strDataRegistroMensagemSPB = xmlCAM0055.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
                    Else
                        blnExisteMensagem = False
                    End If
                    Set xmlCAM0055 = Nothing
                Loop
                
                'Busca Mensagem CAM0005R2
                strDataRegistroMensagemSPB = fgDtHr_To_Xml(DateTime.Now)
                blnExisteMensagem = True
                Do While blnExisteMensagem
                    Set xmlCAM0005R2 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0005R2", , strDataRegistroMensagemSPB, , _
                                                                         xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, , , , , , , _
                                                                         xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text)
                    'Associa Mensagem CAM0005R2 a Operacao
                    If xmlCAM0005R2.xml <> vbNullString Then
                        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                                objMesgRecbEnviSPB.Alterar xmlCAM0005R2.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                           xmlCAM0005R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           xmlCAM0005R2.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , _
                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                            End If
                        End If
                        strDataRegistroMensagemSPB = xmlCAM0005R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
                    Else
                        blnExisteMensagem = False
                    End If
                    Set xmlCAM0005R2 = Nothing
                Loop
                
                Set objMesgRecbEnviSPB = Nothing
                
                'Monta a Mensagem SPB
                Set objMensagem = CreateObject("A8LQS.clsMensagem")
                strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
                If blnEnviarMensagem Then
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                End If
                Set objMensagem = Nothing
            End If
        
        'Operacao 234 com status Liberada e Liberada Automatica
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica Then
                'Busca Mensagem BMC0015
                Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Set xmlBMC0015 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "BMC0015", , , , xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, _
                                                                   , , , xmlRemessa.selectSingleNode("//ChACAM").Text)
                'Se Entrada Manual altera status da BMC0015 para Conciliada Contingencia
                If xmlBMC0015.xml <> vbNullString Then
                
                    If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                        If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                            
                            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                            
                            lngSequencialConciliacaoOperacao = objConciliacao.IncluirJustificativaConciliacaoOperacao()
                            
                            'Concilia BMC0015 com a Operacao
                            Call objConciliacao.IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                                           xmlBMC0015.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                                           xmlBMC0015.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                                           0, _
                                                                           xmlBMC0015.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                            
                            'Altera status da BMC0015 para Conciliada Contingencia
                            objMesgRecbEnviSPB.AlterarStatus xmlBMC0015.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                             xmlBMC0015.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                             enumStatusMensagem.ConciliadaContingencia, _
                                                             xmlBMC0015.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, , , _
                                                             xmlBMC0015.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text
                        End If
                    End If
                    
                    'Append Mensagem BMC0015 no Layout 234 para enviar a Mensagem SPB
                    fgAppendXML xmlRemessa, "MESG", xmlBMC0015.xml
                Else
                    fgAppendNode xmlRemessa, "MESG", "DtMovto", xmlRemessa.documentElement.selectSingleNode("DT_MOVI").Text
                End If
                Set objMesgRecbEnviSPB = Nothing
                Set xmlBMC0015 = Nothing
                
                'Monta a Mensagem SPB
                Set objMensagem = CreateObject("A8LQS.clsMensagem")
                strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
                If blnEnviarMensagem Then
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                End If
                Set objMensagem = Nothing
            End If
        
        'Operacao 237 com status Liberada e Liberada Automatica
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica Then
                'Busca Mensagem CAM0013R2
                Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Set xmlCAM0013R2 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0013R2", , , , xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, _
                                                                    , , , , , , xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text, _
                                                                     xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB2").Text)
                'Se Entrada Manual altera status da CAM0013R2 para Confirmada Contingencia
                If xmlCAM0013R2.xml <> vbNullString Then
                    If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                        If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                            objMesgRecbEnviSPB.AlterarStatus xmlCAM0013R2.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                             xmlCAM0013R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                             enumStatusMensagem.ConfirmadaContingencia, _
                                                             xmlCAM0013R2.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, , , _
                                                             xmlCAM0013R2.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                                                             xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                        End If
                    End If
                End If
                
                Set xmlCAM0013R2 = Nothing
                
                'Busca Mensagem CAM0055
                strDataRegistroMensagemSPB = fgDtHr_To_Xml(DateTime.Now)
                blnExisteMensagem = True
                Do While blnExisteMensagem
                    Set xmlCAM0055 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0055", , strDataRegistroMensagemSPB, , _
                                                                       xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, , , , , , , _
                                                                       xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text, xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB2").Text)
                    'Associa Mensagem CAM0055 a Operacao
                    If xmlCAM0055.xml <> vbNullString Then
                        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                                objMesgRecbEnviSPB.Alterar xmlCAM0055.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                           xmlCAM0055.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           xmlCAM0055.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , _
                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                            End If
                        End If
                        strDataRegistroMensagemSPB = xmlCAM0055.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
                    Else
                        blnExisteMensagem = False
                    End If
                    Set xmlCAM0055 = Nothing
                Loop
                
                'Busca Mensagem CAM0005R2
                strDataRegistroMensagemSPB = fgDtHr_To_Xml(DateTime.Now)
                blnExisteMensagem = True
                Do While blnExisteMensagem
                    Set xmlCAM0005R2 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0005R2", , strDataRegistroMensagemSPB, , _
                                                                         xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, , , , , , , _
                                                                         xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB").Text, xmlRemessa.selectSingleNode("//CO_REG_OPER_CAMB2").Text)
                    'Associa Mensagem CAM0005R2 a Operacao
                    If xmlCAM0005R2.xml <> vbNullString Then
                        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If xmlRemessa.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                                objMesgRecbEnviSPB.Alterar xmlCAM0005R2.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                           xmlCAM0005R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           xmlCAM0005R2.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , _
                                                           xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
                            End If
                        End If
                        strDataRegistroMensagemSPB = xmlCAM0005R2.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
                    Else
                        blnExisteMensagem = False
                    End If
                    Set xmlCAM0005R2 = Nothing
                Loop
                
                Set objMesgRecbEnviSPB = Nothing
                
                'Monta a Mensagem SPB
                Set objMensagem = CreateObject("A8LQS.clsMensagem")
                strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
                If blnEnviarMensagem Then
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                End If
                Set objMensagem = Nothing
            End If
        
        Else
            Set objMensagem = CreateObject("A8LQS.clsMensagem")
            strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
            If blnEnviarMensagem Then
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
            End If
            Set objMensagem = Nothing
        End If
    End If
    
    'Envio Mensagem Retorno Legado
    blnEnviarRetorno = True
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        If Not xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                blnEnviarRetorno = False
            End If
        End If
        If blnEnviarRetorno Then
            'Liberacao Mensagens CAM
            If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao _
            Or xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica _
            Or xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacoesBMC Then
                If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada _
                Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Then
                    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                    Set xmlCAM0002 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0002R1", _
                                                                       , , xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                    If xmlCAM0002.xml <> vbNullString Then
                        fgAppendNode xmlRemessa, "MESG", "NR_CNTR_CAMB_CPRA", xmlCAM0002.selectSingleNode("//NumContrtoCAMCompra").Text
                        fgAppendNode xmlRemessa, "MESG", "NR_CNTR_CAMB_VEND", xmlCAM0002.selectSingleNode("//NumContrtoCAMVenda").Text
                        fgAppendNode xmlRemessa, "MESG", "NR_REFE_UNIC_BACEN", xmlCAM0002.selectSingleNode("//CodRefUniBC").Text
                    Else
                        fgAppendNode xmlRemessa, "MESG", "NR_CNTR_CAMB_CPRA", ""
                        fgAppendNode xmlRemessa, "MESG", "NR_CNTR_CAMB_VEND", ""
                        fgAppendNode xmlRemessa, "MESG", "NR_REFE_UNIC_BACEN", ""
                    End If
                    ' Envia Retorno Legado
                    Set objLegado = CreateObject("A8LQS.clsLegado")
                    strMensagem = objLegado.EnviarRetorno(xmlRemessa)
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
                    Set objMesgRecbEnviSPB = Nothing
                    Set xmlCAM0002 = Nothing
                    Set objLegado = Nothing
                End If
            'Operacoes 230, 231 e 234 com status Registrada ou Registrada Automatica
            ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
            Or xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
            Or xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada _
                Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Then
                    
                    If Not xmlRemessa.documentElement.selectSingleNode("CO_MESG_SPB") Is Nothing Then
                        xmlRemessa.documentElement.selectSingleNode("CO_MESG_SPB").Text = "BMC0005"
                    Else
                        fgAppendNode xmlRemessa, "MESG", "CO_MESG_SPB", "BMC0005"
                    End If
                    
                    ' Appenda BMC0005 no XML de retorno para o legado, pois os campos dela são esperados na TagEntrada da Regra de Transporte 250
                    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                    Set xmlBMC0005 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "BMC0005", , , , , , , , xmlRemessa.selectSingleNode("//ChACAM").Text)
                    If xmlBMC0005.xml <> vbNullString Then
                        fgAppendXML xmlRemessa, "MESG", xmlBMC0005.xml
                    Else
                        fgAppendNode xmlRemessa, "MESG", "DtMovto", xmlRemessa.documentElement.selectSingleNode("DT_MOVI").Text
                    End If
                    
                    ' Envia Retorno Legado
                    Set objLegado = CreateObject("A8LQS.clsLegado")
                    strMensagem = objLegado.EnviarRetorno(xmlRemessa)
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
                    
                    ' Finaliza objetos
                    Set objLegado = Nothing
                    Set objMesgRecbEnviSPB = Nothing
                    Set xmlBMC0005 = Nothing
                End If
            Else
                Set objLegado = CreateObject("A8LQS.clsLegado")
                strMensagem = objLegado.EnviarRetorno(xmlRemessa)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
                Set objLegado = Nothing
            End If
        End If
    End If

    If Not xmlRemessa.selectSingleNode("//CO_LOCA_LIQU") Is Nothing And _
       Not xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
        
        If Val(xmlRemessa.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP And _
          (Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.RegistradaAutomatica Or _
           Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Registrada) Then
           
            Call Conciliacao(xmlRemessa, xmlControleProcessoOperacao, "Liberacao")
            
        End If
    
    End If
    
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objLegado = Nothing
    Set xmlCAM0002 = Nothing
    Set xmlBMC0005 = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controla a liberação mensagem SPB gerada por uma outra mensagem SPB.

Public Function LiberarMensagem(ByVal pstrNumeroControleIF As String, _
                                ByVal penumFuncaoSistema As enumFuncaoSistema, _
                       Optional ByRef xmlRepeticao As MSXML2.DOMDocument40, _
                       Optional ByRef plngRepeticao As Long, _
                       Optional ByVal pblnIgnoraGradeHorario As Boolean = False, _
                       Optional ByVal pstrDataRegistroMensagemSPB As String = vbNullString, _
                       Optional ByVal plngNuOperAtiv As Long) As Boolean

Dim objMensagem                             As A8LQS.clsMensagem
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlCodigoMensagem                       As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strCodigoMensagem                       As String
Dim blnEnviarMensagem                       As Boolean

Dim strMensagemEnvio                        As String
Dim lngLocalLiquidacao                      As Long
Dim dblValorMsgEnvio                        As Double
Dim dblValorMsgAnterior                     As Double
Dim xmlMensagemAnterior                     As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    'Ler a mensagem spb
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMensagem = objMensagemEnviadaRecebida.LerXMLMensagem(pstrNumeroControleIF, _
                                                                plngRepeticao, , , , _
                                                                pstrDataRegistroMensagemSPB)

    If xmlMensagem.xml = vbNullString Then
        ' 3050 - Número de Controle IF inválido.
        lngCodigoErroNegocio = 3050
        GoTo ErrorHandler
    End If

    strCodigoMensagem = xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    strMensagem = objFluxoProcessamento.LerTodos(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, , _
                                                 penumFuncaoSistema)
    Set objFluxoProcessamento = Nothing

    If strMensagem = vbNullString Then
        ' 3051 - Código de mensagem de envio não cadastrado.
        lngCodigoErroNegocio = 3051
        GoTo ErrorHandler
    End If

    Set xmlCodigoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    xmlCodigoMensagem.loadXML strMensagem

    xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = xmlCodigoMensagem.documentElement.selectSingleNode("//CO_MESG_SPB_ENVI").Text

    If penumFuncaoSistema = enumFuncaoSistema.PagarContingencia Or penumFuncaoSistema = enumFuncaoSistema.PagarSTR Then

        Call flDigitacaoAlcada(xmlMensagem)

    Else
        If Not pblnIgnoraGradeHorario Then
            Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
            If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 0, enumLocalLiquidacao.BMA) Then
                ' 3095 - Horário limite para envio da Mensagem excedido.
                lngCodigoErroNegocio = 3095
                GoTo ErrorHandler
            End If
        End If
    End If
        
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    'Inclusão da repetição na mensagens
    If Not xmlRepeticao Is Nothing Then
                
        'KIDA 15/03/2010
        'BOOK
        If Not xmlMensagem.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            If xmlMensagem.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
                If strCodigoMensagem = "LTR0001" Then
                    If xmlRepeticao.xml <> vbNullString Then
                        fgAppendXML xmlMensagem, "SISMSG", xmlRepeticao.xml
                    End If
                End If
            End If
        End If
        
        If Not xmlRepeticao.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            If xmlRepeticao.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Then
                If strCodigoMensagem = "LDL0001" Then
                    objMensagem.MontarRepeticaoLDL0003 xmlMensagem, xmlRepeticao
                    
                    If xmlRepeticao.xml <> vbNullString Then
                        fgAppendXML xmlMensagem, "SISMSG", xmlRepeticao.xml
                    End If
                    
                End If
            End If
        End If
        
        'KIDA - BMF
        If Not xmlMensagem.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            If xmlMensagem.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMD Then
                
                If strCodigoMensagem = "LDL0001" Then
                    If Not xmlMensagem.selectSingleNode("//TpInf") Is Nothing Then
                        If xmlMensagem.selectSingleNode("//TpInf").Text = "P" Then
                            objMensagem.MontarRepeticaoLDL0004Previa_BMF xmlMensagem, xmlRepeticao
                        Else
                            If penumFuncaoSistema = enumFuncaoSistema.Concordar Or _
                               penumFuncaoSistema = enumFuncaoSistema.Discordar Then
                                objMensagem.MontarRepeticaoLDL0003_BMF xmlMensagem, xmlRepeticao
                            Else
                                objMensagem.MontarRepeticaoLDL0004Previa_BMF xmlMensagem, xmlRepeticao
                            End If
                        End If
                        
                        If xmlRepeticao.xml <> vbNullString Then
                            fgAppendXML xmlMensagem, "SISMSG", xmlRepeticao.xml
                        End If

                    Else
                        
                        objMensagem.MontarRepeticaoLDL0003_BMF xmlMensagem, xmlRepeticao
                        
                        If xmlRepeticao.xml <> vbNullString Then
                            fgAppendXML xmlMensagem, "SISMSG", xmlRepeticao.xml
                        End If
                        
                    End If
                End If
                
            End If
        End If
        
    End If

    If Not xmlMensagem.selectSingleNode("//CO_MESG") Is Nothing Then
        strMensagemEnvio = xmlMensagem.selectSingleNode("//CO_MESG").Text
    End If
        
    If Not xmlMensagem.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        lngLocalLiquidacao = Val(xmlMensagem.selectSingleNode("//CO_LOCA_LIQU").Text)
    End If
    
    If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
        Select Case strMensagemEnvio
            Case "LDL0003", "LDL0004", "BMC0102"
        
                Set xmlMensagemAnterior = objMensagemEnviadaRecebida.LerXMLMensagem(, , , _
                                                                                    strMensagemEnvio, , , , _
                                                                                    Left$(xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, 8), , _
                                                                                    lngLocalLiquidacao)
    
                If xmlMensagemAnterior.xml <> vbNullString Then
                    
                    dblValorMsgEnvio = fgVlrXml_To_Decimal(fgObterValorMensagemPelaTag(xmlMensagem))
                    dblValorMsgAnterior = fgVlrXml_To_Decimal(fgObterValorMensagemPelaTag(xmlMensagemAnterior))
                    
                    dblValorMsgEnvio = dblValorMsgEnvio - dblValorMsgAnterior
                    
                    strMensagemEnvio = Replace$(xmlMensagem.xml, _
                                                fgObterValorMensagemPelaTag(xmlMensagem), _
                                                fgVlr_To_Xml(dblValorMsgEnvio))
                                                 
                    Call xmlMensagem.loadXML(strMensagemEnvio)
                
                End If
                
        End Select
    End If
    
    strMensagem = objMensagem.MontarMensagemPorMensagemSPB(xmlMensagem, blnEnviarMensagem, , plngNuOperAtiv)
    If blnEnviarMensagem Then
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, True
        Set objProcessoOperacao = Nothing
    End If

    Set objMensagem = Nothing
    Set objGradeHorario = Nothing
    Set xmlCodigoMensagem = Nothing
    Set objMensagemEnviadaRecebida = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objMensagem = Nothing
    Set objGradeHorario = Nothing
    Set xmlCodigoMensagem = Nothing
    Set objMensagemEnviadaRecebida = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controla a gravação mensagem STR.

Public Function GravarMensagemSTR(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                         Optional ByVal pblnLiquidacaoPorLote As Boolean = False, _
                         Optional ByVal plngTipoBackOfficeLote As Long) As Boolean

Dim objMensagem                             As A8LQS.clsMensagem
Dim lngTipoBackofice                        As Long

    On Error GoTo ErrorHandler

    Call flDigitacaoAlcada(xmlMensagem, pblnLiquidacaoPorLote)
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    Call objMensagem.MontarMensagemPorMensagemSPB(xmlMensagem, False, , , pblnLiquidacaoPorLote, plngTipoBackOfficeLote)
    
    Set objMensagem = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objMensagem = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GravarMensagemSTR Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar a digitação de alçada

Private Function flDigitacaoAlcada(ByRef pxmlremessa As MSXML2.DOMDocument40, Optional ByVal pblnLiquidacaoPorLote As Boolean = False) As Boolean

Dim objPerfil                               As A6A7A8.clsPerfil
Dim vntProtocoloAlcada                      As Variant
Dim strValorMensagem                        As String

    On Error GoTo ErrorHandler

    If pblnLiquidacaoPorLote Then
        vntProtocoloAlcada = Format$(Now, "DDHHMMSS")
    Else
        strValorMensagem = fgObterValorMensagemPelaTag(pxmlremessa)
        
        Set objPerfil = CreateObject("A6A7A8.clsPerfil")
        Call objPerfil.DigitarAlcada(fgUsuarioRede, _
                                     strValorMensagem, _
                                     vntProtocoloAlcada)
    End If
    
    If vntProtocoloAlcada <> vbNullString Then
        flDigitacaoAlcada = True
        
        Call fgAppendNode(pxmlremessa, "SISMSG", "NU_PRTC_MESG_LG", vntProtocoloAlcada)
    
    Else
        flDigitacaoAlcada = False
    
    End If
    
    Set objPerfil = Nothing

    Exit Function

ErrorHandler:
    Set objPerfil = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flDigitacaoAlcada", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se a funcionalidade de Liberação está automático ou não no workflow, controlando o status que a operação deve ficar.

Private Function flLiberarAutomatico(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objAlerta                               As A8LQS.clsAlerta
Dim objMensagem                             As A8LQS.clsMensagem
Dim lngSequencialISPB                       As Long
Dim lngStatusOperacao                       As Long
Dim lngTipoAcao                             As Long
Dim lngCodigoRetornoVerificacao             As Long

    On Error GoTo ErrorHandler

    flLiberarAutomatico = False
    lngStatusOperacao = 0
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    'Libera o Registro para Consulta. Está neste ponto pois a transação pode se encerrar em vários pontos.
    objOperacao.OperacaoDisponivelConsulta xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                           enumIndicadorSimNao.Sim

    If Trim(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text) <> vbNullString Then
        'Verificar a Grade de Horário
        Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
        Set objMensagem = CreateObject("A8LQS.clsMensagem")
        If Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "LDL" Then
            objMensagem.flObterISPBLDL 1, xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, lngSequencialISPB
        Else
            lngSequencialISPB = objMensagem.ObterSequencialISPBDestino(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text)
        End If
        If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                                                  lngSequencialISPB, _
                                                                  CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text)) Then
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         enumJustificativa.GradeHorario
            Set objHistSituacaoOperacao = Nothing
            Set objAlerta = CreateObject("A8LQS.clsAlerta")
            objAlerta.GerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.MensagensBloqueadasGradeHorario
            Set objAlerta = Nothing
            Exit Function
        End If
        Set objMensagem = Nothing
        Set objGradeHorario = Nothing
    End If

    Set objRegra = CreateObject("A8LQS.clsRegra")
    'Verifica regra para liberação automática
    If Not objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Liberar, lngCodigoRetornoVerificacao) Then
        Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
        objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                     xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                     IIf(lngCodigoRetornoVerificacao <> 0, _
                                                            lngCodigoRetornoVerificacao, _
                                                            enumJustificativa.RegraWorkflow)
        Set objHistSituacaoOperacao = Nothing
        Exit Function
    End If

    lngTipoAcao = 0
    Select Case xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text
        
        Case enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
             enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA
            
            lngStatusOperacao = enumStatusOperacao.LiquidadaFisicamenteAutomatica
        
        Case enumTipoOperacaoLQS.DepositoBMA
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1002
        
        Case enumTipoOperacaoLQS.RetiradaBMA
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1006
        
        Case enumTipoOperacaoLQS.TransferenciasBMCDeposito
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1002
        
        Case enumTipoOperacaoLQS.TransferenciasBMCRetirada
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1006
        
        Case enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA

            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1016
            
        Case enumTipoOperacaoLQS.TransferenciaBMA
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1016

        Case enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
             enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
             enumTipoOperacaoLQS.EspecTermoIntermediacao, _
             enumTipoOperacaoLQS.EspecTermoCobertura, _
             enumTipoOperacaoLQS.CancelamentoEspecificacaoBMA
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            
        Case enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
             enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, _
             enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
        
        Case enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA, _
             enumTipoOperacaoLQS.OperacaoTermoInternaBMA, _
             enumTipoOperacaoLQS.OperacaoCompromissadaInternaBMA

            lngStatusOperacao = enumStatusOperacao.Liquidada

        Case enumTipoOperacaoLQS.CompromissadaEspecificaTermo, _
             enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA

            lngStatusOperacao = enumStatusOperacao.Liquidada

        Case enumTipoOperacaoLQS.DespesasCETIP
            lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica

        Case enumTipoOperacaoLQS.EventosJurosSWAP, _
             enumTipoOperacaoLQS.EventosJurosTERMO, _
             enumTipoOperacaoLQS.EventosCETIP

            If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.SemModalidade Then
                lngStatusOperacao = enumStatusOperacao.Liquidada
            Else
                lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            End If

        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
            
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            lngTipoAcao = enumTipoAcao.EnviadaLDL1002
            
        Case enumTipoOperacaoLQS.CompromissadaGenericaATermo, _
             enumTipoOperacaoLQS.CompromissadaGenericaAVista
            
            lngStatusOperacao = enumStatusOperacao.AConciliarBMA0013

        Case enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoSTR, _
             enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoSTR, _
             enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoSTR, _
             enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoSTR
            
            'RATS 946
            lngStatusOperacao = enumStatusOperacao.Liquidada

        Case enumTipoOperacaoLQS.RegistroOperacoesBMC

            lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica

        Case enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBalcao _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBalcaoAutomatica Then
                lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
                lngTipoAcao = enumTipoAcao.EnviadaBMC0001
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarRegistro Then
                lngStatusOperacao = enumStatusOperacao.CancelamentoSolicitado
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaReativacao Then
                lngStatusOperacao = enumStatusOperacao.LiberadaReativacao
                lngTipoAcao = enumTipoAcao.EnviadaBMC0003
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaReativacaoAutomatica Then
                lngStatusOperacao = enumStatusOperacao.LiberadaReativacaoAutomatica
                lngTipoAcao = enumTipoAcao.EnviadaBMC0003
            Else
                lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            End If

        Case enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica

            lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
        
        'BMF
        Case enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF

            lngStatusOperacao = enumStatusOperacao.ConcordanciaAdmArea
        
        'KIDA - CCR
        Case enumTipoOperacaoLQS.ConsultaOperacaoCCR, _
             enumTipoOperacaoLQS.ConsultaLimitesImportacaoCCR, _
             enumTipoOperacaoLQS.EmissaoOperacaoCCR, _
             enumTipoOperacaoLQS.NegociacaoOperacaoCCR, _
             enumTipoOperacaoLQS.DevolucaoRecolhimentoEstornoReembolsoCCR
             
            lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica

        Case Else
            
            If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
                lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            Else
                lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            End If
    
    End Select

    If lngStatusOperacao > 0 Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            lngStatusOperacao, , lngTipoAcao
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
    End If

    Set objOperacao = Nothing
    Set objRegra = Nothing

    flLiberarAutomatico = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiberarAutomatico Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verifica se a funcionalidade de Liberação de operações pela tela, controlando o status que a operação deve ficar.

Private Function flLiberarViaTela(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                  ByVal pblnIgnoraGradeHorario As Boolean, _
                                  ByVal pblnOrigemConciliacao As Boolean, _
                                  ByVal pblnFluxoAutomatico As Boolean) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim xmlHistSituacao                         As MSXML2.DOMDocument40
Dim lngSequencialISPB                       As Long
Dim lngStatusOperacao                       As Long
Dim lngTipoAcao                             As Long
Dim lngCodigoRetornoVerificacao             As Long

On Error GoTo ErrorHandler

    flLiberarViaTela = False

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If Not pblnIgnoraGradeHorario Then
        If Trim(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text) <> vbNullString Then
            'Verificar a Grade de Horário
            Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
            Set objMensagem = CreateObject("A8LQS.clsMensagem")
            If Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "LDL" Then
                objMensagem.flObterISPBLDL 1, xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, lngSequencialISPB
            Else
                lngSequencialISPB = objMensagem.ObterSequencialISPBDestino(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text)
            End If
            If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                                                      lngSequencialISPB, _
                                                                      CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text)) Then
                If Not pblnOrigemConciliacao Then
                    ' 3023 - Horário limite para envio da Operação excedido.
                    lngCodigoErroNegocio = 3023
                    GoTo ErrorHandler
                Else
                    Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                    objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                 xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                 enumJustificativa.GradeHorario
                    Set objHistSituacaoOperacao = Nothing
                    Set objGradeHorario = Nothing
                    Set objMensagem = Nothing
                    Exit Function
                End If
            End If
            Set objGradeHorario = Nothing
            Set objMensagem = Nothing
        End If
    End If
   
    If pblnOrigemConciliacao Then
        Set objRegra = CreateObject("A8LQS.clsRegra")
        'Verifica regra para liberação automática
        If Not objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Liberar, lngCodigoRetornoVerificacao) Then
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         IIf(lngCodigoRetornoVerificacao <> 0, _
                                                                lngCodigoRetornoVerificacao, _
                                                                enumJustificativa.RegraWorkflow)
            Set objHistSituacaoOperacao = Nothing
            Exit Function
        End If
        Set objRegra = Nothing
                
        If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.RegistroOperacoesBMC _
        And xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica _
        And xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Then
              
            lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                lngStatusOperacao
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
                
        End If
        
    End If

    lngTipoAcao = 0
    lngStatusOperacao = 0
             
    Select Case xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text
        
        Case enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA, _
             enumTipoOperacaoLQS.OperacaoTermoInternaBMA, _
             enumTipoOperacaoLQS.OperacaoCompromissadaInternaBMA

            lngStatusOperacao = enumStatusOperacao.Liquidada

        Case enumTipoOperacaoLQS.DepositoBMA, _
             enumTipoOperacaoLQS.TransferenciasBMCDeposito
            
            lngStatusOperacao = enumStatusOperacao.Liberada
            lngTipoAcao = enumTipoAcao.EnviadaLDL1002
        
        Case enumTipoOperacaoLQS.RetiradaBMA, _
             enumTipoOperacaoLQS.TransferenciasBMCRetirada
            
            lngStatusOperacao = enumStatusOperacao.Liberada
            lngTipoAcao = enumTipoAcao.EnviadaLDL1006
        
        Case enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA
            
            lngStatusOperacao = enumStatusOperacao.Liberada
            lngTipoAcao = enumTipoAcao.EnviadaLDL1016
        
        Case enumTipoOperacaoLQS.TransferenciaBMA
            
            lngStatusOperacao = enumStatusOperacao.Liberada
            lngTipoAcao = enumTipoAcao.EnviadaLDL1016

        Case enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
             enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA, _
             enumTipoOperacaoLQS.DefinitivaDescobertaBMA

            lngStatusOperacao = 0
 
        Case enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
             enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
             enumTipoOperacaoLQS.EspecTermoIntermediacao, _
             enumTipoOperacaoLQS.EspecTermoCobertura, _
             enumTipoOperacaoLQS.CancelamentoEspecificacaoBMA

            lngStatusOperacao = enumStatusOperacao.Liberada

        Case enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
             enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, _
             enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA

            lngStatusOperacao = enumStatusOperacao.Liberada

        Case enumTipoOperacaoLQS.CompromissadaEspecificaTermo, _
             enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA

            lngStatusOperacao = enumStatusOperacao.Liquidada

        Case enumTipoOperacaoLQS.DespesasCETIP
            
            lngStatusOperacao = enumStatusOperacao.Registrada

        Case enumTipoOperacaoLQS.EventosJurosSWAP, _
             enumTipoOperacaoLQS.EventosJurosTERMO, _
             enumTipoOperacaoLQS.EventosCETIP, _
             enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP, _
             enumTipoOperacaoLQS.NETEntradaManualMultilateralCETIP
        
            If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.SemModalidade Then
                lngStatusOperacao = enumStatusOperacao.Liquidada
            Else
                lngStatusOperacao = enumStatusOperacao.Registrada
            End If

        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
            
            lngStatusOperacao = enumStatusOperacao.Liberada
            lngTipoAcao = enumTipoAcao.EnviadaLDL1002
        
        Case enumTipoOperacaoLQS.CompromissadaGenericaATermo, _
             enumTipoOperacaoLQS.CompromissadaGenericaAVista
            
            lngStatusOperacao = enumStatusOperacao.AConciliarBMA0013

        Case enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoCBLC, _
             enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoBMF, _
             enumTipoOperacaoLQS.LiqCorretoraExternaCTributacaoSTR, _
             enumTipoOperacaoLQS.LiqCorretoraExternaSTributacaoSTR, _
             enumTipoOperacaoLQS.LiqCorretoraInternaCTributacaoSTR, _
             enumTipoOperacaoLQS.LiqCorretoraInternaSTributacaoSTR
            
            lngStatusOperacao = 0

        Case enumTipoOperacaoLQS.DespesasBMC
            
            lngStatusOperacao = enumStatusOperacao.PagamentoLib

        Case enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC
            
            lngStatusOperacao = enumStatusOperacao.AConciliar

        Case enumTipoOperacaoLQS.EnvioTEDSTR0006Clientes, _
             enumTipoOperacaoLQS.EnvioTEDSTR0007Clientes, _
             enumTipoOperacaoLQS.EnvioTEDSTR0008Clientes, _
             enumTipoOperacaoLQS.EnvioTEDSTR0009Clientes, _
             enumTipoOperacaoLQS.EnvioTEDSTR0025Clientes, _
             enumTipoOperacaoLQS.EnvioTEDSTR0034Clientes, _
             enumTipoOperacaoLQS.EnvioPAG0105Clientes, _
             enumTipoOperacaoLQS.EnvioPAG0106Clientes, _
             enumTipoOperacaoLQS.EnvioPAG0108Clientes, _
             enumTipoOperacaoLQS.EnvioPAG0109Clientes, _
             enumTipoOperacaoLQS.EnvioPAG0121Clientes, _
             enumTipoOperacaoLQS.EnvioPAG0134Clientes, _
             enumTipoOperacaoLQS.EmissaoTEDPAG0106FdosFIDC, _
             enumTipoOperacaoLQS.EmissaoTEDPAG0108FdosFIDC, _
             enumTipoOperacaoLQS.EmissaoTEDSTR0007FdosFIDC, _
             enumTipoOperacaoLQS.EmissaoTEDSTR0008FdosFIDC
             
             lngStatusOperacao = 0
'Pagamento de Despesas Layout 154 - Colocar STR007 e as demais mensagens
        Case enumTipoOperacaoLQS.EnvioSTR0006PagDespesas, _
             enumTipoOperacaoLQS.EnvioSTR0008PagDespesas, _
             enumTipoOperacaoLQS.EnvioSTR0007PagDespesas, _
             enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente, _
             enumTipoOperacaoLQS.EnvioPagDespesasBoleto, _
             enumTipoOperacaoLQS.EnvioPagDespesasTributos, _
             enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta, _
             enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib, _
             enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta, _
             enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib, _
             enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta, _
             enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib, _
             enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta, _
             enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib, _
             enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta, _
             enumTipoOperacaoLQS.EnvioSTR0007PagDespesasTrib, _
             enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta, _
             enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib
             lngStatusOperacao = 0
                
        Case enumTipoOperacaoLQS.RegistroOperacoesBMC

            lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            
        Case enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBalcao _
            Or xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBalcaoAutomatica Then
                lngStatusOperacao = enumStatusOperacao.Liberada
                lngTipoAcao = enumTipoAcao.EnviadaBMC0001
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarRegistro Then
                lngStatusOperacao = enumStatusOperacao.CancelamentoSolicitado
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaReativacao Then
                lngStatusOperacao = enumStatusOperacao.LiberadaReativacao
                lngTipoAcao = enumTipoAcao.EnviadaBMC0003
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaReativacaoAutomatica Then
                lngStatusOperacao = enumStatusOperacao.LiberadaReativacaoAutomatica
                lngTipoAcao = enumTipoAcao.EnviadaBMC0003
            Else
                lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            End If

        Case enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica
            
            lngStatusOperacao = enumStatusOperacao.RegistradaAutomatica
            
        Case enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais
            
            lngStatusOperacao = enumStatusOperacao.Liberada
                
        'KIDA - CCR
        Case enumTipoOperacaoLQS.ConsultaOperacaoCCR, _
             enumTipoOperacaoLQS.ConsultaLimitesImportacaoCCR, _
             enumTipoOperacaoLQS.EmissaoOperacaoCCR, _
             enumTipoOperacaoLQS.NegociacaoOperacaoCCR, _
             enumTipoOperacaoLQS.DevolucaoRecolhimentoEstornoReembolsoCCR
             
            lngStatusOperacao = enumStatusOperacao.Liberada
            
        'CAM - Layout 249 - Operacao 230
        Case enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega
            
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Then
                lngStatusOperacao = enumStatusOperacao.Liberada
                lngTipoAcao = enumTipoAcao.EnviadaCAM0006
            Else
                lngStatusOperacao = 0
            End If
            
        'CAM - Layout 249 - Operacao 231
        Case enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega
        
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Then
                lngStatusOperacao = enumStatusOperacao.Liberada
                lngTipoAcao = enumTipoAcao.EnviadaCAM0007
            Else
                lngStatusOperacao = 0
            End If
        
        'CAM - Layout 249 - Operacao 232
        Case enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara
        
            lngStatusOperacao = enumStatusOperacao.Liberada
            lngTipoAcao = enumTipoAcao.EnviadaCAM0009
        
        'CAM - Layout 249 - Operacao 233
        Case enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara
        
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Then
                lngStatusOperacao = enumStatusOperacao.Liberada
                lngTipoAcao = enumTipoAcao.EnviadaCAM0010
            Else
                lngStatusOperacao = 0
            End If
        
        'CAM - Layout 249 - Operacao 234
        Case enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega

            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Then
                lngStatusOperacao = enumStatusOperacao.Liberada
                lngTipoAcao = enumTipoAcao.EnviadaCAM0054
            Else
                lngStatusOperacao = 0
            End If
            
        'CAM - Layout 249 - Operacao 237
        Case enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais
        
            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Then
                lngStatusOperacao = enumStatusOperacao.Liberada
                lngTipoAcao = enumTipoAcao.EnviadaCAM0014
            Else
                lngStatusOperacao = 0
            End If
        
        'CAM - Layouts 253, 257, 258, 260 - Operacoes 235, 236, 238, 239, 240 - Sem regras especificas
        Case enumTipoOperacaoLQS.InformaContrArbitParceiroExteriorPaisPropriaIF, _
             enumTipoOperacaoLQS.InformaOperacaoArbitragemParceiroPais, _
             enumTipoOperacaoLQS.CAMInformaContratacaoInterbancarioViaLeilao, _
             enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria, _
             enumTipoOperacaoLQS.ConsultaContratosCambioMercadoInterbancario
                
                lngStatusOperacao = enumStatusOperacao.Liberada
            
        Case Else
            
            If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP _
            Or xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CAM Then
                lngStatusOperacao = enumStatusOperacao.Liberada
            Else
                lngStatusOperacao = enumStatusOperacao.Registrada
            End If
    
    End Select
    
    If lngStatusOperacao > 0 Then
        
        If pblnFluxoAutomatico Then
            Select Case lngStatusOperacao
                Case enumStatusOperacao.Liberada
                    lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
            End Select
        Else
            Select Case lngStatusOperacao
                Case enumStatusOperacao.RegistradaAutomatica
                    lngStatusOperacao = enumStatusOperacao.Registrada
            End Select
        End If
                
        Call objOperacao.AlterarStatusAcao(xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                           lngStatusOperacao, _
                                           xmlRemessa.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, , _
                                           lngTipoAcao)
        
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
    
    End If

    Set objOperacao = Nothing

    flLiberarViaTela = True

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiberarViaTela Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executa a chamada para o geração do registro para o conta corrente, verificando regras especificas por operação.

Private Function flContaCorrente(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                        Optional ByVal blnEnviaRetornoLegado As Boolean = True) As Boolean

Dim objContaCorrente                        As A8LQS.clsContaCorrente

On Error GoTo ErrorHandler

    flContaCorrente = False

    'Para estas operações só envia conta corrente se o Tipo do Titular da Conta for um "PLC"
    If fgTipoOperacaoBMA(CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)) Then
        If Not xmlRemessa.documentElement.selectSingleNode("TP_TITL") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("TP_TITL").Text <> "PLC" Then
                Exit Function
            End If
        End If
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            Exit Function
        Else
            If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.Bruta Then
                'xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
                Exit Function
            End If
        End If
    End If

    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    objContaCorrente.Disponibilizar xmlRemessa, blnEnviaRetornoLegado
    Set objContaCorrente = Nothing

    flContaCorrente = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executa a chamada para o estorno do registro para o conta corrente, verificando regras especificas por operação.

Private Function flEstornarContaCorrente(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objContaCorrente                        As A8LQS.clsContaCorrente

On Error GoTo ErrorHandler

    flEstornarContaCorrente = False

    'Para estas operações só envia conta corrente se o Tipo do Titular da Conta for um "PLC"
    If fgTipoOperacaoBMA(CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)) Then
        If xmlRemessa.documentElement.selectSingleNode("TP_TITL").Text <> "PLC" Then
            Exit Function
        End If
    End If

    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    objContaCorrente.Cancelar xmlRemessa, True
    Set objContaCorrente = Nothing

    flEstornarContaCorrente = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEstornarContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar a chamada dos métodos de envio de estrono para lançamento de conta corrente, PJ, A6, mensagem SPB e retorno para o Legado,
' no cancelamento de uma operação.

Public Function GerenciarCancelamento(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                      ByVal pblnAlterarStatus As Boolean) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objContabilidade                        As A8LQS.clsContabilidade
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objLegado                               As A8LQS.clsLegado
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim lngTipoLiquidacao                       As Long
Dim intCanalVenda                           As Integer
Dim intModalidadeLiquidacao                 As Integer
Dim intTipoContraparte                      As Integer

    On Error GoTo ErrorHandler

    intCanalVenda = 0
    If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
        intCanalVenda = Val(xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text)
    End If
    
    intModalidadeLiquidacao = 99
    If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
        intModalidadeLiquidacao = Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
    End If
    
    intTipoContraparte = 0
    If Not xmlOperacao.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = Val(xmlOperacao.selectSingleNode("//TP_CNPT").Text)
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Conciliada Or _
       xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica Then
        '3025 - Situação da Operação não permite cancelamento.
        lngCodigoErroNegocio = 3025
        GoTo ErrorHandler
    
    ElseIf xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada Or _
           xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Then
        
        If xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
           xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
            '3025 - Situação da Operação não permite cancelamento.
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
        End If
    
    ElseIf xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada Then
        
        If xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text <> enumTipoOperacaoLQS.EventosJurosSWAP Then
            '3025 - Situação da Operação não permite cancelamento.
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
            
        Else
            If intCanalVenda = enumCanalDeVenda.SGC And intModalidadeLiquidacao = enumTipoLiquidacao.SemModalidade Then
                If flVerificarContaCorrente(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) Then
                    
                    If xmlOperacao.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
                        Call fgAppendNode(xmlOperacao, "MESG", "TP_LANC_ITGR", vbNullString)
                    End If
                    
                    xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text = enumTipoLancamentoIntegracao.Estorno
                    
                    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
                    Call objContaCorrente.Disponibilizar(xmlOperacao)
                    Set objContaCorrente = Nothing
                    
                End If
            ElseIf intTipoContraparte <> enumTipoContraparte.Cliente1 Then
                '3025 - Situação da Operação não permite cancelamento.
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
            End If
        
        End If
    
    End If

    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
        If xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DefinitivaCobertaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DefinitivaDescobertaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.OperacaoTermoCobertaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
            xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
            
            If xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada Or _
                xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Or _
                xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarBMA0013 Then
                If Not flConsistirBMA003(xmlOperacao) Then
                    '3109 - Status da Operação não permite cancelamento sem mensagem BMA0003
                    lngCodigoErroNegocio = 3109
                    GoTo ErrorHandler
                End If
            End If
        End If
    End If

    'Altera Status quando vem o cancelamento via Legado
    If pblnAlterarStatus Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.CanceladaOrigem
        
        If Not xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.CanceladaOrigem
        End If
        
        Set objOperacao = Nothing
    End If

    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val(xmlOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.EnvioTEDClientes Then
            If xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.EmSer Then
                Exit Function
            End If
        End If
    End If
    
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Esta seleção de nós está assim pois a estrutura do XML é diferente quando é chamada pela tela e quando
    'é enviada por remessa
    If Not xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
        'lngTipoLiquidacao = CLng(xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    End If
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    strXML = objProcessoOperacao.LerTodos("RecebimentoCancelamento", _
                                          CLng(xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text), _
                                          enumIndicadorSimNao.Sim, , _
                                          lngTipoLiquidacao)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    objProcessoOperacao.EstornarPrevisao xmlOperacao, xmlControleProcessoOperacao

    'KIDA - CBLC - 16/09/2008
    If Not xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                Set objGestaoCaixa = Nothing
            End If
        End If
    End If
        
        
    'NICK - CETIP - ESTORNO PJ/COE - 09/04/2014
    If intCanalVenda = enumCanalDeVenda.SGC And intModalidadeLiquidacao = enumTipoLiquidacao.SemModalidade Then
        If (xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.CanceladaOrigem) Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                Set objGestaoCaixa = Nothing
            End If
        End If
    End If
    
        
    'Esta seleção de nós está assim pois a estrutura do XML é diferente quando é chamada pela tela e quando
    'é enviada por remessa
    If (xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAutomatica Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBackoffice Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAdmArea Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBackofficePrevia Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAdmArea Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarBMA0013) Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            Set objGestaoCaixa = Nothing
        End If
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            objContaCorrente.Cancelar xmlOperacao, True, False
            Set objContaCorrente = Nothing
        End If
        If xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada Or _
            xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Or _
            xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarBMA0013 Then
            If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
                Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
                strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                Set objSubReserva = Nothing
            End If
        End If
        
        'Pikachu - 03/05/2005
        'RATS 244 - PM 1515661
        Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        objMensagemEnviadaRecebida.AlterarStatusBMA0002LDL0001 xmlOperacao.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                               enumStatusMensagem.AConciliar
        Set objMensagemEnviadaRecebida = Nothing
    
    End If
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        Set objLegado = CreateObject("A8LQS.clsLegado")
        strMensagem = objLegado.EnviarRetorno(xmlOperacao)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        Set objLegado = Nothing
    End If

    Set xmlControleProcessoOperacao = Nothing
    Set objProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarCancelamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Verificar a situação do lançamento em conta corrente da operação
'-------------------------------------------------------------------------------------------------

Private Function flVerificarContaCorrente(ByVal pvntNumeroSequencial As Variant) As Boolean

Dim objHistSituacaoIntegracao               As A8LQS.clsHistSituacaoIntegracao
Dim xmlHistSituacaoIntegracao               As MSXML2.DOMDocument40
Dim strHistSituacaoIntegracao               As String

    On Error GoTo ErrorHandler

    flVerificarContaCorrente = False
    
    Set objHistSituacaoIntegracao = CreateObject("A8LQS.clsHistSituacaoIntegracao")

    strHistSituacaoIntegracao = objHistSituacaoIntegracao.Ler( _
                                    "A.NU_SEQU_OPER_ATIV = " & pvntNumeroSequencial & " AND " & _
                                    "A.CO_SITU_PROC IN (" & enumStatusIntegracao.Integrado & ", " & _
                                                            enumStatusIntegracao.IntegradoCC & ")")
                            
    If strHistSituacaoIntegracao <> vbNullString Then
        flVerificarContaCorrente = True
    End If

    Set objHistSituacaoIntegracao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objHistSituacaoIntegracao = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flVerificarContaCorrente", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Controlar o status da operação na liberação de uma mensagem de cancelamento.

Public Function GerenciarCancelamentoMensagem(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean
    
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagem
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim lngTipoLiquidacao                       As Long

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                        , _
                                        enumTipoAcao.CancelamentoSolicitadoComMensagem

    strMensagem = objTipoOperacao.LerTipoOperacaoAcao(CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text), _
                                                      enumTipoAcao.CancelamentoSolicitado)

    If strMensagem = vbNullString Then
        ' 3026 - Não é permitido estornar ou cancelar este tipo de operação.
        lngCodigoErroNegocio = 3026
        GoTo ErrorHandler
    End If
    xmlTipoOperacao.loadXML strMensagem

    If xmlOperacao.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CO_MESG", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_ACAO").Text
    Else
        xmlOperacao.documentElement.selectSingleNode("CO_MESG").Text = xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_ACAO").Text
    End If
    If xmlOperacao.documentElement.selectSingleNode("CodMsg") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CodMsg", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_ACAO").Text
    Else
        xmlOperacao.documentElement.selectSingleNode("CodMsg").Text = xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_ACAO").Text
    End If

    objMensagem.MontarMensagem xmlOperacao, False, False

    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objMensagem = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarCancelamentoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar as chamadas para geração de estorno da operação.

Public Function GerenciarChamadasEstorno(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                         ByVal pstrFuncionalidade As String) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim lngTipoLiquidacao                       As Long
Dim objLegado                               As A8LQS.clsLegado
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    'Carlos - Gerenciamento de chamada das rotinas
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
        'lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    End If
    strXML = objProcessoOperacao.LerTodos(pstrFuncionalidade, _
                                          CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                                          enumIndicadorSimNao.Sim, , _
                                          lngTipoLiquidacao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    'Realizado PJ
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
        flEstornarRealizadoPJ xmlRemessa
    End If

    'Realizado Solicitado A6
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        flEstornarRealizadoSolicitadoA6 xmlRemessa
    End If

    'Lançamento Conta Corrente
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralCBLC Then
            flExcluirContaCorrente xmlRemessa
        Else
            flEstornarContaCorrente xmlRemessa
        End If
    End If

    'mensagem retorno paro Legado
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        Set objLegado = CreateObject("A8LQS.clsLegado")
        strMensagem = objLegado.EnviarRetorno(xmlRemessa)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        Set objLegado = Nothing
    End If

    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarChamadasEstorno Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consistir as informações da mensagem BMA0003.

Private Function flConsistirBMA003(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim strMensagem                             As String
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlMensagemBMA0002                      As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagemBMA0002 = CreateObject("MSXML2.DOMDocument.4.0")

    strMensagem = objMensagemEnviadaRecebida.LerMensagemConciliada(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                   "BMA0002")

    If strMensagem = vbNullString Then
        flConsistirBMA003 = False
        Exit Function
    End If

    xmlMensagem.loadXML strMensagem

    Set xmlMensagemBMA0002 = objMensagemEnviadaRecebida.LerXMLMensagem(xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text)

    If xmlMensagemBMA0002.xml = vbNullString Then
        flConsistirBMA003 = False
        Exit Function
    End If

    xmlMensagem.loadXML strMensagem

    strMensagem = vbNullString
    strMensagem = objMensagemEnviadaRecebida.LerTodos(, , _
                                                      xmlMensagemBMA0002.documentElement.selectSingleNode("//NumCtrlBMA").Text, _
                                                      "BMA0003")

    Set objMensagemEnviadaRecebida = Nothing

    If Trim(strMensagem) = vbNullString Then
        flConsistirBMA003 = False
    Else
        flConsistirBMA003 = True
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConsistirBMA003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controla a liquidação dos eventos de juros CETIP

Private Function flLiquidarEventosJurosCETIP(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim strMensagem                             As String

    On Error GoTo ErrorHandler

    'Enviar Realizado Confirmado para o sistem A6
    flEnviarRealizadoConfirmadoA6 xmlRemessa

    'Lançamento Conta Corrente
    flContaCorrente xmlRemessa
    

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiquidarEventosJurosCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executa a chamada para a exclusao do registro de conta corrente, verificando regras especificas por operação.
Private Function flExcluirContaCorrente(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objContaCorrente                        As A8LQS.clsContaCorrente

On Error GoTo ErrorHandler

    flExcluirContaCorrente = False

    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    objContaCorrente.ExcluirDisponivel xmlRemessa
    Set objContaCorrente = Nothing

    flExcluirContaCorrente = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluirContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verificar inatividade de operações e mensagens em tabela interna
Public Function ObterInatividadeFluxoRecebimento(Optional ByVal plngTP_OPER As Long = 0, _
                                                 Optional ByVal pstrCO_MESG_SPB As String = vbNullString, _
                                                 Optional ByVal plngCO_LOCA_LIQU As Long = 0, _
                                                 Optional ByVal plngCO_SITU_PROC As Long = 0, _
                                                 Optional ByVal plngTP_LIQU_OPER_ATIV As Long = 0) As String

Dim strSQL                                  As String
Dim strWhere                                As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_OPER, CO_MESG_SPB, CO_LOCA_LIQU " & vbNewLine & _
             " FROM     A8.TB_CTRL_FLUX_ATIV               " & vbNewLine

    If plngTP_OPER > 0 Then
        strWhere = strWhere & " AND     TP_OPER    =   " & plngTP_OPER & vbNewLine
    End If

    If pstrCO_MESG_SPB <> vbNullString Then
        strWhere = strWhere & " AND     CO_MESG_SPB    =   '" & pstrCO_MESG_SPB & "'" & vbNewLine
    End If

    If plngCO_LOCA_LIQU > 0 Then
        strWhere = strWhere & " AND     CO_LOCA_LIQU    =   " & plngCO_LOCA_LIQU & vbNewLine
    End If

    If plngCO_SITU_PROC > 0 Then
        strWhere = strWhere & " AND     CO_SITU_PROC    =   " & plngCO_SITU_PROC & vbNewLine
    End If

    If plngTP_LIQU_OPER_ATIV > 0 Then
        strWhere = strWhere & " AND     TP_LIQU_OPER_ATIV    =   " & plngTP_LIQU_OPER_ATIV & vbNewLine
    End If
    
    If strWhere <> vbNullString Then
        strWhere = " WHERE " & Mid$(strWhere, 5)
        strSQL = strSQL & strWhere
    End If

    ObterInatividadeFluxoRecebimento = fgQueryXMLLerTodos("ControleFluxoRecebimento", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterInatividadeFluxoRecebimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------------------------
'Geração de Alerta para o grupo de Usuario
'------------------------------------------------------------------------------------------------------------

Private Sub flGeraAlertaCamara(ByVal pstrCodigoMensagem As String, _
                               ByVal plngTipoBackOffice As Long, _
                               ByVal pstrValor As String, _
                               ByRef xmlOperacao As MSXML2.DOMDocument40)
                               
Dim objAlerta                               As A8LQS.clsAlerta
Dim xmlPropriedadesAlerta                   As MSXML2.DOMDocument40
Dim lngCodFatorGeraAlerta                   As Long
                              
On Error GoTo ErrorHandler
                               
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    Set xmlPropriedadesAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlPropriedadesAlerta.loadXML objAlerta.ObterPropriedades
    
    Select Case pstrCodigoMensagem
        Case "BMC0015NLOCA"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0015NaoLocalizada
        Case "BMC0005NLOCA"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0005NaoLocalizada
        Case Else
            Exit Sub
    End Select
    
    xmlPropriedadesAlerta.selectSingleNode("//NU_SEQU_OPER_ATIV").Text = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    xmlPropriedadesAlerta.selectSingleNode("//CO_FATO_GERA_ALER").Text = lngCodFatorGeraAlerta
    xmlPropriedadesAlerta.selectSingleNode("//TP_BKOF").Text = plngTipoBackOffice
    xmlPropriedadesAlerta.selectSingleNode("//VA_OPER_ATIV").Text = pstrValor
    xmlPropriedadesAlerta.selectSingleNode("//TX_ANEX").Text = xmlOperacao.xml

    Call objAlerta.GerarAlertaCamara(xmlPropriedadesAlerta)
                               
    Set objAlerta = Nothing
    Set xmlPropriedadesAlerta = Nothing

    Exit Sub
ErrorHandler:
    
    Set objAlerta = Nothing
    Set xmlPropriedadesAlerta = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGeraAlertaCamara Private", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Function flEnviarRealizadoMEPJ(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEnviarRealizadoMEPJ = False

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EnviarRealizadoME(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing

    flEnviarRealizadoMEPJ = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoMEPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
