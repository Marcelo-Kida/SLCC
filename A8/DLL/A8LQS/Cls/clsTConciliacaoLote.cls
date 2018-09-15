VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTConciliacaoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variável utilizada para tratamento de erros
Private xmlRetornoErro                      As MSXML2.DOMDocument40
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
Public Function LiberarLoteParaAlcada() As String
Dim rsLotesALiberar                         As ADODB.Recordset
Dim rsOperacao                              As ADODB.Recordset
Dim objRegra                                As A8LQS.clsRegra
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim blnLoteValido                           As Boolean
Dim strLogErro                              As String
Dim blnConfirmacaoAutomatica                As Boolean
Dim blnLiberacaoAutomatica                  As Boolean
Dim blnIntegracaoCCAutomatica               As Boolean
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim intTipoCorretora                        As Integer
Dim blnLoteRecebimentoSemMsgR2Recebida      As Boolean
Dim strNumeroControleIF                     As String
Dim strDataMensagem                         As String
Dim intNumeroSequenciaRepeticao             As Integer

On Error GoTo ErrorHandler

    'instancia objetos
    Set rsLotesALiberar = CreateObject("ADODB.Recordset")
    Set objRegra = CreateObject("A8LQS.clsRegra")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
     
    'obtem Lotes a Liberar
    Set rsLotesALiberar = SelecionarLotesALiberar()
    
    If Not rsLotesALiberar.EOF Then
        rsLotesALiberar.MoveFirst
        Do Until rsLotesALiberar.EOF

            'Validar as Operações do Lote
            blnLoteValido = ValidarOperacoesPorLote(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
            If blnLoteValido Then
            
                'verifica se o workflow (Confirmacao, Liberacao e Integracao C/C) das operacoes que compoem o lote é Automático
                Set rsOperacao = CreateObject("ADODB.Recordset")
                Set rsOperacao = ObterOperacoesPorLote(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
                rsOperacao.MoveFirst
                Do Until rsOperacao.EOF
                    blnConfirmacaoAutomatica = objRegra.VerificarRegraAutomatica2(rsOperacao!NU_SEQU_OPER_ATIV, enumFuncaoSistema.Confirmar)
                    blnLiberacaoAutomatica = objRegra.VerificarRegraAutomatica2(rsOperacao!NU_SEQU_OPER_ATIV, enumFuncaoSistema.Liberar)
                    blnIntegracaoCCAutomatica = objRegra.VerificarRegraAutomatica2(rsOperacao!NU_SEQU_OPER_ATIV, enumFuncaoSistema.IntegracaoCC)
                    Set xmlOperacao = objOperacao.LerXMLOperacao(rsOperacao!NU_SEQU_OPER_ATIV)
                    Exit Do
                Loop
                
                'obter o tipo de corretora (externa ou interna)
                If Not xmlOperacao.selectSingleNode("//TP_IF_CRED_DEBT") Is Nothing Then
                    intTipoCorretora = xmlOperacao.selectSingleNode("//TP_IF_CRED_DEBT").Text
                End If
                
                If blnConfirmacaoAutomatica = True Then
                    
                    'se Corretora = Externa, e Lote = Recebimento/Crédito, então verifica se a mensagem R2 correspondente já foi recebida,
                    'se Não recebeu ainda então aborta o processamento Automático.
                    blnLoteRecebimentoSemMsgR2Recebida = False
                    If rsLotesALiberar!TP_DEBT_CRED_LOTE = enumTipoDebitoCredito.Credito And intTipoCorretora = enumTipoInstituicaoCreditadaDebitada.CorretoraExterna Then
                        blnLoteRecebimentoSemMsgR2Recebida = IIf(ValidarConciliacaoLoteCorretoraMensagemR2(rsLotesALiberar, rsOperacao, strNumeroControleIF, strDataMensagem, intNumeroSequenciaRepeticao) = False, True, False)
                    End If
                    
                    If rsLotesALiberar!TP_DEBT_CRED_LOTE = enumTipoDebitoCredito.Debito Or blnLoteRecebimentoSemMsgR2Recebida = False Then
                    
                        'Confirma todas as operações do lote
                        ProcessarLote rsLotesALiberar, enumAcaoConciliacao.BOConcordar
                        
                        If blnLiberacaoAutomatica = True Then
                        
                            'Libera todas as operações do lote
                            ProcessarLote rsLotesALiberar, enumAcaoConciliacao.AdmAreaLiberar
                            
                            'Marcar Lote como Liberado
                            AlterarStatusLote rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV, enumSituacaoLote.Liberado
                        
                            'se Corretora = Externa, e Lote = Recebimento, então atualiza status da MensagemR2 conciliada para "Liquidada Automática"
                            If intTipoCorretora = enumTipoInstituicaoCreditadaDebitada.CorretoraExterna And rsLotesALiberar!TP_DEBT_CRED_LOTE = enumTipoDebitoCredito.Credito Then
                                Call objMesgRecbEnviSPB.AlterarStatus(Trim(strNumeroControleIF), fgDtHr_To_Xml(strDataMensagem), _
                                                                      enumStatusMensagem.MensagemLiquidadaAutomatica, vbNullString, , False, intNumeroSequenciaRepeticao)

                            End If
                        
                            If blnIntegracaoCCAutomatica = True Then
                                
                                'se Corretora = Externa, e Lote = Pagamento, então não pode Integrar o ContaCorrente neste momento,
                                'visto que este ainda nem foi disponibilizado, e ele será disponibilizado e integrado apenas quando receber a STR0007R1.
                                If Not (intTipoCorretora = enumTipoInstituicaoCreditadaDebitada.CorretoraExterna And rsLotesALiberar!TP_DEBT_CRED_LOTE = enumTipoDebitoCredito.Debito) Then
                                    Call IntegrarContaCorrente(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
                                End If
                                
                            End If
                        End If
                    End If
                End If
            End If
            rsLotesALiberar.MoveNext
        Loop
    End If
    
    Set rsLotesALiberar = Nothing
    Set rsOperacao = Nothing
    Set objRegra = Nothing
    Set objOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

Exit Function
ErrorHandler:
    Set rsLotesALiberar = Nothing
    Set rsOperacao = Nothing
    Set objRegra = Nothing
    Set objOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing
    
    strLogErro = Err.Description
    
    fgGravaArquivo App.EXEName & "." & "LiberarLoteParaAlcada", strLogErro
    
End Function
Public Function AlterarStatusLote(ByVal plngIdLote As Long, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv As Date, ByVal pintSituacaoLote As Integer)
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    Dim objConciliacao                          As A8LQS.clsConciliacao
    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    AlterarStatusLote = objConciliacao.AlterarStatusLote(plngIdLote, pstrSiglaSistema, pdtDataOperAtiv, pintSituacaoLote)
    Set objConciliacao = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:
    Set objConciliacao = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
Public Function ProcessarLote(prsLotesALiberar As ADODB.Recordset, ByVal pintAcaoProcessamento As Integer)
Dim rsSomaLote                  As ADODB.Recordset
Dim rsOperacao                  As ADODB.Recordset
Dim strXMLProc                  As String
Dim strResultadoOperacao        As String

    On Error GoTo ErrorHandler


    Set rsSomaLote = CreateObject("ADODB.Recordset")
    Set rsSomaLote = ObterSomatoriaOperacaoPorLote(prsLotesALiberar!CD_LOTE, prsLotesALiberar!SG_SIST, prsLotesALiberar!DT_OPER_ATIV)
    
    'Obter as operações do Lote
    Set rsOperacao = CreateObject("ADODB.Recordset")
    Set rsOperacao = ObterUltimaOperacaoPorLote(prsLotesALiberar!CD_LOTE, prsLotesALiberar!SG_SIST, prsLotesALiberar!DT_OPER_ATIV)
    strXMLProc = flMontarXMLProcessamento(rsOperacao, rsSomaLote)
                
    Set rsOperacao = Nothing
    Set rsSomaLote = Nothing
                
    'Concordar o Lote
    flProcessar strXMLProc, pintAcaoProcessamento
    Exit Function
ErrorHandler:
    Set rsSomaLote = Nothing
    Set rsOperacao = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessarLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'Enviar itens de mensagem e operações para liquidação
Public Function flProcessar(ByVal pstrXMLProc As String, ByVal pintAcaoProcessamento As Integer)

Dim strXMLRetorno                           As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

        
    If pstrXMLProc <> vbNullString Then
                
        'Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        'strXMLRetorno = objOperacaoMensagem.LiquidarCorretoras(intAcaoProcessamento, _
        '                                                       strXMLProc, _
        '                                                       vntCodErro, _
        '                                                       vntMensagemErro)
        
        strXMLRetorno = LiquidarLoteCorretoras(pintAcaoProcessamento, _
                                                pstrXMLProc, _
                                                vntCodErro, _
                                                vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
    End If
    
    If strXMLRetorno <> vbNullString Then
        Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetornoErro.loadXML(strXMLRetorno)
    Else
        Set xmlRetornoErro = Nothing
    End If
    
    flProcessar = strXMLRetorno
    
Exit Function
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
End Function

'Monta XML com as chaves das operações que serão processadas
Public Function flMontarXMLProcessamento(ByVal rsOperacao As ADODB.Recordset, ByVal rsSomaLote As ADODB.Recordset) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer
Dim strItemKey                              As String
Dim rsOperacoes                             As ADODB.Recordset
    
On Error GoTo ErrorHandler
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")
    
    Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_EMPR", _
                                       rsOperacao!CO_EMPR)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_LOCA_LIQU", _
                                       rsOperacao!CO_LOCA_LIQU)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_ISPB_CNPT", _
                                       rsOperacao!CO_ISPB_BANC_LIQU_CNPT)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_CNPJ_CNPT", _
                                       rsOperacao!CO_CNPJ_CNPT)
                                       
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "NO_CNPT", _
                                       rsOperacao!NO_CNPT)
                                       
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "TP_IF_CRED_DEBT", _
                                       rsOperacao!TP_IF_CRED_DEBT)
    
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_AGEN_COTR", _
                                       rsOperacao!CO_AGEN_COTR)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "NU_CC_COTR", _
                                       rsOperacao!NU_CC_COTR)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "ID_PGTO_RECB_GRUP", _
                                       IIf(rsSomaLote!VL_TOTL_OPER < 0, "P", "R"))
                                       
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "VA_LANC", _
                                       fgVlr_To_Xml(rsOperacao!VL_TOTL_LOTE))
            
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "Repeat_Operacao", _
                                       "")
                                       
    Set rsOperacoes = CreateObject("ADODB.Recordset")
    Set rsOperacoes = ObterOperacoesPorLote(rsOperacao!CD_LOTE, rsOperacao!SG_SIST, rsOperacao!DT_OPER_ATIV)
    
    rsOperacoes.MoveFirst
    Do Until rsOperacoes.EOF
        Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "NU_SEQU_OPER_ATIV", _
                                           rsOperacoes!NU_SEQU_OPER_ATIV)
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "DH_ULTI_ATLZ", _
                                           fgDtHr_To_Xml(rsOperacoes!dh_ulti_atlz))
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "CO_ULTI_SITU_PROC", _
                                            rsOperacoes!CO_ULTI_SITU_PROC)
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "ID_PGTO_RECB", _
                                            IIf(rsSomaLote!VL_TOTL_OPER < 0, "P", "R"))
                    
        Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)
                    
        Set xmlItemOperacao = Nothing
        rsOperacoes.MoveNext
    Loop
    
    Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
    Set xmlItemEnvioMsg = Nothing

    
    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLProcessamento = vbNullString
    Else
        flMontarXMLProcessamento = xmlProcessamento.xml
    End If
    
    Set xmlProcessamento = Nothing
    
Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLProcessamento", 0

End Function
Public Function ObterOperacoesPorLote(ByVal plngIdLote As Long, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As ADODB.Recordset
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    On Error GoTo ErrorHandler
    
        Set rsOperacoes = CreateObject("ADODB.Recordset")
        strSQL = "Select O.*, " & vbCrLf & _
                 "L.TP_DEBT_CRED_LOTE, " & vbCrLf & _
                 "L.VL_TOTL_LOTE, " & vbCrLf & _
                 "L.QT_TOTL_LOTE, " & vbCrLf & _
                 "L.IN_LIQU_LOTE " & vbCrLf & _
                 " From  A8.TB_OPER_ATIV O, " & vbCrLf & _
                 "       A8.TB_LOTE_OPER_ATIV L  " & vbCrLf & _
                 " Where O.CD_LOTE = L.CD_LOTE " & vbCrLf & _
                 " And   O.SG_SIST = L.SG_SIST " & vbCrLf & _
                 " And   O.DT_OPER_ATIV = L.DT_OPER_ATIV " & vbCrLf & _
                 " And   L.CD_LOTE = " & plngIdLote & vbCrLf & _
                 " And   L.SG_SIST         =   '" & pstrSiglaSistema & "' " & vbCrLf & _
                 " And   L.DT_OPER_ATIV    =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv))
                 
                 
        Set rsOperacoes = fgQuerySQL(strSQL)
        
        Set ObterOperacoesPorLote = rsOperacoes
        
        Set rsOperacoes = Nothing
        
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    
    Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacoesPorLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
Private Function ObterUltimaOperacaoPorLote(ByVal plngIdLote As Long, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As ADODB.Recordset
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    On Error GoTo ErrorHandler
    
        Set rsOperacoes = CreateObject("ADODB.Recordset")
        strSQL = "Select O.*, " & vbCrLf & _
                 "L.TP_DEBT_CRED_LOTE, " & vbCrLf & _
                 "L.VL_TOTL_LOTE, " & vbCrLf & _
                 "L.QT_TOTL_LOTE, " & vbCrLf & _
                 "L.IN_LIQU_LOTE " & vbCrLf & _
                 " From  A8.TB_OPER_ATIV O, " & vbCrLf & _
                 "       A8.TB_LOTE_OPER_ATIV L  " & vbCrLf & _
                 " Where O.CD_LOTE = L.CD_LOTE " & vbCrLf & _
                 " And   O.SG_SIST = L.SG_SIST " & vbCrLf & _
                 " And   O.DT_OPER_ATIV = L.DT_OPER_ATIV " & vbCrLf & _
                 " And   L.CD_LOTE = " & plngIdLote & vbCrLf & _
                 " And   L.SG_SIST         =   '" & pstrSiglaSistema & "' " & vbCrLf & _
                 " And   L.DT_OPER_ATIV    =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv)) & vbCrLf & _
                 " And   rownum = 1 " & vbCrLf & _
                 " Order by O.dh_ulti_atlz desc "
                 
        Set rsOperacoes = fgQuerySQL(strSQL)
        
        Set ObterUltimaOperacaoPorLote = rsOperacoes
        
        Set rsOperacoes = Nothing
        
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    
    Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterUltimaOperacaoPorLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Public Function ObterSomatoriaOperacaoPorLote(ByVal plngIdLote As Long, _
                                               ByVal pstrSiglaSistema As String, _
                                               ByVal pdtDataOperAtiv, _
                                      Optional ByVal pblnValidaStatusOperacoes As Boolean = False) As ADODB.Recordset

Dim rsLote                                  As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    strSQL = " select decode(l.tp_debt_cred_lote,1,l.vl_totl_lote*-1,l.vl_totl_lote) vl_totl_lote, " & vbCrLf & _
             "        l.qt_totl_lote, " & vbCrLf & _
             "        sum (decode(o.in_oper_debt_cred,2,o.va_oper_ativ*-1,o.va_oper_ativ)) vl_totl_oper, " & vbCrLf & _
             "        count(o.co_oper_ativ) QT_OPER " & vbCrLf & _
             " from a8.tb_oper_ativ o, " & vbCrLf & _
             "      a8.tb_lote_oper_ativ l " & vbCrLf & _
             " where o.cd_lote = l.cd_lote " & vbCrLf & _
             "   and o.dt_oper_ativ = l.dt_oper_ativ " & vbCrLf & _
             "   and o.sg_sist = l.sg_sist " & vbCrLf & _
             "   and l.cd_lote = " & plngIdLote & vbCrLf & _
             "   and l.sg_sist = '" & pstrSiglaSistema & "' " & vbCrLf & _
             "   and l.dt_oper_ativ = " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv)) & vbCrLf
   
    If pblnValidaStatusOperacoes = True Then
        strSQL = strSQL & " and o.co_ulti_situ_proc in (" & enumStatusOperacao.EmSer & "," & enumStatusOperacao.ManualEmSer & ") " & vbCrLf
    End If
    
    strSQL = strSQL & " group by decode(l.tp_debt_cred_lote,1,l.vl_totl_lote*-1,l.vl_totl_lote), " & vbCrLf & _
                      "          l.qt_totl_lote "
        
    Set rsLote = CreateObject("ADODB.Recordset")
    Set rsLote = fgQuerySQL(strSQL)
    Set ObterSomatoriaOperacaoPorLote = rsLote
        
    Set rsLote = Nothing
        
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
        
Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSomatoriaOperacaoPorLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function ValidarOperacoesPorLote(ByVal plngIdLote As Long, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As Boolean
Dim strSQL                                  As String
Dim rsLote                                  As ADODB.Recordset
Dim intDifQuantidade                        As Integer
Dim decDifValor                             As Currency

    On Error GoTo ErrorHandler
    
    Set rsLote = CreateObject("ADODB.Recordset")
    Set rsLote = ObterSomatoriaOperacaoPorLote(plngIdLote, pstrSiglaSistema, pdtDataOperAtiv, True)
    
    If Not rsLote.EOF Then
        intDifQuantidade = (rsLote!QT_TOTL_LOTE) - CInt(rsLote!QT_OPER)
        decDifValor = CDec(rsLote!VL_TOTL_LOTE) - CDec(rsLote!VL_TOTL_OPER)
    Else
        ValidarOperacoesPorLote = False
        Exit Function
    End If
    
    'Se houver diferenças do valor e da quantidade no lote não liberar
    If intDifQuantidade <> 0 Or decDifValor <> 0 Then
        ValidarOperacoesPorLote = False
        Exit Function
    End If
    ValidarOperacoesPorLote = True
    
    Set rsLote = Nothing
    
    Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ValidarOperacoesPorLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function
Public Function SelecionarLotesALiberar() As ADODB.Recordset
Dim strSQL                                  As String
Dim rsLotes                                 As ADODB.Recordset
    On Error GoTo ErrorHandler

        Set rsLotes = CreateObject("ADODB.Recordset")
        strSQL = "Select * " & vbCrLf & _
                 " From  A8.TB_LOTE_OPER_ATIV " & vbCrLf & _
                 " Where IN_LIQU_LOTE = " & enumSituacaoLote.ALiberar & vbCrLf & _
                 " and   DT_OPER_ATIV = TRUNC(SYSDATE)"
                 
        Set rsLotes = fgQuerySQL(strSQL)
            
        Set SelecionarLotesALiberar = rsLotes
        
        Set rsLotes = Nothing
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    Exit Function
ErrorHandler:
    Set rsLotes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "SelecionarLotesALiberar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método instancia a classe clsTConciliacao para a liquidação com corretoras, de operações e mensagens,
' na camada de negócios, e retorna para a camada de interface.
' Também devolve à camada de interface o resultado do processamento em lote.
Public Function LiquidarLoteCorretoras(ByVal pintAcao As Integer, _
                                   ByVal strXMLOperacoes As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlDocProcessar                         As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40

Dim strxmlErro                              As String
Dim lngCodErro                              As Long
Dim intTipoErro                             As Integer
Dim strErro                                 As String

Dim blnIgnoraGradeHorario                   As Boolean

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    'Monta XML para controle de erros
    Set xmlControleErro = CreateObject("MSXML2.DOMDocument.4.0")

    'Monta XML para controle de resultados
    Set xmlControleResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlControleResultado, "", "Repeat_ControleErro", "")

    Set xmlDocProcessar = CreateObject("MSXML2.DOMDocument.4.0")
    xmlDocProcessar.loadXML strXMLOperacoes
    For Each xmlDomNode In xmlDocProcessar.documentElement.childNodes
        'blnIgnoraGradeHorario = Val(xmlDomNode.selectSingleNode("IgnoraGradeHorario").Text)

        'Incrementa o controle de erros com a MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", 0, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", vbNullString, "Repeat_ControleErro")

        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_LOCA_LIQU", xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_ISPB_CNPT", xmlDomNode.selectSingleNode("CO_ISPB_CNPT").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_CNPJ_CNPT", xmlDomNode.selectSingleNode("CO_CNPJ_CNPT").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TP_IF_CRED_DEBT", xmlDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_AGEN_COTR", xmlDomNode.selectSingleNode("CO_AGEN_COTR").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "NU_CC_COTR", xmlDomNode.selectSingleNode("NU_CC_COTR").Text, "Repeat_ControleErro")
                        
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        strxmlErro = objConciliacao.LiquidarCorretoras(pintAcao, xmlDomNode, True, True)
        

        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)

            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 0, "Repeat_ControleErro")

            strErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Description").Text

            'Incrementa o controle de erros com o STATUS do processo (ERRO)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoErro", intTipoErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", strErro, "Repeat_ControleErro")
        Else
            lngCodErro = 0

            'Incrementa o controle de erros com o STATUS do processo (OK)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", "OK", "Repeat_ControleErro")
        End If
    Next

    LiquidarLoteCorretoras = xmlControleResultado.xml

    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    Set xmlDocProcessar = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objConciliacao = Nothing
    Set xmlDocProcessar = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

Private Function ObterSomatoriaOperacaoPorContaCorrente(ByVal plngIdLote As Long, _
                                                        ByVal pstrSiglaSistema As String, _
                                                        ByVal pdtDataOperAtiv As Date) As ADODB.Recordset

Dim rsLote                                  As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler

'                    strListItemKey = "|" & objDomNode.selectSingleNode("CO_EMPR").Text & _
'                                 "|" & objDomNode.selectSingleNode("DT_OPER").Text & _
'                                 "|" & objDomNode.selectSingleNode("TP_OPER").Text & _
'                                 "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
'                                 "|" & objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & _
'                                 "|" & objDomNode.selectSingleNode("CO_BANC").Text & _
'                                 "|" & objDomNode.selectSingleNode("CO_AGEN").Text & _
'                                 "|" & objDomNode.selectSingleNode("NU_CC").Text & _
'                                 "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
    
    strSQL = " select o.co_empr, " & vbCrLf & _
             "        o.TP_OPER, " & vbCrLf & _
             "        o.CO_VEIC_LEGA, " & vbCrLf & _
             "        o.CO_LOCA_LIQU, " & vbCrLf & _
             "        l.CO_AGEN, " & vbCrLf & _
             "        l.NU_CC, " & vbCrLf & _
             "        o.CO_ULTI_SITU_PROC, " & vbCrLf & _
             "        sum (decode(o.in_oper_debt_cred,2,o.va_oper_ativ*-1,o.va_oper_ativ)) vl_totl_oper " & vbCrLf & _
             " from a8.tb_oper_ativ o, a8.tb_lanc_cc_cntb l " & vbCrLf & _
             " where o.nu_sequ_oper_ativ = l.nu_sequ_oper_ativ and l.nr_sequ_lanc = 0 and o.cd_lote      = " & plngIdLote & vbCrLf & _
             "   and o.sg_sist      = '" & pstrSiglaSistema & "' " & vbCrLf & _
             "   and o.dt_oper_ativ = " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv)) & vbCrLf
    
    strSQL = strSQL & " group by o.co_empr, " & vbCrLf & _
                      "          o.TP_OPER, " & vbCrLf & _
                      "          o.CO_VEIC_LEGA, " & vbCrLf & _
                      "          o.CO_LOCA_LIQU, " & vbCrLf & _
                      "          l.CO_AGEN, " & vbCrLf & _
                      "          l.NU_CC, " & vbCrLf & _
                      "          o.CO_ULTI_SITU_PROC " & vbCrLf

    Set rsLote = CreateObject("ADODB.Recordset")
    Set rsLote = fgQuerySQL(strSQL)
    Set ObterSomatoriaOperacaoPorContaCorrente = rsLote
        
    Set rsLote = Nothing
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSomatoriaOperacaoPorContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function ObterOperacoesPorLoteContaCorrente(ByVal plngIdLote As Long, _
                                                    ByVal pstrSiglaSistema As String, _
                                                    ByVal pdtDataOperAtiv, _
                                                    ByVal pintCO_EMPR As Integer, _
                                                    ByVal pintTP_OPER As Integer, _
                                                    ByVal pstrCO_VEIC_LEGA As String, _
                                                    ByVal pintCO_LOCA_LIQU As Integer, _
                                                    ByVal pintCO_AGEN As Integer, _
                                                    ByVal plngNU_CC As Long, _
                                                    ByVal pintCO_ULTI_SITU_PROC As Integer) As ADODB.Recordset
    
Dim strSQL                                  As String
Dim rsOperacoes                             As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsOperacoes = CreateObject("ADODB.Recordset")
    strSQL = "Select o.nu_sequ_oper_ativ " & vbCrLf & _
             " From  A8.TB_OPER_ATIV o, a8.tb_lanc_cc_cntb l " & vbCrLf & _
             " Where o.nu_sequ_oper_ativ = l.nu_sequ_oper_ativ and l.nr_sequ_lanc = 0 And o.CD_LOTE           =  " & plngIdLote & vbCrLf & _
             "   And o.SG_SIST           = '" & pstrSiglaSistema & "' " & vbCrLf & _
             "   And o.DT_OPER_ATIV      =  " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv)) & vbCrLf & _
             "   And o.CO_EMPR           =  " & pintCO_EMPR & vbCrLf & _
             "   And o.TP_OPER           =  " & pintTP_OPER & vbCrLf & _
             "   And o.CO_VEIC_LEGA      = '" & pstrCO_VEIC_LEGA & "' " & vbCrLf & _
             "   And o.CO_LOCA_LIQU      =  " & pintCO_LOCA_LIQU & vbCrLf & _
             "   And l.CO_AGEN      =  " & pintCO_AGEN & vbCrLf & _
             "   And l.NU_CC        =  " & plngNU_CC & vbCrLf & _
             "   And o.CO_ULTI_SITU_PROC =  " & pintCO_ULTI_SITU_PROC & vbCrLf

    Set rsOperacoes = fgQuerySQL(strSQL)
    
    Set ObterOperacoesPorLoteContaCorrente = rsOperacoes
    
    Set rsOperacoes = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacoesPorLoteContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Sub IntegrarContaCorrente(ByVal plngIdLote As Long, _
                                 ByVal pstrSiglaSistema As String, _
                                 ByVal pdtDataOperAtiv As Date)

Dim objFluxoProcessamento       As A8LQS.clsFluxoProcessamento
Dim objContaCorrente            As A8LQS.clsContaCorrente
Dim objOperacao                 As A8LQS.clsOperacao
Dim rsSomaNetCC                 As ADODB.Recordset
Dim rsOperacoesNetCC            As ADODB.Recordset
Dim xmlOperacao                 As MSXML2.DOMDocument40
Dim strNetOperacoes             As String
Dim intDebitoCredito            As Integer
Dim dblValorNetCC               As Double
Dim lngNU_SEQU_OPER_ATIV        As Long

On Error GoTo ErrorHandler

    'obtem somatorio de operações por conta corrente
    Set rsSomaNetCC = CreateObject("ADODB.Recordset")
    Set rsSomaNetCC = ObterSomatoriaOperacaoPorContaCorrente(plngIdLote, pstrSiglaSistema, pdtDataOperAtiv)
    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    If Not rsSomaNetCC.EOF Then
        rsSomaNetCC.MoveFirst
        Do Until rsSomaNetCC.EOF
        
            dblValorNetCC = rsSomaNetCC!VL_TOTL_OPER
            
            If dblValorNetCC <> 0 Then
            
                If dblValorNetCC > 0 Then
                    intDebitoCredito = enumTipoDebitoCredito.Credito
                Else
                    intDebitoCredito = enumTipoDebitoCredito.Debito
                End If
            
                'obter NetOperacoes
                Set rsOperacoesNetCC = CreateObject("ADODB.Recordset")
                Set rsOperacoesNetCC = ObterOperacoesPorLoteContaCorrente(plngIdLote, _
                                                                          pstrSiglaSistema, _
                                                                          pdtDataOperAtiv, _
                                                                          rsSomaNetCC!CO_EMPR, _
                                                                          rsSomaNetCC!TP_OPER, _
                                                                          rsSomaNetCC!CO_VEIC_LEGA, _
                                                                          rsSomaNetCC!CO_LOCA_LIQU, _
                                                                          rsSomaNetCC!CO_AGEN, _
                                                                          rsSomaNetCC!NU_CC, _
                                                                          rsSomaNetCC!CO_ULTI_SITU_PROC)
                
                strNetOperacoes = vbNullString
                rsOperacoesNetCC.MoveFirst
                lngNU_SEQU_OPER_ATIV = 0
                Do Until rsOperacoesNetCC.EOF
                    If lngNU_SEQU_OPER_ATIV = 0 Then
                        lngNU_SEQU_OPER_ATIV = rsOperacoesNetCC!NU_SEQU_OPER_ATIV
                    End If
                    strNetOperacoes = strNetOperacoes & "|" & rsOperacoesNetCC!NU_SEQU_OPER_ATIV
                    rsOperacoesNetCC.MoveNext
                Loop
                
                'integrar net por conta corrente
                Call objContaCorrente.Integrar(lngNU_SEQU_OPER_ATIV, _
                                               enumIntegrarCCContabilidade.ContaCorrente, _
                                               False, _
                                               enumTipoLancamentoIntegracao.Normal, _
                                               "", _
                                               Abs(dblValorNetCC), _
                                               intDebitoCredito, _
                                               strNetOperacoes, _
                                               0, _
                                               True)
                                               
                '--------------------------------------------------------------------
                'envia retorno para o legado de cada operacao que compõe
                'o net de c/c que acabou de ser integrado acima
                '--------------------------------------------------------------------
                rsOperacoesNetCC.MoveFirst
                Do Until rsOperacoesNetCC.EOF
                    lngNU_SEQU_OPER_ATIV = rsOperacoesNetCC!NU_SEQU_OPER_ATIV
                    Set xmlOperacao = objOperacao.LerXMLOperacao(lngNU_SEQU_OPER_ATIV)
                    
                    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
                    If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
                        Call fgAppendNode(xmlOperacao, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
                    End If
                    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.EnviadoCCAutomatico
                    Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao, False)
                    
                    rsOperacoesNetCC.MoveNext
                Loop
                '--------------------------------------------------------------------
                                                                                                                      
            End If
            rsSomaNetCC.MoveNext
        Loop
    End If

    Set objFluxoProcessamento = Nothing
    Set objContaCorrente = Nothing
    Set objOperacao = Nothing
    Set rsSomaNetCC = Nothing
    Set rsOperacoesNetCC = Nothing
    Set xmlOperacao = Nothing

Exit Sub
ErrorHandler:
    Set objFluxoProcessamento = Nothing
    Set objContaCorrente = Nothing
    Set objOperacao = Nothing
    Set rsSomaNetCC = Nothing
    Set rsOperacoesNetCC = Nothing
    Set xmlOperacao = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "Sub IntegrarContaCorrente", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub
'processamento a partir do Lote de Operacoes, portanto verificar se há Mensagem R2 correspondente
Private Function ValidarConciliacaoLoteCorretoraMensagemR2(ByVal prsLoteRecebimento As ADODB.Recordset, _
                                                           ByVal prsOperacao As ADODB.Recordset, _
                                                           ByRef pstrNumeroControleIF As String, _
                                                           ByRef pstrDataMensagem As String, _
                                                           ByRef pintNumeroSequenciaRepeticao As Integer) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim rsMensagensR2                           As ADODB.Recordset
Dim strLoteISPBIFDebitada                   As String
Dim strLoteNomeRemetDebitado                As String
Dim strLoteCNPJCPFDebitado                  As String
Dim lngLoteAgenciaDebitada                  As Long
Dim lngLoteContaDebitada                    As Long
Dim lngLoteLocalLiquidacao                  As Long
Dim dblLoteValor                            As Double
Dim strMsgR2ISPBIFDebitada                   As String
Dim strMsgR2NomeRemetDebitado                As String
Dim strMsgR2CNPJCPFDebitado                  As String
Dim lngMsgR2AgenciaDebitada                  As Long
Dim lngMsgR2ContaDebitada                    As Long
Dim lngMsgR2LocalLiquidacao                  As Long
Dim dblMsgR2Valor                            As Double

On Error GoTo ErrorHandler
    
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set rsMensagensR2 = objMensagemEnviadaRecebida.LerXMLMensagensPagamentoCorretoras()
     
    'verifica MensagemR2 a MensagemR2 se concilia com o Lote de Recebimento
    If Not rsMensagensR2.EOF Then
        
        'obtem informações do Lote necessárias para verificar conciliação
        strLoteISPBIFDebitada = IIf(IsNull(prsOperacao!CO_ISPB_BANC_LIQU_CNPT), "", prsOperacao!CO_ISPB_BANC_LIQU_CNPT)
        strLoteCNPJCPFDebitado = IIf(IsNull(prsOperacao!CO_CNPJ_CNPT), "", prsOperacao!CO_CNPJ_CNPT)
        strLoteNomeRemetDebitado = IIf(IsNull(prsOperacao!NO_CNPT), "", prsOperacao!NO_CNPT)
        lngLoteAgenciaDebitada = IIf(IsNull(prsOperacao!CO_AGEN_COTR), "", prsOperacao!CO_AGEN_COTR)
        lngLoteContaDebitada = IIf(IsNull(prsOperacao!NU_CC_COTR), "", prsOperacao!NU_CC_COTR)
        lngLoteLocalLiquidacao = IIf(IsNull(prsOperacao!CO_LOCA_LIQU), "", prsOperacao!CO_LOCA_LIQU)
        dblLoteValor = IIf(IsNull(prsLoteRecebimento!VL_TOTL_LOTE), "", prsOperacao!VL_TOTL_LOTE)
        
        rsMensagensR2.MoveFirst
        Do Until rsMensagensR2.EOF
            
            strMsgR2ISPBIFDebitada = IIf(IsNull(rsMensagensR2!CO_ISPB_BANC_LIQU_CNPT), "", rsMensagensR2!CO_ISPB_BANC_LIQU_CNPT)
            strMsgR2CNPJCPFDebitado = IIf(IsNull(rsMensagensR2!CO_CNPJ_CNPT), "", rsMensagensR2!CO_CNPJ_CNPT)
            strMsgR2NomeRemetDebitado = IIf(IsNull(rsMensagensR2!NO_CNPT), "", rsMensagensR2!NO_CNPT)
            lngMsgR2AgenciaDebitada = IIf(IsNull(rsMensagensR2!CO_AGEN_COTR), "", rsMensagensR2!CO_AGEN_COTR)
            lngMsgR2ContaDebitada = IIf(IsNull(rsMensagensR2!NU_CC_COTR), "", rsMensagensR2!NU_CC_COTR)
            lngMsgR2LocalLiquidacao = IIf(IsNull(rsMensagensR2!CO_LOCA_LIQU), "", rsMensagensR2!CO_LOCA_LIQU)
            dblMsgR2Valor = IIf(IsNull(rsMensagensR2!VA_FINC), "", rsMensagensR2!VA_FINC)
            
            Select Case rsMensagensR2!CO_MESG_SPB
                Case "STR0008R2", "PAG0108R2"
                    If Val(strLoteISPBIFDebitada) = Val(strMsgR2ISPBIFDebitada) _
                    And Val(strLoteCNPJCPFDebitado) = Val(strMsgR2CNPJCPFDebitado) _
                    And Trim(UCase(strLoteNomeRemetDebitado)) = Trim(UCase(strMsgR2NomeRemetDebitado)) _
                    And lngLoteAgenciaDebitada = Val(lngMsgR2AgenciaDebitada) _
                    And lngLoteContaDebitada = Val(lngMsgR2ContaDebitada) _
                    And lngLoteLocalLiquidacao = lngMsgR2LocalLiquidacao _
                    And dblLoteValor = dblMsgR2Valor Then
                        ValidarConciliacaoLoteCorretoraMensagemR2 = True
                        pstrNumeroControleIF = Trim(rsMensagensR2!NU_CTRL_IF)
                        pstrDataMensagem = Trim(rsMensagensR2!DH_REGT_MESG_SPB)
                        pintNumeroSequenciaRepeticao = Val(rsMensagensR2!NU_SEQU_CNTR_REPE)
                        Exit Do
                    End If
                Case "STR0006R2"
                    If Val(strLoteISPBIFDebitada) = Val(strMsgR2ISPBIFDebitada) _
                    And Val(strLoteCNPJCPFDebitado) = Val(strMsgR2CNPJCPFDebitado) _
                    And Trim(UCase(strLoteNomeRemetDebitado)) = Trim(UCase(strMsgR2NomeRemetDebitado)) _
                    And lngLoteAgenciaDebitada = Val(lngMsgR2AgenciaDebitada) _
                    And lngLoteLocalLiquidacao = lngMsgR2LocalLiquidacao _
                    And dblLoteValor = dblMsgR2Valor Then
                        ValidarConciliacaoLoteCorretoraMensagemR2 = True
                        pstrNumeroControleIF = Trim(rsMensagensR2!NU_CTRL_IF)
                        pstrDataMensagem = Trim(rsMensagensR2!DH_REGT_MESG_SPB)
                        pintNumeroSequenciaRepeticao = Val(rsMensagensR2!NU_SEQU_CNTR_REPE)
                        Exit Do
                    End If
                Case "STR0007R2"
                    If Val(strLoteISPBIFDebitada) = Val(strMsgR2ISPBIFDebitada) _
                    And lngLoteLocalLiquidacao = lngMsgR2LocalLiquidacao _
                    And dblLoteValor = dblMsgR2Valor Then
                        ValidarConciliacaoLoteCorretoraMensagemR2 = True
                        pstrNumeroControleIF = Trim(rsMensagensR2!NU_CTRL_IF)
                        pstrDataMensagem = Trim(rsMensagensR2!DH_REGT_MESG_SPB)
                        pintNumeroSequenciaRepeticao = Val(rsMensagensR2!NU_SEQU_CNTR_REPE)
                        Exit Do
                    End If
            End Select
            
            rsMensagensR2.MoveNext
        Loop
     End If
    
    Set objMensagemEnviadaRecebida = Nothing
    Set rsMensagensR2 = Nothing
    
Exit Function
ErrorHandler:

    Set objMensagemEnviadaRecebida = Nothing
    Set rsMensagensR2 = Nothing

    Call fgRaiseError(App.EXEName, TypeName(Me), "Function ValidarConciliacaoLoteCorretoraMensagemR2", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


