VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : A8Liquidacao\r\nClasse         : clsOperacao\r\nData Criação   : 02-07-2001 10:26\r\nObjetivo       : Controle Transacional de acesso ao MQSeries/Oracle, distribuição\r\n                 das Mensagens aos objetos responsáveis pelo tratamento.\r\nAnalista       : Adilson Gonçalves Damasceno\r\n\r\nProgramador    : Adilson G. Damasceno\r\nData           : 02-07-2001 10:26\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F4621F402AC"
 
'' Objeto responsável pelo fornecimento de informações relativas às operações, bem
'' como a manipulação dos dados cadastrados

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
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

'Método genérico para a execução de outros métodos desta classe
Public Function Executar(ByRef pxmlExecucao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlExecucao.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        
        Case "LerTodosTiposOperacao"
            Executar = flLerTodosTiposOperacao(pxmlExecucao)
        
        Case "RejeitarOperacao"
            
            If Not pxmlExecucao.selectSingleNode("//DH_REGT_MESG_SPB") Is Nothing Then
                
                Call RejeitarOperacao(pxmlExecucao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                              pxmlExecucao.selectSingleNode("//DH_REGT_MESG_SPB").Text)
            
            ElseIf Not pxmlExecucao.selectSingleNode("//DH_ULTI_ATLZ") Is Nothing Then
                
                Call RejeitarOperacao(pxmlExecucao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, pxmlExecucao.selectSingleNode("//DH_ULTI_ATLZ").Text)
            
            End If
        
        Case "Reenviar"
            Call Reenviar(pxmlExecucao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                          pxmlExecucao.selectSingleNode("//DH_ULTI_ATLZ").Text)
                        
        Case "RejeitarTED"
            Call RejeitarTED(pxmlExecucao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                             pxmlExecucao.selectSingleNode("//DH_ULTI_ATLZ").Text)
                        
        Case "ObterComposicaoNet"
            Executar = flObterComposicaoNet(pxmlExecucao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        
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

'-------------------------------------------------------------------------------------------------
'Alterar o status de operações indisponiveis para consulta
'-------------------------------------------------------------------------------------------------

Public Function AlterarStatusOperIndisp(ByVal pvntSequenciaOperacao As Variant, _
                                        ByVal plngCodigoStatus As Long, _
                               Optional ByVal plngJustificativa As Long, _
                               Optional ByVal pintTipoAcao As Integer = 0, _
                               Optional ByVal pstrSituacaoRecebida As String = vbNullString) As Boolean

Dim strSQL                                  As String
Dim objHistoricoSituacaoOperacao            As A8LQS.clsHistSituacaoOperacao

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_OPER_ATIV " & vbCrLf & _
             " Set      CO_ULTI_SITU_PROC        =   " & plngCodigoStatus & vbCrLf

    If pintTipoAcao <> 0 Then
        strSQL = strSQL & _
            ",           TP_ACAO_OPER_ATIV_EXEC  =  " & pintTipoAcao & vbNewLine
    Else
        strSQL = strSQL & _
            ",           TP_ACAO_OPER_ATIV_EXEC  =  NULL " & vbNewLine
    End If

    If pstrSituacaoRecebida <> vbNullString Then
        strSQL = strSQL & _
            ",           CO_SITU_PROC_MESG_SPB_RECB   = '" & pstrSituacaoRecebida & "'             " & vbNewLine
    End If
             
    strSQL = strSQL & " Where    NU_SEQU_OPER_ATIV   =   " & pvntSequenciaOperacao

    fgExecuteSQL strSQL

    'Armazenar o Histórico da Situacao da Operação
    Set objHistoricoSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoSituacaoOperacao.Incluir pvntSequenciaOperacao, _
                                         plngCodigoStatus, _
                                         plngJustificativa, _
                                         pintTipoAcao
    Set objHistoricoSituacaoOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarStatusAcaoOperIndisp Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter detalhe de operações de corretoras
'-------------------------------------------------------------------------------------------------

Public Function ObterDetalheCorretoras(ByVal pvntSequenciaOperacao As Variant) As String

Dim strRetorno                              As String
Dim xmlRetorno                              As MSXML2.DOMDocument40
Dim objMensagem                             As A8LQS.clsMensagemSLCC
Dim xmlDominioAtributo                      As MSXML2.DOMDocument40
Dim strTipoMovimento                        As String

On Error GoTo ErrorHandler

    Set xmlRetorno = LerXMLOperacao(pvntSequenciaOperacao)
    
    If xmlRetorno.selectSingleNode("//REPE_DTLH_OPER") Is Nothing Then
        ObterDetalheCorretoras = vbNullString
    Else
        Set xmlDominioAtributo = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMensagem = CreateObject("A8LQS.clsMensagemSLCC")
        If Not xmlDominioAtributo.loadXML(objMensagem.ObterDominioAtributo("TP_MOVI_FINC")) Then
            fgErroLoadXML xmlDominioAtributo, "A8LQS", "clsOperacao", "ObterDetalheCorretoras"
        End If
        Set objMensagem = Nothing

        strTipoMovimento = xmlRetorno.selectSingleNode("//TP_MOVI_FINC").Text
        strTipoMovimento = strTipoMovimento & " - " & xmlDominioAtributo.selectSingleNode("//Repeat_DominioAtributo/Grupo_DominioAtributo[CO_DOMI='" & strTipoMovimento & "']/DE_DOMI").Text

        fgAppendNode xmlRetorno, "REPE_DTLH_OPER", "DE_MOVI", strTipoMovimento
        ObterDetalheCorretoras = xmlRetorno.selectSingleNode("//REPE_DTLH_OPER").xml
    End If

    Set xmlRetorno = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set xmlRetorno = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Alterar o tipo de compromisso da operação
'-------------------------------------------------------------------------------------------------

Public Function AlterarTipoCompromisso(ByVal plngSequenciaOperacao As Variant, _
                                       ByVal pstrDataUltimaAtualizacao As String, _
                                       ByVal pstrTipoCompromisso As String, _
                                       ByVal pstrTipoCompromissoRet As String, _
                                       ByVal pintStatusOperacao As String, _
                                       ByVal pstrAcaoAnterior As String _
                                       ) As Boolean

Dim strSQL                                  As String
Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao

'Pikachu - Release 2 (r2) , demanda 953
'Restringir o acesso à complementação de operações com data inferior a D-3.

Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim strDataMensagem                         As String

'--------------------------------------------------------------------------

On Error GoTo ErrorHandler

    'Pikachu - Release 2 (r2) , demanda 953
    'Restringir o acesso à complementação de operações com data inferior a D-3.
            
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                
    Set xmlOperacao = objOperacao.LerXMLOperacao(plngSequenciaOperacao)
    
    If Not xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
    
        Select Case Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        
            Case enumTipoOperacaoLQS.CompromissadaIda
            
                 If Not xmlOperacao.documentElement.selectSingleNode("DT_MESG") Is Nothing Then
                    If xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text <> fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) Then
                        '3100 - Data da Operação Inválida.
                        lngCodigoErroNegocio = 3116
                        GoTo ErrorHandler
                    End If
                End If
            
            Case enumTipoOperacaoLQS.CompromissadaVolta
                
                 If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                    If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text <> fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) Then
                        '3100 - Data da Operação Inválida.
                        lngCodigoErroNegocio = 3116
                        GoTo ErrorHandler
                    End If
                End If
    
        End Select
    End If
    
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
            
    'fim demanda 953 ---------------------------------------------------------------------------

    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoOperacao.Incluir plngSequenciaOperacao, _
                                 enumStatusOperacao.EmSer, _
                                 0, _
                                 enumTipoAcao.AlteracaoTipoCompromisso, _
                                 pstrAcaoAnterior
    Set objHistoricoOperacao = Nothing

    strSQL = " UPDATE   A8.TB_OPER_ATIV           " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          DH_ULTI_ATLZ = Sysdate,      " & vbCrLf & vbCrLf & _
             "          CO_ULTI_SITU_PROC       =    " & enumStatusOperacao.EmSer & vbCrLf
             
    If pstrTipoCompromisso <> vbNullString Then
       strSQL = strSQL & _
             " , TP_CPRO_OPER_ATIV              ='" & pstrTipoCompromisso & "' " & vbCrLf
    End If

    If pstrTipoCompromissoRet <> vbNullString Then
       strSQL = strSQL & _
             ", TP_CPRO_RETN_OPER_ATIV          ='" & pstrTipoCompromissoRet & "' " & vbCrLf
    End If

    strSQL = strSQL & _
                " Where    NU_SEQU_OPER_ATIV    =    " & plngSequenciaOperacao & vbCrLf & _
                " And      DH_ULTI_ATLZ         =    " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao)

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    'Verificar Confirmação/liberação Automático
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.GerenciarChamadaViaTela plngSequenciaOperacao, _
                                                enumFuncionalidade.Confirmacao

    Set objProcessoOperacao = Nothing

    AlterarTipoCompromisso = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objProcessoOperacao = Nothing
    Set objHistoricoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarTipoCompromisso Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Armazenar o numero de comando que retorna da Câmara
'-------------------------------------------------------------------------------------------------

Public Function AlterarComandoOperacaoRetn(ByVal plngSequenciaOperacao As Variant, _
                                           ByVal plngNroComandoOperacaoRetn As Long) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " UPDATE   A8.TB_OPER_ATIV           " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          DH_ULTI_ATLZ            =   Sysdate, " & vbCrLf & _
             "          NU_COMD_OPER_RETN       =    " & plngNroComandoOperacaoRetn & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV       =    " & plngSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    AlterarComandoOperacaoRetn = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarComandoOperacaoRetn Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Executar a chamada do estorno de lançamento em conta corrente
'-------------------------------------------------------------------------------------------------

Private Function flEstornarLancamento(ByVal pvntNumeroSequenciaOperacao As Variant) As Boolean

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil

On Error GoTo ErrorHandler

    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    Call objLancamentoCCContabil.Estornar(pvntNumeroSequenciaOperacao)
    Set objLancamentoCCContabil = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objLancamentoCCContabil = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEstornarLancamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Estornar a operação alterando o status para estorno solicitado
'-------------------------------------------------------------------------------------------------

Public Function Estornar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pintStatusOperacao As Integer, _
                         ByVal plngNumeroComando As Long, _
                         ByVal pstrDataUltimaAtualizacao As String _
                         ) As Boolean

Dim strSQL                                  As String
Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao

On Error GoTo ErrorHandler
    
    Call flEstornarLancamento(pvntSequenciaOperacao)
    
    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoOperacao.Incluir pvntSequenciaOperacao, _
                                 pintStatusOperacao, _
                                 , enumTipoAcao.EstornoSolicitado

    Set objHistoricoOperacao = Nothing
    
    strSQL = " UPDATE   A8.TB_OPER_ATIV              " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          DH_ULTI_ATLZ = Sysdate,      " & vbCrLf & vbCrLf & _
             "          TP_ACAO_OPER_ATIV_EXEC  =    " & enumTipoAcao.EstornoSolicitado & ", " & vbCrLf & _
             "          NU_COMD_ACAO_EXEC       =    " & IIf(plngNumeroComando = 0, " NULL ", plngNumeroComando) & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV       =    " & pvntSequenciaOperacao & vbCrLf & _
             " And      DH_ULTI_ATLZ            =    " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao)

     
    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    Estornar = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objHistoricoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Estornar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Gerenciar o cancelamento de operações geradas através da entrada manual
'-------------------------------------------------------------------------------------------------

Public Function CancelarEntradaManual(ByVal pvntSequenciaOperacao As Variant, _
                                      ByVal pintStatus As Integer, _
                                      ByVal pstrDataUltimaAtualizacao As String) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlOperacao = LerXMLOperacao(pvntSequenciaOperacao)
    
    If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacoesRodaDolar Then
            Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
            Call objProcessoOperacao.CancelarRodaDolarBMC(xmlOperacao)
            Set objProcessoOperacao = Nothing
            GoTo ExitSub
        End If
    End If
        
    Call AlterarStatusAcao(pvntSequenciaOperacao, _
                           enumStatusOperacao.CanceladaOrigem, _
                           pstrDataUltimaAtualizacao)

    If pintStatus = enumStatusOperacao.Concordancia Then
        
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.GerenciarCancelamento xmlOperacao, False
        Set objProcessoOperacao = Nothing
        
        If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
            If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DespesasBMC Then
                
                Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
                
                Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                Call xmlMensagem.loadXML(objConciliacao.ConsultaConciliacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
                Set objConciliacao = Nothing
                
                If xmlMensagem.xml <> vbNullString Then
                    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                    Call objMensagem.AlterarStatus(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                   xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                   enumStatusMensagem.AConciliar, _
                                                   vbNullString, _
                                                   , _
                                                   False, _
                                                   xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    Set objMensagem = Nothing
                End If
            
            End If
        End If
    
        Set xmlOperacao = Nothing
        
    End If

ExitSub:
    CancelarEntradaManual = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objProcessoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "CancelarEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o reenvio da mensagem SPB
'-------------------------------------------------------------------------------------------------

Public Function Reenviar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pstrDHUltimaAtualizacao As String) As Boolean

Dim objMensagem                             As A8LQS.clsMensagem
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim strMensagem                             As String

On Error GoTo ErrorHandler
    
    AlterarStatusAcao pvntSequenciaOperacao, enumStatusOperacao.Liberada, pstrDHUltimaAtualizacao
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    strMensagem = objMensagem.Reenviar(pvntSequenciaOperacao)
    Set objMensagem = Nothing
    
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.GerenciarReenvio strMensagem
    Set objProcessoOperacao = Nothing

    Reenviar = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Reenviar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar Rejeicao de Operações Pagamento de Despesas
Public Function RejeitarOperacao(ByVal pvntSequenciaOperacao As Variant, _
                                 ByVal pstrDHUltimaAtualizacao As String) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

    On Error GoTo ErrorHandler
    
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    
    Set xmlOperacao = LerXMLOperacao(pvntSequenciaOperacao, , , , , , , , , , , False)
        
    If xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoInterbancaria _
    Or xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoArbitragem _
    Or xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.IFInformaLiquidacaoInterbancaria _
    Or xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.IFCamaraConsultaContratosCambioMercadoInterbancario Then
        Call AlterarStatusAcao(pvntSequenciaOperacao, enumStatusOperacao.Rejeitada, pstrDHUltimaAtualizacao)
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    Else
        If Not xmlOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
            If xmlOperacao.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
                Call AlterarStatusAcao(pvntSequenciaOperacao, enumStatusOperacao.ManualEmSer, pstrDHUltimaAtualizacao)
            Else
                Call AlterarStatusAcao(pvntSequenciaOperacao, enumStatusOperacao.EmSer, pstrDHUltimaAtualizacao)
            End If
        Else
            Call AlterarStatusAcao(pvntSequenciaOperacao, enumStatusOperacao.EmSer, pstrDHUltimaAtualizacao)
        End If
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        
        If xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.DevolucaoRecolhimentoEstornoReembolsoCCR _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ContratacaoMercadoPrimario And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.EdicaoContratacaoMercadoPrimario _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConfirmacaoEdicaoContratacaoMercadoPrimario And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.AlteracaoContrato _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.EdicaoAlteracaoContrato And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConfirmacaoEdicaoAlteracaoContrato _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.LiquidacaoMercadoPrimario And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.BaixaValorLiquidar _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.RestabelecimentoBaixa And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.CancelamentoValorLiquidar _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.EdicaoCancelamentoValorLiquidar And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConfirmacaoEdicaoCancelamentoValorLiquidar _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.VinculacaoContratos And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.AnulacaoEvento _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.CorretoraRequisitaClausulasEspecificas And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.IFInformaClausulasEspecificas _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ManutencaoCadastroAgenciaCentralizadoraCambio And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.CredenciamentoDescredenciamentoDispostoRMCCI _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.IncorporacaoContratos And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.AceiteRejeicaoIncorporacaoContratos _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaContratosEmSer And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaEventosUmDia _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaDetalhamentoContratoInterbancario And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaEventosContratoMercadoPrimario _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaEventosContratoIntermediadoMercadoPrimario And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaHistoricoIncorporacoes _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaContratosIncorporacao And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaCadeiaIncorporacoesContrato _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaPosicaoCambioMoeda And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.AtualizaçãoInclusãoInstrucoesPagamento _
        And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.ConsultaInstrucoesPagamento And xmlOperacao.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemBUS.IFCamaraConsultaContratosCambioMercadoInterbancario Then
        
            Call fgAppendNode(xmlControleProcessoOperacao, vbNullString, "Repet_ProcessoOperacao", vbNullString)
            Call fgAppendNode(xmlControleProcessoOperacao, "Repet_ProcessoOperacao", "Grupo_ProcessoOperacao", vbNullString)
            Call fgAppendNode(xmlControleProcessoOperacao, "Grupo_ProcessoOperacao", "IN_ENVI_PREV_PJ", "1")
            Call fgAppendNode(xmlControleProcessoOperacao, "Grupo_ProcessoOperacao", "IN_ENVI_PREV_A6", "1")
            
            'Verifica layouts que devem Estornar Previsao Moeda Estrangeira PJ
            If xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoInterbancaria Or xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoArbitragem Then
                Call fgAppendNode(xmlControleProcessoOperacao, "Grupo_ProcessoOperacao", "IN_ENVI_PREV_PJ_ME", "1")
            End If
            
            'Estorar Previsao PJ Moeda Nacional
            Call objProcessoOperacao.EstornarPrevisaoBMACetip(xmlOperacao, xmlControleProcessoOperacao)
        
            'Verifica layouts que devem Estornar Realizado Moeda Estrangeira PJ
            If xmlOperacao.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.IFInformaLiquidacaoInterbancaria Then

                Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                
                'Estornar PJ Moeda Nacional
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            
                'Estornar PJ Moeda Estrangeira
                strMensagem = objGestaoCaixa.EstornarRealizadoME(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                
                Set objGestaoCaixa = Nothing

            End If
        
        End If
        
    End If
    
    Set objProcessoOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    RejeitarOperacao = True
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objProcessoOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "RejeitarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o rejeição de TEDs enviadas a Clientes
'-------------------------------------------------------------------------------------------------
Public Function RejeitarTED(ByVal pvntSequenciaOperacao As Variant, _
                            ByVal pstrDHUltimaAtualizacao As String) As Boolean

Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim intLayoutEntrada                        As Integer
Dim lngStatusFinal                          As Long

    On Error GoTo ErrorHandler
    
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = LerXMLOperacao(pvntSequenciaOperacao, , , , , , , , , , , False)
    
    intLayoutEntrada = 0
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        intLayoutEntrada = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
    End If
    
    lngStatusFinal = IIf(intLayoutEntrada = enumTipoMensagemBUS.EnvioTEDClientes, enumStatusOperacao.Cancelada, enumStatusOperacao.Rejeitada)
    
    Call AlterarStatusAcao(pvntSequenciaOperacao, lngStatusFinal, pstrDHUltimaAtualizacao)
    
    If Not xmlOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
        If xmlOperacao.selectSingleNode("//IN_ENTR_MANU").Text <> enumIndicadorSimNao.Sim Then
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        End If
    Else
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    End If
    
    Set objFluxoProcessamento = Nothing
    Set xmlOperacao = Nothing

    RejeitarTED = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objFluxoProcessamento = Nothing
    Set xmlOperacao = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "RejeitarTED Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Gerenciar a baixa de uma operação via contingência
'-------------------------------------------------------------------------------------------------

Public Function BaixarViaContingencia(ByVal pvntSequenciaOperacao As Variant, _
                                      ByVal pstrDHUltimaAtualizacao As String) As Boolean

On Error GoTo ErrorHandler
    
    AlterarStatusAcao pvntSequenciaOperacao, enumStatusOperacao.BaixadaViaContingencia, pstrDHUltimaAtualizacao
    
    'Falta disponibilizar lançamento em conta corrente quando houver
    
    BaixarViaContingencia = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "BaixarViaContingencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Gerenciar a liquidação de uma operação via contingência
'-------------------------------------------------------------------------------------------------

Public Function LiquidarViaContingencia(ByVal pvntSequenciaOperacao As Variant, _
                                        ByVal pstrDHUltimaAtualizacao As String) As Boolean

On Error GoTo ErrorHandler
    
    AlterarStatusAcao pvntSequenciaOperacao, enumStatusOperacao.LiquidadaViaContingencia, pstrDHUltimaAtualizacao
    
    'Falta Retornar Status LIQUIDADA para o sistema legado
    
    LiquidarViaContingencia = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarViaContingencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Cancelar a operação alterando o status para cancelamento solicitado
'-------------------------------------------------------------------------------------------------

Public Function Cancelar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pintStatusOperacao As Integer, _
                         ByVal plngNumeroComando As Long, _
                         ByVal pstrDataUltimaAtualizacao As String _
                         ) As Boolean

Dim strSQL                                  As String
Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlOperacao = LerXMLOperacao(pvntSequenciaOperacao)

    If Not xmlOperacao.selectSingleNode("//DT_MESG") Is Nothing Then
        If fgDtXML_To_Date(xmlOperacao.selectSingleNode("//DT_MESG").Text) <> fgDataHoraServidor(enumFormatoDataHora.Data) Then
            ' 3102 - Data da Operação Inválida para o Cancelamento.
            lngCodigoErroNegocio = 3102
            GoTo ErrorHandler
        End If
    End If
    
    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoOperacao.Incluir pvntSequenciaOperacao, _
                                 pintStatusOperacao, _
                                 , enumTipoAcao.CancelamentoSolicitado

    Set objHistoricoOperacao = Nothing
    
    strSQL = " UPDATE   A8.TB_OPER_ATIV              " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          DH_ULTI_ATLZ = Sysdate,      " & vbCrLf & vbCrLf & _
             "          TP_ACAO_OPER_ATIV_EXEC  =    " & enumTipoAcao.CancelamentoSolicitado & ", " & vbCrLf & _
             "          NU_COMD_ACAO_EXEC       =    " & IIf(plngNumeroComando = 0, " NULL ", plngNumeroComando) & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV       =    " & pvntSequenciaOperacao & vbCrLf & _
             " And      DH_ULTI_ATLZ            =    " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao)

     
    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    Cancelar = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objHistoricoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Cancelar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Alterar o status e/ou ação da operação
'-------------------------------------------------------------------------------------------------

Public Function AlterarStatusAcao(ByVal pvntSequenciaOperacao As Variant, _
                                  ByVal pintStatusOperacao As enumStatusOperacao, _
                                  ByVal pstrDHUltimaAtualizacao As String, _
                         Optional ByVal plngJustificativa As Long, _
                         Optional ByVal pintTipoAcao As Integer = 0, _
                         Optional ByVal pblnComparaDHUltimaAtualizacao As Boolean = True, _
                         Optional ByVal pstrNumCtrlMesgSPBOriginal As String = vbNullString, _
                         Optional ByVal pstrSituacaoRecebida As String = vbNullString, _
                         Optional ByVal pstrSequenciaConcilOperMesg As String = vbNullString) As Boolean

Dim strSQL                                  As String
Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_OPER_ATIV                                                             " & vbNewLine & _
             " Set      CO_ULTI_SITU_PROC       =   " & pintStatusOperacao & ",                     " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "',                        " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "',             " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate                                         " & vbNewLine

    If pintTipoAcao <> 0 Then
        strSQL = strSQL & _
            ",           TP_ACAO_OPER_ATIV_EXEC  =  " & pintTipoAcao & "                            " & vbNewLine
    Else
        strSQL = strSQL & _
            ",           TP_ACAO_OPER_ATIV_EXEC  =  NULL                                            " & vbNewLine
    End If
             
    If pstrNumCtrlMesgSPBOriginal <> vbNullString Then
        strSQL = strSQL & _
            ",           NU_CTRL_MESG_SPB_ORIG   = '" & pstrNumCtrlMesgSPBOriginal & "'             " & vbNewLine
    End If
             
    If pstrSituacaoRecebida <> vbNullString Then
        strSQL = strSQL & _
            ",           CO_SITU_PROC_MESG_SPB_RECB   = '" & pstrSituacaoRecebida & "'             " & vbNewLine
    End If
             
    If pstrSequenciaConcilOperMesg <> vbNullString Then
        strSQL = strSQL & _
            ",           NU_SEQU_CNCL_OPER_ATIV_MESG   = " & pstrSequenciaConcilOperMesg & "             " & vbNewLine
    End If
             
    strSQL = strSQL & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao
    
    If pblnComparaDHUltimaAtualizacao Then
        strSQL = strSQL & vbNewLine & _
                 " And      DH_ULTI_ATLZ            =   " & fgDtHrXML_To_Oracle(pstrDHUltimaAtualizacao)
    End If
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    'Armazenar o Histórico da Situacao da Operação
    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoOperacao.Incluir pvntSequenciaOperacao, _
                                 pintStatusOperacao, _
                                 plngJustificativa, _
                                 pintTipoAcao
    Set objHistoricoOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objHistoricoOperacao = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarStatusAcao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Alterar qualquer campo da tabela TB_OPER_ATIV
'-------------------------------------------------------------------------------------------------

Public Function Alterar(ByVal pvntSequenciaOperacao As Variant, _
               Optional ByVal pintStatusOperacao As Integer = 0, _
               Optional ByVal pintTipoAcao As Integer = 0, _
               Optional ByVal pstrNumCtrlMesgSPBOriginal As String = vbNullString, _
               Optional ByVal pstrSituacaoRecebida As String = vbNullString, _
               Optional ByVal pstrSequenciaConcilOperMesg As String = vbNullString, _
               Optional ByVal pstrCanalOperacaoInterna As String = vbNullString) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_OPER_ATIV                                                             " & vbNewLine & _
             " Set      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "',                        " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "',             " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate                                         " & vbNewLine

    If pintStatusOperacao <> 0 Then
        strSQL = strSQL & _
            ",           CO_ULTI_SITU_PROC  =  " & pintStatusOperacao & "                            " & vbNewLine
    End If
             
    If pintTipoAcao <> 0 Then
        strSQL = strSQL & _
            ",           TP_ACAO_OPER_ATIV_EXEC  =  " & pintTipoAcao & "                            " & vbNewLine
    Else
        strSQL = strSQL & _
            ",           TP_ACAO_OPER_ATIV_EXEC  =  NULL                                            " & vbNewLine
    End If
             
    If pstrNumCtrlMesgSPBOriginal <> vbNullString Then
        strSQL = strSQL & _
            ",           NU_CTRL_MESG_SPB_ORIG   = '" & pstrNumCtrlMesgSPBOriginal & "'             " & vbNewLine
    End If
             
    If pstrSituacaoRecebida <> vbNullString Then
        strSQL = strSQL & _
            ",           CO_SITU_PROC_MESG_SPB_RECB   = '" & pstrSituacaoRecebida & "'             " & vbNewLine
    End If
             
    If pstrSequenciaConcilOperMesg <> vbNullString Then
        strSQL = strSQL & _
            ",           NU_SEQU_CNCL_OPER_ATIV_MESG   = " & pstrSequenciaConcilOperMesg & "             " & vbNewLine
    End If
             
    If pstrCanalOperacaoInterna <> vbNullString Then
        strSQL = strSQL & _
            ",           CO_CNAL_OPER_INTE   = '" & pstrCanalOperacaoInterna & "'             " & vbNewLine
    End If
             
    strSQL = strSQL & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Alterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Armazenar o numero de controle SPB original que retorna da Câmara
'-------------------------------------------------------------------------------------------------

Public Function GravarControleSPBOriginal(ByVal pvntSequenciaOperacao As Variant, _
                                          ByVal pstrControleRDCOriginal As String) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_OPER_ATIV " & vbNewLine & _
             " Set      NU_CTRL_MESG_SPB_ORIG   =   '" & pstrControleRDCOriginal & "', " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GravarControleSPBOriginal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'RATS -942
Public Function GravarControleSPBOriginalCTP(ByVal pvntSequenciaOperacao As Variant, _
                                             ByVal pstrControleSPBOriginal As String, _
                                    Optional ByVal pstrNumeroOperacaoCTP As String, _
                                    Optional ByVal pstrNumAtivMerc As String) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_OPER_ATIV " & vbNewLine & _
             " Set      NU_CTRL_MESG_SPB_ORIG   =   '" & pstrControleSPBOriginal & "', " & vbNewLine & _
             "          NU_COMD_OPER_RETN       =   '" & pstrNumeroOperacaoCTP & "', " & vbNewLine & _
             "          NU_ATIV_MERC            =   '" & pstrNumAtivMerc & "', " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GravarControleSPBOriginal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar a disponibilização ou não da operação para consulta
'-------------------------------------------------------------------------------------------------

Public Function OperacaoDisponivelConsulta(ByRef pvntSequenciaOperacao As Variant, _
                                           ByRef plngIndicadorSimNao As enumIndicadorSimNao) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_OPER_ATIV " & vbCrLf & _
             " Set      IN_DISP_CONS        =   " & plngIndicadorSimNao & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV   =   " & pvntSequenciaOperacao

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "OperacaoDisponivelConsulta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Alterar o tipo de ação da mensagem SPB
'-------------------------------------------------------------------------------------------------

Public Function AlterarAcaoMensagemSPB(ByRef pvntSequenciaOperacao As Variant, _
                                       ByRef plngAcaoMensagemSPB As Long) As Boolean

Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler
    
    strSQL = " Update   A8.TB_OPER_ATIV " & vbCrLf & _
             " Set      TP_ACAO_OPER_ATIV_EXEC  =   " & IIf(plngAcaoMensagemSPB = 0, _
                                                            "NULL", plngAcaoMensagemSPB) & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao
    
    fgExecuteSQL strSQL
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAcaoMensagemSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Gerar dinamicamente a alteração de uma operação previamente cadastrada
'-------------------------------------------------------------------------------------------------

Public Function ComplementarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                     ByVal pvntSequenciaOperacao As Variant, _
                                     ByVal plngTipoOperacao As Long) As Boolean

Dim strSQL                                  As String
Dim strColunasValores                       As String
Dim strTipoLiquidacao                       As String
Dim rsOperacao                              As ADODB.Recordset
Dim objField                                As ADODB.Field
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlTipoLiquidacao                       As MSXML2.DOMDocument40
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objHistoricoSituacaoOperacao            As A8LQS.clsHistSituacaoOperacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objTipoLiquidacao                       As A6A7A8.clsTipoLiquidacao
Dim strDataInclusao                         As String

    On Error GoTo ErrorHandler

    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EspecificacaoOperacaoBMA Then
        ObterComandoEspecificacao xmlRemessa
    End If

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")

    'Se o local de liquidação BMA é sempre multilateral
    If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
        
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.TransferenciaBMA Or _
            xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EspecificacaoOperacaoBMA Or _
            xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.IntermediacaoOperInternas Then
            fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", 0
        Else
            fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral
        End If
        
    ElseIf xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then

        If Not xmlRemessa.documentElement.selectSingleNode("QT_TITU_CETIP") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "QT_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("QT_TITU_CETIP").Text
        End If

        If Not xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC_CETIP") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "NU_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC_CETIP").Text
        End If

        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.DespesasCETIP Then
                fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Bruta
            End If
        End If
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.DespesasCETIP Then
            fgAppendNode xmlRemessa, "MESG", "TP_CNPT", 1
        End If

        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoesCETIP Then
            If xmlRemessa.documentElement.selectSingleNode("IN_TITL").Text = "S" Then
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
            Else
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito
            End If
        End If

        If Not xmlRemessa.documentElement.selectSingleNode("SG_SIST_CETIP") Is Nothing Then
            objTipoOperacao.ObterTipoOperacaoPorMsgBUS xmlRemessa, plngTipoOperacao, "", ""
        End If

    End If

    Set rsOperacao = flObterPropriedades

    'Alterado para que no BACEN não de erro, pois o A8 envia com 0
    If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
        If Val("0" & xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text) = 0 Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text = vbNullString
        End If
    End If

    For Each objField In rsOperacao.fields
        Select Case objField.Name
            Case "NU_SEQU_OPER_ATIV"
                'Não fazer nada
            '    strColunasValores = strColunasValores & objField.Name & " = " & "1,"
            Case "TP_OPER"
                strColunasValores = strColunasValores & objField.Name & " = " & plngTipoOperacao & ","
            Case "CO_USUA_ULTI_ATLZ"
                strColunasValores = strColunasValores & objField.Name & " = '" & fgUsuarioRede & "',"
            Case "CO_ETCA_TRAB_ULTI_ATLZ"
                strColunasValores = strColunasValores & objField.Name & " = '" & fgEstacaoTrabalhoUsuario & "',"
            Case "DH_ULTI_ATLZ"
                strDataInclusao = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
                strColunasValores = strColunasValores & objField.Name & " = " & fgDtHrXML_To_Oracle(strDataInclusao) & ", "
            Case "IN_DISP_CONS" 'Disponível para consulta
                strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorSimNao.Sim & ","
            Case "CO_ULTI_SITU_PROC" 'Status da Operação
                strColunasValores = strColunasValores & objField.Name & " = " & enumStatusOperacao.EmSer & ","
            Case "SG_SIST"
                strColunasValores = strColunasValores & objField.Name & " = '" & xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "',"
            Case "TP_ACAO_OPER_ATIV_EXEC"
                strColunasValores = strColunasValores
            Case "IN_ENTR_SAID_RECU_FINC"
                Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
                    Case enumTipoMensagemBUS.Redesconto, enumTipoMensagemBUS.ConversaoRedesconto
                        strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorEntradaSaida.ENTRADA & ","
                    Case enumTipoMensagemBUS.PagamentoRedesconto
                        strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorEntradaSaida.Saida & ","
                    Case Else
                        If xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                            strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorEntradaSaida.Saida & ","
                        Else
                            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                                strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorEntradaSaida.ENTRADA & ","
                            Else
                                strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorEntradaSaida.Saida & ","
                            End If
                        End If
                End Select
            Case "IN_OPER_DEBT_CRED"
                Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
                    Case enumTipoMensagemBUS.Redesconto, enumTipoMensagemBUS.ConversaoRedesconto
                        strColunasValores = strColunasValores & objField.Name & " = " & enumTipoDebitoCredito.Debito & ","
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito
                    Case enumTipoMensagemBUS.PagamentoRedesconto
                        strColunasValores = strColunasValores & objField.Name & " = " & enumTipoDebitoCredito.Credito & ","
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
                    Case enumTipoMensagemBUS.DespesasSelic
                        strColunasValores = strColunasValores & objField.Name & " = " & enumTipoEntradaSaida.Saida & ","
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
                    Case Else
                        If xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                            strColunasValores = strColunasValores & objField.Name & " = " & "Null,"
                        Else
                            strColunasValores = strColunasValores & objField.Name & " = " & xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text & ","
                        End If
                End Select
            Case "HO_ENVI_MESG_SPB"
                If Not xmlRemessa.documentElement.selectSingleNode("HO_AGND") Is Nothing Then
                    If Trim(xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text) = "" Or _
                        Format(xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text, "0000") = "0000" Then
                        strColunasValores = strColunasValores & objField.Name & " = " & "Null, "
                    Else
                        strColunasValores = strColunasValores & objField.Name & " = " & fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(enumFormatoDataHora.Data), "YYYYMMDD") & fgCompletaString(xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text, "0", 4, True) & "00") & ","
                    End If
                Else
                    strColunasValores = strColunasValores & objField.Name & " = " & "Null, "
                End If
            Case "IN_ENTR_MANU"
                If Not xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                    strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorSimNao.Sim & ", "
                Else
                    strColunasValores = strColunasValores & objField.Name & " = " & enumIndicadorSimNao.Nao & ", "
                End If
            Case "DT_OPER_ATIV"
                If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                    strColunasValores = strColunasValores & objField.Name & " = " & fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text) & ", "
                Else
                    strColunasValores = strColunasValores & objField.Name & " = " & fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))) & " , "
                End If
            Case "CD_LOTE"
                If Not xmlRemessa.documentElement.selectSingleNode("ID_LOTE") Is Nothing Then
                   If Trim(xmlRemessa.documentElement.selectSingleNode("ID_LOTE").Text) <> "" Then
                        strColunasValores = strColunasValores & objField.Name & " = " & xmlRemessa.documentElement.selectSingleNode("ID_LOTE").Text & ", "
                   Else
                        strColunasValores = strColunasValores & objField.Name & " = " & "Null,"
                   End If
                Else
                    strColunasValores = strColunasValores & objField.Name & " = " & "Null,"
                End If
                
            Case Else
                Set objNode = xmlRemessa.documentElement.selectSingleNode("//" & objField.Name)
                If objNode Is Nothing Then
                    If objField.Name = "VA_OPER_ATIV" Then
                        strColunasValores = strColunasValores & objField.Name & " = " & "0, "
                    ElseIf objField.Name = "DE_ATIV_MERC" Then
                        Set objNode = xmlRemessa.documentElement.selectSingleNode("//CO_SUB_TIPO_ATIV")
                        If objNode Is Nothing Then
                            strColunasValores = strColunasValores & objField.Name & " = " & "Null, "
                        Else
                            strColunasValores = strColunasValores & objField.Name & " = " & "'" & objNode.Text & "', "
                        End If
                    Else
                        strColunasValores = strColunasValores & objField.Name & " = " & "Null, "
                    End If
                Else
                    Select Case objField.Type
                        Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                            If CLng("0" & objNode.Text) = 0 Then
                                strColunasValores = strColunasValores & objField.Name & " = NULL, "
                            Else
                                strColunasValores = strColunasValores & objField.Name & " = " & fgDtXML_To_Oracle(objNode.Text) & ", "
                            End If
                        Case Is = ADODB.adNumeric, ADODB.adDouble
                            If objField.DefinedSize > 0 Then
                                strColunasValores = strColunasValores & objField.Name & " = " & fgVlrToDBServer(objNode.Text) & ", "
                            Else
                                strColunasValores = strColunasValores & objField.Name & " = " & objNode.Text & ", "
                            End If
                        Case ADODB.adChar, ADODB.adVarChar
                            strColunasValores = strColunasValores & objField.Name & " = " & "'" & objNode.Text & "', "
                        Case Else
                            Err.Raise vbObjectError + 100, TypeName(Me), "Tipo de dado não previsto no Insert da Operacao"
                    End Select
                End If
        End Select
    Next

    strColunasValores = Left$(Trim$(strColunasValores), Len(Trim$(strColunasValores)) - 1) & " "

    strSQL = " Update   A8.TB_OPER_ATIV  Set " & vbCrLf & strColunasValores & vbCrLf & _
             " Where    CO_OPER_ATIV    =   '" & xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text & "'" & vbCrLf & _
             " And     (TP_ACAO_OPER_ATIV_EXEC != " & enumTipoAcao.RejeicaoPorDuplicidade & vbCrLf & _
             " Or       TP_ACAO_OPER_ATIV_EXEC IS NULL)"

    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Then
        If Val(xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text) > 0 Then
            strSQL = strSQL & " And    SG_SIST      =   '" & xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "'" & vbCrLf & _
                              " And    NU_ATIV_MERC =   '" & xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text & "'" & vbCrLf & _
                              " And    DT_VENC_ATIV =    " & fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text)
        Else
            strSQL = strSQL & " And    SG_SIST      =   '" & xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "'" & vbCrLf & _
                              " And    NU_ATIV_MERC =   '" & xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text & "'" & vbCrLf
        End If
    End If

    'Para o Layout 50 e sistema LQC (Se houver o Id do Lote), realizar a inclusão do lote
    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.OperacoesCorretoras _
        And Trim(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) = "LQC" Then
        If Not xmlRemessa.documentElement.selectSingleNode("ID_LOTE") Is Nothing Then
            If Trim(xmlRemessa.documentElement.selectSingleNode("ID_LOTE").Text) <> "" Then
                IncluirLote xmlRemessa
            End If
        End If
    End If

    fgExecuteSQL strSQL

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir pvntSequenciaOperacao, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    Set objOperacaoMensagemInterna = Nothing

    'Armazenar o Histórico da Situacao da Operação
    Set objHistoricoSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoSituacaoOperacao.Incluir pvntSequenciaOperacao, _
                                         enumStatusOperacao.EmSer, _
                                         0
    Set objHistoricoSituacaoOperacao = Nothing

    'Armazenar a Sequencia para futura utilização
    'Verificar o nome da TAG principal da Mensagem que virá do BUS
    If xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", pvntSequenciaOperacao
    Else
        xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = pvntSequenciaOperacao
    End If
    'Incluir o Tipo de Operacao
    If xmlRemessa.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "TP_OPER", plngTipoOperacao
    Else
        xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = plngTipoOperacao
    End If
    'Incluir o Status da Operação
    If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", enumStatusOperacao.EmSer
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.EmSer
    End If
    'Tipo ação da mensagem
    If xmlRemessa.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "TP_ACAO_OPER_ATIV_EXEC", "0"
    Else
        xmlRemessa.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = "0"
    End If
    'Tipo ação da mensagem
    If xmlRemessa.documentElement.selectSingleNode("DH_ULTI_ATLZ") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "DH_ULTI_ATLZ", strDataInclusao
    Else
        xmlRemessa.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text = strDataInclusao
    End If

    'Incluir o Codigo da Mensagem SPB
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML objTipoOperacao.Ler(plngTipoOperacao)
    If Trim(xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text) = vbNullString Then
        If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            Set objTipoLiquidacao = CreateObject("A6A7A8.clsTipoLiquidacao")
            strTipoLiquidacao = objTipoLiquidacao.Ler(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
            If strTipoLiquidacao <> vbNullString Then
                Set xmlTipoLiquidacao = CreateObject("MSXML2.DOMDocument.4.0")
                xmlTipoLiquidacao.loadXML strTipoLiquidacao
                fgAppendNode xmlRemessa, "MESG", "CO_MESG", xmlTipoLiquidacao.documentElement.selectSingleNode("CO_MESG_SPB_LIQU").Text
                Set xmlTipoLiquidacao = Nothing
            Else
                fgAppendNode xmlRemessa, "MESG", "CO_MESG", vbNullString
            End If
            Set objTipoLiquidacao = Nothing
        Else
            fgAppendNode xmlRemessa, "MESG", "CO_MESG", vbNullString
        End If
    Else
        fgAppendNode xmlRemessa, "MESG", "CO_MESG", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
    End If
    fgAppendNode xmlRemessa, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacaoMensagemInterna = Nothing
    Set objHistoricoSituacaoOperacao = Nothing
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ComplementarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar a confirmalção de uma operação
'-------------------------------------------------------------------------------------------------

Public Function Confirmar(ByVal pvntSequenciaOperacao As Variant, _
                          ByVal pintStatusOperacao As enumStatusOperacao, _
                          ByVal pstrDHUltimaAtualizacao As String, _
                          ByVal pvntProtocolo As Variant) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objAlcadasLimites                       As A6A7A8.clsPerfil
Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40

Dim lngStatusOperacao                       As enumStatusOperacao
Dim lngFuncFuncao                           As enumFuncionalidade
Dim blnCompararDataAtualizacao              As Boolean

    On Error GoTo ErrorHandler

    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")

    Set xmlRemessa = LerXMLOperacao(pvntSequenciaOperacao)

    'Se o Status da Operação for MANUAL EM SER, verificar Limites e Alçadas
    If pintStatusOperacao = enumStatusOperacao.ManualEmSer Or _
       xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
       If Not xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
            
            Set objAlcadasLimites = CreateObject("A6A7A8.clsPerfil")
            Call objAlcadasLimites.DigitarAlcada(fgUsuarioRede, _
                                                 xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                                 pvntProtocolo)
            Set objAlcadasLimites = Nothing
            
            Call AlterarNumeroProtocoloLG(pvntSequenciaOperacao, _
                                          pvntProtocolo)
            
            Call fgAppendNode(xmlRemessa, xmlRemessa.documentElement.nodeName, "NU_PRTC_MESG_LG", pvntProtocolo)
        End If
        
    End If
    
    'Lancamento C/C - Operacao Manual
    If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.LancamentoContaCorrenteBG Then
        
        Set objAlcadasLimites = CreateObject("A6A7A8.clsPerfil")
        Call objAlcadasLimites.DigitarAlcada(fgUsuarioRede, _
                                             xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                             pvntProtocolo)
        Set objAlcadasLimites = Nothing
            
        Call AlterarNumeroProtocoloLG(pvntSequenciaOperacao, _
                                      pvntProtocolo)
            
        Call fgAppendNode(xmlRemessa, xmlRemessa.documentElement.nodeName, "NU_PRTC_MESG_LG", pvntProtocolo)
        
    End If
        
    'Pagamento de Despesas
    If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas Then
        Select Case Val(xmlRemessa.selectSingleNode("//TP_OPER").Text)
            Case enumTipoOperacaoLQS.EnvioSTR0006PagDespesas, _
                 enumTipoOperacaoLQS.EnvioSTR0008PagDespesas, _
                 enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta, _
                 enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib, _
                 enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta, _
                 enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib, _
                 enumTipoOperacaoLQS.EnvioPagDespesasBoleto, _
                 enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta, _
                 enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib, _
                 enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente, _
                 enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta, _
                 enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib, _
                 enumTipoOperacaoLQS.EnvioPagDespesasTributos, _
                 enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta, _
                 enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib
                If Not xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                    Set objAlcadasLimites = CreateObject("A6A7A8.clsPerfil")
                    Call objAlcadasLimites.DigitarAlcada(fgUsuarioRede, _
                                                 xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, _
                                                 pvntProtocolo)
                    Set objAlcadasLimites = Nothing

                    Call AlterarNumeroProtocoloLG(pvntSequenciaOperacao, _
                                          pvntProtocolo)

                    Call fgAppendNode(xmlRemessa, xmlRemessa.documentElement.nodeName, "NU_PRTC_MESG_LG", pvntProtocolo)

               End If


            Case Else

        End Select

    End If
    
    lngFuncFuncao = enumFuncionalidade.Confirmacao
    blnCompararDataAtualizacao = True
    
    If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
        Select Case Val(xmlRemessa.selectSingleNode("//TP_OPER").Text)
            Case enumTipoOperacaoLQS.RegistroLiquidacaoEventoCBLC
                lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
            Case enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao
                If Not xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
                    If Val(xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ReativacaoSolicitada Then
                        lngStatusOperacao = enumStatusOperacao.ConcordanciaReativacao
                    Else
                        lngStatusOperacao = enumStatusOperacao.ConcordanciaBalcao
                    End If
                End If
            Case enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica
                lngStatusOperacao = enumStatusOperacao.AConciliarRegistro
            Case enumTipoOperacaoLQS.RegistroOperacoesBMC
                lngStatusOperacao = AConciliarAceite
            
            'KIDA - BMF
            Case enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF
                lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
            
            Case enumTipoOperacaoLQS.NegociacaoOperacaoCCR
                lngStatusOperacao = enumStatusOperacao.Concordancia
                
                'CCR - KIDA
                Dim objProcessoCCR As A8LQS.clsProcessoMensagemCCR
                'Set objProcessoCCR = CreateObject("A8LQS.clsProcessoMensagemCCR")
                
                'If Not xmlRemessa.selectSingleNode("//TpManut") Is Nothing Then
                '    If xmlRemessa.documentElement.selectSingleNode("//TpManut").Text = "I" Then
                '        If Not objProcessoCCR.ConciliarOperacao(CLng(xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text)) Then
                '            lngStatusOperacao = enumStatusOperacao.ManualEmSer
                '            'lngCodigoErroNegocio = 4433
                '            'GoTo ErrorHandler
                '        End If
                '    End If
                'End If
                'Set objProcessoCCR = Nothing
                

            Case Else
                
                lngStatusOperacao = enumStatusOperacao.Concordancia
                
        End Select
    Else
        lngStatusOperacao = enumStatusOperacao.Concordancia
    End If
    
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        
        If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
            lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
            lngFuncFuncao = enumFuncionalidade.Liberacao
            blnCompararDataAtualizacao = True
        
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas Then

                lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
                Select Case Val("0" & xmlRemessa.selectSingleNode("//TP_OPER").Text)
                    Case enumTipoOperacaoLQS.EnvioPagDespesasBoleto, _
                         enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente, _
                         enumTipoOperacaoLQS.EnvioPagDespesasTributos, _
                         enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta, _
                         enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib, _
                         enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib, _
                         enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta, _
                         enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta, _
                         enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib
                        lngFuncFuncao = enumFuncionalidade.Confirmacao
                    Case Else
                        lngFuncFuncao = enumFuncionalidade.Liberacao
                End Select
                
                blnCompararDataAtualizacao = True
        
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.LancamentoContaCorrenteBG Then
            
            lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
            lngFuncFuncao = enumFuncionalidade.Confirmacao
            blnCompararDataAtualizacao = True
        
        End If
        
    End If
    
    Call AlterarStatusAcao(pvntSequenciaOperacao, _
                           lngStatusOperacao, _
                           pstrDHUltimaAtualizacao, _
                           , , _
                           blnCompararDataAtualizacao)

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.GerenciarChamadaViaTela pvntSequenciaOperacao, lngFuncFuncao
    Set objProcessoOperacao = Nothing

    If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
        If xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DespesasBMC Then
            
            Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Call xmlMensagem.loadXML(objConciliacao.ConsultaConciliacao(xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
            Set objConciliacao = Nothing
            
            If xmlMensagem.xml <> vbNullString Then
                Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Call objMensagem.AlterarStatus(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                               xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                               enumStatusMensagem.Concordancia, _
                                               vbNullString, _
                                               , _
                                               False, _
                                               xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                Set objMensagem = Nothing
            End If
        
        End If
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set xmlRemessa = Nothing
    Set objAlcadasLimites = Nothing
    Set objProcessoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Confirmar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Public Function IncluirLote(ByVal xmlRemessa As MSXML2.DOMDocument40)
Dim rsLote                                  As ADODB.Recordset
Dim intIdLote                               As Integer
Dim strSiglaSistema                         As String
Dim intTipoDebitoCredito                    As Integer
Dim dblValorTotalLote                       As Double
Dim intQuantidadeOperLote                   As Integer
Dim strSQL                                  As String
Dim strDtOperAtiv                           As String


    On Error GoTo ErrorHandler

        intIdLote = fgVlrXml_To_Decimal(xmlRemessa.selectSingleNode("//ID_LOTE").Text)
        strSiglaSistema = xmlRemessa.selectSingleNode("//SG_SIST_ORIG").Text
        intTipoDebitoCredito = fgVlrXml_To_Decimal(xmlRemessa.selectSingleNode("//TP_DEB_CRED_LOTE").Text)
        dblValorTotalLote = fgVlrXml_To_Decimal(xmlRemessa.selectSingleNode("//VA_TOT_LOTE").Text)
        intQuantidadeOperLote = fgVlrXml_To_Decimal(xmlRemessa.selectSingleNode("//QT_OPER_LOTE").Text)
        strDtOperAtiv = fgDtXML_To_Oracle(xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text)
        
        Set rsLote = CreateObject("ADODB.Recordset")
        
        'Verifica se já há um lote cadastrado
        Set rsLote = LerLote(intIdLote, strSiglaSistema, strDtOperAtiv)
        If rsLote.EOF Then
            'Incluir o Lote na Base
            strSQL = " Insert Into  A8.TB_LOTE_OPER_ATIV (CD_LOTE, SG_SIST, DT_OPER_ATIV, TP_DEBT_CRED_LOTE, VL_TOTL_LOTE, QT_TOTL_LOTE, IN_LIQU_LOTE, DH_ULTI_ATLZ)" & vbCrLf & _
            " Values (" & vbCrLf & intIdLote & vbCrLf & _
                      " ,'" & strSiglaSistema & "'" & vbCrLf & _
                      "," & strDtOperAtiv & vbCrLf & _
                      "," & intTipoDebitoCredito & vbCrLf & _
                      "," & fgVlrToDBServer(dblValorTotalLote) & vbCrLf & _
                      "," & intQuantidadeOperLote & vbCrLf & _
                      ", 2" & vbCrLf & _
                      ", SYSDATE" & vbCrLf & ")"
            
            fgExecuteSQL strSQL
        Else
            If Trim(xmlRemessa.selectSingleNode("//TP_SOLI").Text) = enumTipoSolicitacao.Complementacao _
                And rsLote!IN_LIQU_LOTE = enumSituacaoLote.ALiberar Then
                    strSQL = "Update A8.TB_LOTE_OPER_ATIV set TP_DEBT_CRED_LOTE = " & intTipoDebitoCredito & "," & vbCrLf & _
                             "VL_TOTL_LOTE  = " & fgVlrToDBServer(dblValorTotalLote) & "," & vbCrLf & _
                             "QT_TOTL_LOTE = " & intQuantidadeOperLote & "," & vbCrLf & _
                             "DH_ULTI_ATLZ = SYSDATE" & vbCrLf & _
                             "WHERE CD_LOTE = " & intIdLote & vbCrLf & _
                             "AND SG_SIST = '" & strSiglaSistema & "'" & vbCrLf & _
                             "AND DT_OPER_ATIV = " & strDtOperAtiv
                    
                    fgExecuteSQL strSQL
            End If
        End If
        
        Set rsLote = Nothing
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    

End Function


'-------------------------------------------------------------------------------------------------
'Gerar dinamicamente a inclusão de uma operação
'-------------------------------------------------------------------------------------------------

Public Function IncluirOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                ByRef plngStatusOperacao As Long) As Boolean

Dim strSQL                                  As String
Dim strColunas                              As String
Dim strValores                              As String
Dim vntSequencialOperacao                   As Variant
Dim lngTipoOperacao                         As Long
Dim strCodigoMensagem                       As String
Dim strOperacao                             As String
Dim rsOperacao                              As ADODB.Recordset
Dim objField                                As ADODB.Field
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objTipoLiquidacao                       As A6A7A8.clsTipoLiquidacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objHistoricoSituacaoOperacao            As A8LQS.clsHistSituacaoOperacao
Dim xmlTipoLiquidacao                       As MSXML2.DOMDocument40
Dim strIdentificadorOperacao                As String
Dim strMensagemRetornoLegado                As String
Dim strTipoLiquidacao                       As String
Dim strDataInclusao                         As String
Dim blnPrimeiraEspecificacao                As Boolean

Dim intEntradaManual                        As enumIndicadorSimNao

On Error GoTo ErrorHandler

    If plngStatusOperacao <> enumStatusOperacao.Inativa Then
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.LiquidacaoFisicaOperacaoBMA Then
            flIncluirLiquidacaoFisicaOper xmlRemessa, enumStatusOperacao.AConciliar
            Exit Function
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EspecificacaoOperacaoBMA Then
            ObterComandoEspecificacao xmlRemessa
            
            If xmlRemessa.documentElement.selectSingleNode("//TP_ESFC").Text <> enumTipoEspecificacao.Cancelamento Then
                'Demanda A822
                'Alteração no SLCC para permitir o tratamento correto do Fluxo da BMA0006,
                'fechado com a Câmara
                blnPrimeiraEspecificacao = Not ExisteOperacaoEspecificacao(xmlRemessa)
            
                If Not blnPrimeiraEspecificacao Then
                    If xmlRemessa.documentElement.selectSingleNode("//TP_NEGO_BMA").Text = enumTipoNegociacaoBMA.Compromissada Then
                        fgAppendNode xmlRemessa, "MESG", "CO_MESG", "BMA0005"
                    Else
                        fgAppendNode xmlRemessa, "MESG", "CO_MESG", "BMA0004"
                    End If
                End If
                '--------------------------------------------------------------------------
            Else
                blnPrimeiraEspecificacao = True
            End If
        End If
    End If

    'Se o local de liquidação BMA é sempre multilateral
    If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.TransferenciaBMA Or _
            xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EspecificacaoOperacaoBMA Or _
            xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.IntermediacaoOperInternas Then
            fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", 0
        Else
            fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral
        End If
    
    ElseIf xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        
        If Not xmlRemessa.documentElement.selectSingleNode("QT_TITU_CETIP") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "QT_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("QT_TITU_CETIP").Text
        End If

        If Not xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC_CETIP") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "NU_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC_CETIP").Text
        End If

        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.DespesasCETIP Then
                fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Bruta
            End If
        Else
            '----------------- RATS- 945
            If Not xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                
                If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
                    If Trim(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text) = "" Then
                        xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.SemModalidade
                    End If
                Else
                    fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.SemModalidade
                End If
                
                If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                    If Trim(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = "" Then
                        xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = "0"
                    End If
                Else
                    fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", "0"
                    
                End If
                
            End If
            '----------------- RATS- 945
        
        End If
        
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.DespesasCETIP Then
            fgAppendNode xmlRemessa, "MESG", "TP_CNPT", 1
        End If

        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoesCETIP Then
            If xmlRemessa.documentElement.selectSingleNode("IN_TITL").Text = "S" Then
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
            Else
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito
            End If
        End If

    ElseIf xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMC Then
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.TransferenciasBMC Then
            fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", 0
        End If
    
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
        intEntradaManual = Val(xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU").Text)
    Else
        intEntradaManual = enumIndicadorSimNao.Nao
    End If

    'Irá segurar a inclusão, pois o Codigo da Operção e a Sigla do Sistema não podem duplicar e o indice
    ' não pode ser único pois temos a compromissada de ida que vem com estes iguais modificando somente
    ' o identificado do título
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    'objTipoOperacao.BloquearTipoOperacaoParaInclusao

    'Obter a Operação correspondente a Mensagem do BUS
    If xmlRemessa.selectSingleNode("//NET_ENTRADA_MANUAL") Is Nothing Then
        objTipoOperacao.ObterTipoOperacaoPorMsgBUS xmlRemessa, lngTipoOperacao, strCodigoMensagem, strMensagemRetornoLegado
    Else
        lngTipoOperacao = Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
    End If

    If intEntradaManual = enumIndicadorSimNao.Nao Then
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemBUS.CompromissadaIda And _
           xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemBUS.Redesconto And _
           xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemBUS.PagamentoRedesconto And _
           xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemBUS.OperacaoCompromissadaCETIP Then
            'Verificar se existe o Identificador da Operação para o Sistema Origem, caso exista é erro!
            strOperacao = LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    
            If strOperacao <> vbNullString Then
            
                fgAppendNode xmlRemessa, "MESG", "TP_OPER", lngTipoOperacao
                fgAppendNode xmlRemessa, "MESG", "TP_MESG_RETN_INTE", strMensagemRetornoLegado
            
                '3037 - Identificador da Operação já existe.
                lngCodigoErroNegocio = 3037
                GoTo ErrorHandler
            End If
        Else
            'Verificar se existe o Identificador da Operação para o Sistema Origem e o Identificador do lastro
            strOperacao = LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                   , , , , , , , , , , _
                                   xmlRemessa.documentElement.selectSingleNode("CO_IDEF_LAST").Text)

            If strOperacao <> vbNullString Then
                '3037 - Identificador da Operação já existe.
                lngCodigoErroNegocio = 3037
                GoTo ErrorHandler
            End If
        End If
    End If

    'Alterado para que no BACEN não de erro, pois o A8 envia com 0
    If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
        If Val("0" & xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text) = 0 Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text = vbNullString
        End If
    End If

    If strCodigoMensagem = vbNullString Then
        If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            Set objTipoLiquidacao = CreateObject("A6A7A8.clsTipoLiquidacao")
            strTipoLiquidacao = objTipoLiquidacao.Ler(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
            If strTipoLiquidacao <> vbNullString Then
                Set xmlTipoLiquidacao = CreateObject("MSXML2.DOMDocument.4.0")
                xmlTipoLiquidacao.loadXML strTipoLiquidacao
                strCodigoMensagem = xmlTipoLiquidacao.documentElement.selectSingleNode("CO_MESG_SPB_LIQU").Text
                Set xmlTipoLiquidacao = Nothing
            End If
            Set objTipoLiquidacao = Nothing
        End If
    End If
    
    Set objTipoOperacao = Nothing

    Set rsOperacao = flObterPropriedades
    strValores = ""
    For Each objField In rsOperacao.fields
        strColunas = strColunas & objField.Name & ", "
        Select Case objField.Name
            Case "NU_SEQU_OPER_ATIV"
                vntSequencialOperacao = fgExecuteSequence("A8.SQ_A8_NU_SEQU_OPER_ATIV")
                strValores = strValores & vntSequencialOperacao & ","
            Case "TP_OPER"
                strValores = strValores & lngTipoOperacao & ","
            Case "CO_USUA_ULTI_ATLZ"
                strValores = strValores & "'" & fgUsuarioRede & "',"
            Case "CO_ETCA_TRAB_ULTI_ATLZ"
                strValores = strValores & "'" & fgEstacaoTrabalhoUsuario & "',"
            Case "DH_ULTI_ATLZ"
                strDataInclusao = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
                strValores = strValores & fgDtHrXML_To_Oracle(strDataInclusao) & ", "
            Case "IN_DISP_CONS" 'Disponível para consulta
                strValores = strValores & enumIndicadorSimNao.Sim & ","
            Case "CO_ULTI_SITU_PROC" 'Status da Operação
                If plngStatusOperacao <> enumStatusOperacao.Inativa Then
                    If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.EventosJurosCETIP And _
                           fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text) > fgDataHoraServidor(enumFormatoDataHora.Data) Then
                            If Not xmlRemessa.documentElement.selectSingleNode("TP_CNPT") Is Nothing And _
                               Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
                                If Val(xmlRemessa.documentElement.selectSingleNode("TP_CNPT").Text) <> enumTipoContraparte.Cliente1 Or _
                                   Val(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text) <> enumTipoLiquidacao.SemModalidade Then
                                    plngStatusOperacao = enumStatusOperacao.LiquidacaoFutura
                                End If
                            End If
                        End If
                    End If
                End If
                strValores = strValores & plngStatusOperacao & ","
            Case "SG_SIST"
                strValores = strValores & "'" & xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "',"
            Case "IN_ENTR_SAID_RECU_FINC"
                Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
                    Case enumTipoMensagemBUS.Redesconto, enumTipoMensagemBUS.ConversaoRedesconto
                        strValores = strValores & enumIndicadorEntradaSaida.ENTRADA & ","
                    Case enumTipoMensagemBUS.PagamentoRedesconto
                        strValores = strValores & enumIndicadorEntradaSaida.Saida & ","
                    Case Else
                        If xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                            If Not xmlRemessa.documentElement.selectSingleNode("TP_OPER_CAMB") Is Nothing Then
                                Select Case xmlRemessa.documentElement.selectSingleNode("TP_OPER_CAMB").Text
                                    Case "V"
                                        strValores = strValores & enumIndicadorEntradaSaida.ENTRADA & ","
                                    Case "C"
                                        strValores = strValores & enumIndicadorEntradaSaida.Saida & ","
                                    Case Else
                                        strValores = strValores & "0,"
                                End Select
                            Else
                                strValores = strValores & "0,"
                            End If
                        Else
                            'RATS- 945
                            If Not IsNumeric(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) Then
                                strValores = strValores & enumIndicadorEntradaSaida.ENTRADA & ","
                            Else
                                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                                    strValores = strValores & enumIndicadorEntradaSaida.ENTRADA & ","
                                Else
                                    strValores = strValores & enumIndicadorEntradaSaida.Saida & ","
                                End If
                            End If
                        End If
                End Select
            Case "IN_OPER_DEBT_CRED"
                Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
                    Case enumTipoMensagemBUS.Redesconto, enumTipoMensagemBUS.ConversaoRedesconto
                        strValores = strValores & enumTipoDebitoCredito.Debito & ", "
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito
                    Case enumTipoMensagemBUS.PagamentoRedesconto
                        strValores = strValores & enumTipoDebitoCredito.Credito & ", "
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
                    Case enumTipoMensagemBUS.DespesasSelic
                        strValores = strValores & enumTipoEntradaSaida.Saida & ", "
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
                    Case Else
                        If xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                            strValores = strValores & " 0, "
                        Else
                            strValores = strValores & xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text & ", "
                        End If
                End Select
            Case "HO_ENVI_MESG_SPB"
                If Not xmlRemessa.documentElement.selectSingleNode("HO_AGND") Is Nothing Then
                    If Trim(xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text) = "" Or _
                        Format(xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text, "0000") = "0000" Then
                        strValores = strValores & "Null, "
                    Else
                        strValores = strValores & fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(enumFormatoDataHora.Data), "YYYYMMDD") & fgCompletaString(xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text, "0", 4, True) & "00") & ", "
                    End If
                Else
                    strValores = strValores & "Null, "
                End If
            Case "IN_ENTR_MANU"
                If Not xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                    strValores = strValores & enumIndicadorSimNao.Sim & ", "
                Else
                    strValores = strValores & enumIndicadorSimNao.Nao & ", "
                End If
            Case "DT_OPER_ATIV"
                If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                    strValores = strValores & fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text) & ", "
                Else
                    strValores = strValores & fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))) & " , "
                End If
            Case "VA_OPER_ATIV"
                If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                    strValores = strValores & fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text) & ", "
                Else
                    If Not xmlRemessa.documentElement.selectSingleNode("VA_MOED_NACIO") Is Nothing Then
                        strValores = strValores & fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("VA_MOED_NACIO").Text) & ", "
                    Else
                        strValores = strValores & "0, "
                    End If
                End If
            Case "VA_MOED_ESTR"
                If Not xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTRG") Is Nothing Then
                    strValores = strValores & fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTRG").Text) & ", "
                Else
                    strValores = strValores & "Null, "
                End If
            Case "NO_CLIE"
                If Not xmlRemessa.documentElement.selectSingleNode("//NM_CLIE") Is Nothing Then
                    strValores = strValores & "'" & xmlRemessa.documentElement.selectSingleNode("//NM_CLIE").Text & "', "
                Else
                    strValores = strValores & "Null, "
                End If
            Case "CD_MOED_ISO"
                If Not xmlRemessa.documentElement.selectSingleNode("CO_MOED_ISO") Is Nothing Then
                    strValores = strValores & "'" & xmlRemessa.documentElement.selectSingleNode("CO_MOED_ISO").Text & "', "
                Else
                    strValores = strValores & "Null, "
                End If
             Case "NR_PERC_TAXA_CAMB"
                If Not xmlRemessa.documentElement.selectSingleNode("VA_TAXA_CAMB") Is Nothing Then
                    strValores = strValores & fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("VA_TAXA_CAMB").Text) & ","
                Else
                    strValores = strValores & "Null, "
                End If
             Case "NU_COMD_OPER"
                If Not xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER") Is Nothing Then
                    strValores = strValores & "'" & xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text & "', "
                Else
                    If Not xmlRemessa.documentElement.selectSingleNode("CO_REG_OPER_CAMB") Is Nothing Then
                        If Trim(xmlRemessa.documentElement.selectSingleNode("CO_REG_OPER_CAMB").Text) <> "" Then
                            strValores = strValores & xmlRemessa.documentElement.selectSingleNode("CO_REG_OPER_CAMB").Text & ", "
                        Else
                            strValores = strValores & "Null,"
                        End If
                    Else
                        strValores = strValores & "Null,"
                    End If
                End If
             Case "NR_OPER_CAMB_2"
                If Not xmlRemessa.documentElement.selectSingleNode("CO_REG_OPER_CAMB2") Is Nothing Then
                    If Trim(xmlRemessa.documentElement.selectSingleNode("CO_REG_OPER_CAMB2").Text) <> "" Then
                        strValores = strValores & xmlRemessa.documentElement.selectSingleNode("CO_REG_OPER_CAMB2").Text & ", "
                    Else
                        strValores = strValores & "Null,"
                    End If
                Else
                    strValores = strValores & "Null,"
                End If
             Case "CD_ASSO_CAMB"
                If Not xmlRemessa.documentElement.selectSingleNode("ChACAM") Is Nothing Then
                    If Trim(xmlRemessa.documentElement.selectSingleNode("ChACAM").Text) <> "" Then
                        strValores = strValores & "'" & xmlRemessa.documentElement.selectSingleNode("ChACAM").Text & "', "
                    Else
                        strValores = strValores & "Null,"
                    End If
                Else
                    strValores = strValores & "Null,"
                End If
              Case "CD_LOTE"
                If Not xmlRemessa.documentElement.selectSingleNode("ID_LOTE") Is Nothing Then
                   If Trim(xmlRemessa.documentElement.selectSingleNode("ID_LOTE").Text) <> "" Then
                      strValores = strValores & fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("ID_LOTE").Text) & ","
                   Else
                      strValores = strValores & "Null,"
                   End If
                Else
                    strValores = strValores & "Null,"
                End If
            Case Else
                'Para a inclusão de entrada manual, será incluso o código da operação neste momento...
                If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing And _
                    objField.Name = "CO_OPER_ATIV" Then
                    If Trim$(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text) = vbNullString Then
                        strIdentificadorOperacao = "A8" & vntSequencialOperacao
                        strValores = strValores & "'" & strIdentificadorOperacao & "',"
                        xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text = strIdentificadorOperacao
                    Else
                        strIdentificadorOperacao = xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text
                        strValores = strValores & "'" & strIdentificadorOperacao & "',"
                    End If
                ElseIf objField.Name = "NR_CNPJ_CPF" And (Trim(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) = "GPC" Or Trim(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) = "R2") Then  'operações de COMEX
                    Select Case strCodigoMensagem
                        Case "CAM0021"
                            If Not xmlRemessa.documentElement.selectSingleNode("//CNPJ_CPF_CLI") Is Nothing Then
                               strValores = strValores & Trim(xmlRemessa.documentElement.selectSingleNode("//CNPJ_CPF_CLI").Text) & ","
                            Else
                               strValores = strValores & "Null,"
                            End If
                        Case "CAM0023"
                        'RATS 1178 - IM02716447  - GPC - ERRO CAM23 MANUAL
                            If Not xmlRemessa.documentElement.selectSingleNode("//CNPJ_CPF_CLI") Is Nothing Then
                               strValores = strValores & Trim(xmlRemessa.documentElement.selectSingleNode("//CNPJ_CPF_CLI").Text) & ","
                            ElseIf Not xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CORRT") Is Nothing Then
                                If xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CORRT").Text <> vbNullString Then
                                    strValores = strValores & Trim(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CORRT").Text) & ","
                                Else
                                    strValores = strValores & "Null,"
                                End If
                            ElseIf Not xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_BASE_CORRT") Is Nothing Then
                               strValores = strValores & Trim(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_BASE_CORRT").Text) & ","
                            Else
                               strValores = strValores & "Null,"
                            End If
                        Case "CAM0026", "CAM0032", "CAM0036"
                            If Not xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_BASE_CORRT") Is Nothing Then
                               strValores = strValores & Trim(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_BASE_CORRT").Text) & ","
                            Else
                               strValores = strValores & "Null,"
                            End If
                        Case Else
                            strValores = strValores & "Null,"
                    End Select
                ElseIf (objField.Name = "NO_CNPT" Or objField.Name = "DT_LIQU_OPER_ATIV" Or objField.Name = "PE_TAXA_NEGO") And (xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoInterbancaria Or xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacaoArbitragem Or xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.IFInformaLiquidacaoInterbancaria) Then
                    If objField.Name = "NO_CNPT" Then
                        If Not xmlRemessa.documentElement.selectSingleNode("//NM_CLIE_MOED_ESTR") Is Nothing Then
                           strValores = strValores & "'" & Trim(xmlRemessa.documentElement.selectSingleNode("//NM_CLIE_MOED_ESTR").Text) & "',"
                        Else
                           strValores = strValores & "Null,"
                        End If
                    ElseIf objField.Name = "DT_LIQU_OPER_ATIV" Then
                        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.IFInformaLiquidacaoInterbancaria Then
                            If Not xmlRemessa.documentElement.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
                               strValores = strValores & fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("//DT_OPER_ATIV").Text) & ","
                            Else
                               strValores = strValores & "Null,"
                            End If
                        Else
                            If Not xmlRemessa.documentElement.selectSingleNode("//DT_LIQU_OPER") Is Nothing Then
                               strValores = strValores & fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("//DT_LIQU_OPER").Text) & ","
                            Else
                               strValores = strValores & "Null,"
                            End If
                        End If
                    ElseIf objField.Name = "PE_TAXA_NEGO" Then
                        If Not xmlRemessa.documentElement.selectSingleNode("//VA_TAXA_CAMB") Is Nothing And xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemBUS.RegistroOperacaoArbitragem Then
                           strValores = strValores & fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("//VA_TAXA_CAMB").Text) & ","
                        Else
                           strValores = strValores & "Null,"
                        End If
                    End If
                Else
                    Set objNode = xmlRemessa.documentElement.selectSingleNode("//" & objField.Name)
                    If objNode Is Nothing Then
                        If objField.Name = "DE_ATIV_MERC" Then
                            Set objNode = xmlRemessa.documentElement.selectSingleNode("//CO_SUB_TIPO_ATIV")
                            If objNode Is Nothing Then
                                strValores = strValores & "Null, "
                            Else
                                strValores = strValores & "'" & objNode.Text & "', "
                            End If
                        Else
                            strValores = strValores & "Null, "
                        End If
                    Else
                        Select Case objField.Type
                            Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                                If Val(objNode.Text) = 0 Then
                                    strValores = strValores & " NULL ,"
                                Else
                                    strValores = strValores & fgDtXML_To_Oracle(objNode.Text) & ", "
                                End If
                            Case Is = ADODB.adNumeric, ADODB.adDouble
                                If objField.DefinedSize > 0 Then
                                    strValores = strValores & fgVlrToDBServer(objNode.Text) & ", "
                                Else
                                    strValores = strValores & objNode.Text & ", "
                                End If
                            Case ADODB.adChar, ADODB.adVarChar
                                strValores = strValores & "'" & objNode.Text & "', "
                            Case Else
                                Err.Raise vbObjectError + 100, TypeName(Me), "Tipo de dado não previsto no Insert da Operacao"
                        End Select
                    End If
                End If
        End Select
    Next

    strColunas = Left$(Trim$(strColunas), Len(Trim$(strColunas)) - 1) & " "
    strValores = Left$(Trim$(strValores), Len(Trim$(strValores)) - 1) & " "
    
    'Para o Layout 50 e sistema LQC (Se houver o Id do Lote), realizar a inclusão do lote
    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.OperacoesCorretoras _
        And Trim(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) = "LQC" Then
        If Not xmlRemessa.documentElement.selectSingleNode("ID_LOTE") Is Nothing Then
            If Trim(xmlRemessa.documentElement.selectSingleNode("ID_LOTE").Text) <> "" Then
                IncluirLote xmlRemessa
            End If
        End If
    End If
    
    strSQL = " Insert Into  A8.TB_OPER_ATIV (" & vbCrLf & strColunas & ")" & vbCrLf & " Values (" & vbCrLf & strValores & ")"
    fgExecuteSQL strSQL

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir vntSequencialOperacao, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    Set objOperacaoMensagemInterna = Nothing

    'Armazenar o Histórico da Situacao da Operação
    Set objHistoricoSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoSituacaoOperacao.Incluir vntSequencialOperacao, _
                                         plngStatusOperacao
    Set objHistoricoSituacaoOperacao = Nothing

    'Armazenar a Sequencia para futura utilização
    'Verificar o nome da TAG principal da Mensagem que virá do BUS
    If xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", vntSequencialOperacao
    Else
        xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vntSequencialOperacao
    End If
    'Incluir o Tipo de Operacao
    fgAppendNode xmlRemessa, "MESG", "TP_OPER", lngTipoOperacao
    'Incluir o Status da Operação
    fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", plngStatusOperacao
    
    If xmlRemessa.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
        'Incluir o Codigo da Mensagem SPB
        fgAppendNode xmlRemessa, "MESG", "CO_MESG", strCodigoMensagem
    End If
    
    'Tipo ação da mensagem
    fgAppendNode xmlRemessa, "MESG", "TP_ACAO_OPER_ATIV_EXEC", "0"
    'Tipo de mensagem de Retorno para o Legado
    fgAppendNode xmlRemessa, "MESG", "TP_MESG_RETN_INTE", strMensagemRetornoLegado
    'Data da inclusão para chamada da conciliacao
    fgAppendNode xmlRemessa, "MESG", "DH_ULTI_ATLZ", strDataInclusao

    'Demanda A822
    'Alteração no SLCC para permitir o tratamento correto do Fluxo da BMA0006,
    'fechado com a Câmara
    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EspecificacaoOperacaoBMA Then
        If Not blnPrimeiraEspecificacao Then
            If xmlRemessa.documentElement.selectSingleNode("//TP_NEGO_BMA").Text = enumTipoNegociacaoBMA.Compromissada Then
                xmlRemessa.selectSingleNode("//CO_MESG").Text = "BMA0005"
            Else
                xmlRemessa.selectSingleNode("//CO_MESG").Text = "BMA0004"
            End If
        End If
    End If
    '---------------------------------------------------------------------------------

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objTipoOperacao = Nothing
    Set objTipoLiquidacao = Nothing
    Set objOperacaoMensagemInterna = Nothing
    Set objHistoricoSituacaoOperacao = Nothing
    Set xmlTipoLiquidacao = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Function LerLote(ByVal pintIdLote As Integer, ByVal pstrSiglaSistema As String, ByVal pstrDtOperAtiv As String) As ADODB.Recordset
Dim strSQL                                  As String
Dim rsLote                                  As ADODB.Recordset
    On Error GoTo ErrorHandler

        Set rsLote = CreateObject("ADODB.Recordset")
        strSQL = "Select * " & vbCrLf & _
                 " From  A8.TB_LOTE_OPER_ATIV " & vbCrLf & _
                 " Where CD_LOTE = " & pintIdLote & vbCrLf & _
                 " And   SG_SIST         =   '" & pstrSiglaSistema & "'" & vbCrLf & _
                 " And   DT_OPER_ATIV    =   " & pstrDtOperAtiv
                 
        Set rsLote = fgQuerySQL(strSQL)
            
        Set LerLote = rsLote
        
        Set rsLote = Nothing
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function
'-------------------------------------------------------------------------------------------------
'Ler uma operação conforme o filtro
'-------------------------------------------------------------------------------------------------
Public Function Ler(ByVal pvntSequenciaOperacao As Variant) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Select NU_SEQU_OPER_ATIV, TP_OPER, CO_LOCA_LIQU, TP_LIQU_OPER_ATIV, CO_EMPR, " & vbCrLf & _
             "       CO_USUA_CADR_OPER, HO_ENVI_MESG_SPB, CO_OPER_ATIV, NU_COMD_OPER, " & vbCrLf & _
             "       NU_COMD_OPER_RETN, DT_OPER_ATIV, DT_OPER_ATIV_RETN, CO_VEIC_LEGA, " & vbCrLf & _
             "       SG_SIST, CO_CNTA_CUTD_SELIC_VEIC_LEGA, CO_CNPJ_CNPT, NO_CNPT, " & vbCrLf & _
             "       CO_CNTA_CUTD_SELIC_CNPT, IN_OPER_DEBT_CRED, NU_ATIV_MERC, DE_ATIV_MERC, " & vbCrLf & _
             "       PU_ATIV_MERC, QT_ATIV_MERC, IN_ENTR_SAID_RECU_FINC, DT_VENC_ATIV, " & vbCrLf & _
             "       VA_OPER_ATIV, VA_OPER_ATIV_REAJ, DT_LIQU_OPER_ATIV, TP_CPRO_OPER_ATIV, " & vbCrLf & _
             "       TP_CPRO_RETN_OPER_ATIV, IN_DISP_CONS, IN_ENVI_PREV_SIST_PJ, IN_ENVI_RELZ_SIST_PJ, " & vbCrLf & _
             "       IN_ENVI_PREV_SIST_A6, IN_ENVI_RELZ_SIST_A6, CO_ULTI_SITU_PROC, TP_ACAO_OPER_ATIV_EXEC, " & vbCrLf & _
             "       NU_COMD_ACAO_EXEC, IN_ENTR_MANU, NU_SEQU_CNCL_OPER_ATIV_MESG, NU_CTRL_MESG_SPB_ORIG, " & vbCrLf & _
             "       PE_TAXA_NEGO, CO_TITL_CUTD, CO_USUA_ULTI_ATLZ, CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "       DH_ULTI_ATLZ, CO_OPER_CETIP, CO_ISPB_BANC_LIQU_CNPT, TP_IF_CRED_DEBT, " & vbCrLf & _
             "       CO_AGEN_COTR, NU_CC_COTR, PZ_DIAS_RETN_OPER_ATIV, VA_OPER_ATIV_RETN, " & vbCrLf & _
             "       TP_CNPT, CO_CNPT_CAMR, NU_PRTC_OPER_LG, CO_IDEF_LAST, CO_GRUP_LANC_FINC, " & vbCrLf & _
             "       CO_PARP_CAMR , TP_PGTO_LDL , TP_CNAL_VEND" & vbCrLf & _
             " From  A8.TB_OPER_ATIV " & vbCrLf & _
             " Where NU_SEQU_OPER_ATIV  =   " & pvntSequenciaOperacao

    Ler = fgQueryXMLLer("Operacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'-------------------------------------------------------------------------------------------------
'Ler uma operação ou n operações conforme o filtro especificado
'-------------------------------------------------------------------------------------------------

'ATENÇÃO: O Codigo da Operação é unico por Sistema
Public Function LerTodos(Optional ByVal pstrCodigoOperacao As String = "", _
                         Optional ByVal pstrSiglaSistema As String = "", _
                         Optional ByVal pstrComandoOperacao As String = "", _
                         Optional ByVal pstrComandoOperacaoRet As String = "", _
                         Optional ByVal pdatDataOperacao As Date = gstrDataVazia, _
                         Optional ByVal plngTipoOperacao As Long = 0, _
                         Optional ByVal pstrNumCtrlRDCOriginal As String = "", _
                         Optional ByVal pstrIdentificadorTitulo As String = "", _
                         Optional ByVal plngCodigoEmpresa As Long = 0, _
                         Optional ByVal pstrCodigoVeiculoLegal As String = "", _
                         Optional ByVal pdatDataVencimento As Date = gstrDataVazia, _
                         Optional ByVal pvntCNPJContraparte As Variant = 0, _
                         Optional ByVal pstrIdentificadorLastro As String = "", _
                         Optional ByVal pblnForUpdate As Boolean = True) As String

Dim strSQL                                  As String
Dim blnWhere                                As Boolean

    On Error GoTo ErrorHandler

    strSQL = "Select NU_SEQU_OPER_ATIV, TP_OPER, CO_LOCA_LIQU, TP_LIQU_OPER_ATIV, CO_EMPR, " & vbCrLf & _
             "       CO_USUA_CADR_OPER, HO_ENVI_MESG_SPB, CO_OPER_ATIV, NU_COMD_OPER, " & vbCrLf & _
             "       NU_COMD_OPER_RETN, DT_OPER_ATIV, DT_OPER_ATIV_RETN, CO_VEIC_LEGA, " & vbCrLf & _
             "       SG_SIST, CO_CNTA_CUTD_SELIC_VEIC_LEGA, CO_CNPJ_CNPT, NO_CNPT, " & vbCrLf & _
             "       CO_CNTA_CUTD_SELIC_CNPT, IN_OPER_DEBT_CRED, NU_ATIV_MERC, DE_ATIV_MERC, " & vbCrLf & _
             "       PU_ATIV_MERC, QT_ATIV_MERC, IN_ENTR_SAID_RECU_FINC, DT_VENC_ATIV, " & vbCrLf & _
             "       VA_OPER_ATIV, VA_OPER_ATIV_REAJ, DT_LIQU_OPER_ATIV, TP_CPRO_OPER_ATIV, " & vbCrLf & _
             "       TP_CPRO_RETN_OPER_ATIV, IN_DISP_CONS, IN_ENVI_PREV_SIST_PJ, IN_ENVI_RELZ_SIST_PJ, " & vbCrLf & _
             "       IN_ENVI_PREV_SIST_A6, IN_ENVI_RELZ_SIST_A6, CO_ULTI_SITU_PROC, TP_ACAO_OPER_ATIV_EXEC, " & vbCrLf & _
             "       NU_COMD_ACAO_EXEC, IN_ENTR_MANU, NU_SEQU_CNCL_OPER_ATIV_MESG, NU_CTRL_MESG_SPB_ORIG, " & vbCrLf & _
             "       PE_TAXA_NEGO, CO_TITL_CUTD, CO_USUA_ULTI_ATLZ, CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "       DH_ULTI_ATLZ, CO_OPER_CETIP, CO_ISPB_BANC_LIQU_CNPT, TP_IF_CRED_DEBT, " & vbCrLf & _
             "       CO_AGEN_COTR, NU_CC_COTR, PZ_DIAS_RETN_OPER_ATIV, VA_OPER_ATIV_RETN, " & vbCrLf & _
             "       TP_CNPT, CO_CNPT_CAMR, NU_PRTC_OPER_LG, CO_IDEF_LAST, CO_GRUP_LANC_FINC, " & vbCrLf & _
             "       CO_PARP_CAMR , TP_PGTO_LDL , TP_CNAL_VEND" & vbCrLf & _
             " From  A8.TB_OPER_ATIV " & vbCrLf

    blnWhere = False
    If Trim(pstrCodigoOperacao) <> vbNullString Then
        strSQL = strSQL & " Where   CO_OPER_ATIV    =   '" & pstrCodigoOperacao & "'" & vbCrLf & _
                          " And     SG_SIST         =   '" & pstrSiglaSistema & "' "
        blnWhere = True
    End If

    If pstrComandoOperacao <> "" Then
        If blnWhere Then
            strSQL = strSQL & " And     NU_COMD_OPER   =   '" & pstrComandoOperacao & "'"
        Else
            strSQL = strSQL & " Where   NU_COMD_OPER   =   '" & pstrComandoOperacao & "'"
            blnWhere = True
        End If
    End If

    If pstrComandoOperacaoRet <> "" Then
        If blnWhere Then
            strSQL = strSQL & " And     NU_COMD_OPER_RETN   =   '" & pstrComandoOperacaoRet & "'"
        Else
            strSQL = strSQL & " Where   NU_COMD_OPER_RETN   =   '" & pstrComandoOperacaoRet & "'"
            blnWhere = True
        End If
    End If

    If pdatDataOperacao <> gstrDataVazia Then
        If blnWhere Then
            strSQL = strSQL & " And     DT_OPER_ATIV        =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdatDataOperacao))
        Else
            strSQL = strSQL & " Where   DT_OPER_ATIV        =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdatDataOperacao))
            blnWhere = True
        End If
    End If

    If plngTipoOperacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     TP_OPER             =   " & plngTipoOperacao
        Else
            strSQL = strSQL & " Where   TP_OPER             =   " & plngTipoOperacao
            blnWhere = True
        End If
    End If

    If pstrNumCtrlRDCOriginal <> vbNullString Then
        If blnWhere Then
            strSQL = strSQL & " And     NU_CTRL_MESG_SPB_ORIG   =   '" & pstrNumCtrlRDCOriginal & "'"
        Else
            strSQL = strSQL & " Where   NU_CTRL_MESG_SPB_ORIG   =   '" & pstrNumCtrlRDCOriginal & "'"
            blnWhere = True
        End If
    End If

    If pstrIdentificadorTitulo <> vbNullString Then
        If blnWhere Then
            strSQL = strSQL & " And     NU_ATIV_MERC   =   '" & pstrIdentificadorTitulo & "'"
        Else
            strSQL = strSQL & " Where   NU_ATIV_MERC   =   '" & pstrIdentificadorTitulo & "'"
            blnWhere = True
        End If
    End If

    If plngCodigoEmpresa > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     CO_EMPR   =   " & plngCodigoEmpresa
        Else
            strSQL = strSQL & " Where   CO_EMPR   =   " & plngCodigoEmpresa
            blnWhere = True
        End If
    End If
    
    If Trim(pstrCodigoVeiculoLegal) <> vbNullString Then
        If blnWhere Then
            strSQL = strSQL & " And     CO_VEIC_LEGA   =   '" & pstrCodigoVeiculoLegal & "'"
        Else
            strSQL = strSQL & " Where   CO_VEIC_LEGA   =   '" & pstrCodigoVeiculoLegal & "'"
            blnWhere = True
        End If
    End If

    If pdatDataVencimento <> gstrDataVazia Then
        If blnWhere Then
            strSQL = strSQL & " And     DT_VENC_ATIV        =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdatDataVencimento))
        Else
            strSQL = strSQL & " Where   DT_VENC_ATIV        =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdatDataVencimento))
            blnWhere = True
        End If
    End If

    If pvntCNPJContraparte > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     CO_CNPJ_CNPT    =   " & pvntCNPJContraparte
        Else
            strSQL = strSQL & " Where   CO_CNPJ_CNPT    =   " & pvntCNPJContraparte
            blnWhere = True
        End If
    End If

    If pstrIdentificadorLastro <> "" Then
        If blnWhere Then
            strSQL = strSQL & " And     CO_IDEF_LAST   =   '" & pstrIdentificadorLastro & "'"
        Else
            strSQL = strSQL & " Where   CO_IDEF_LAST   =   '" & pstrIdentificadorLastro & "'"
            blnWhere = True
        End If
    End If

    strSQL = strSQL & " Order By NU_SEQU_OPER_ATIV "

    If pblnForUpdate And Not fgEstacaoDesenvolvimento Then
        strSQL = strSQL & " For Update"
    End If

    LerTodos = fgQueryXMLLerTodos("Operacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Ler o xml de uma operação conforme o filtro especificado
'-------------------------------------------------------------------------------------------------

'ATENÇÃO: O Codigo da Operação é unico por Sistema
'Operações compromissada e redesconto repete o codigo da operação e o que desempata é o identificador do lastro
Public Function LerXMLOperacao(Optional ByVal pvntSequenciaOperacao As Variant = 0, _
                               Optional ByVal pstrCodigoOperacao As String = "", _
                               Optional ByVal pstrSiglaSistema As String = "", _
                               Optional ByVal pstrNumeroOperacaoSelic As String = "", _
                               Optional ByVal pstrIdentificadorTitulo As String = "", _
                               Optional ByVal pstrComandoOperacaoRet As String = "", _
                               Optional ByVal pblnSolicitacaoCancelamento As Boolean = False, _
                               Optional ByVal pstrIdentificadorLastro As String = "", _
                               Optional ByRef pobjRs As ADODB.Recordset = Nothing, _
                               Optional ByVal pvntNrIdefNegoBMC As Variant = 0, _
                               Optional ByVal pstrChACAM As String = "", _
                               Optional ByVal pblnForUpdate As Boolean = True) As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim rsOperacao                              As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsOperacao = CreateObject("ADODB.Recordset")
    Set LerXMLOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    strSQL = " Select   A.TP_OPER,                     " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          A.DT_OPER_ATIV,                " & vbCrLf & _
             "          A.IN_ENTR_MANU,                " & vbCrLf & _
             "          A.HO_ENVI_MESG_SPB,            " & vbCrLf & _
             "          A.TP_ACAO_OPER_ATIV_EXEC,      " & vbCrLf & _
             "          A.NU_COMD_ACAO_EXEC,           " & vbCrLf & _
             "          A.DH_ULTI_ATLZ,                " & vbCrLf & _
             "          A.TP_CPRO_OPER_ATIV,           " & vbCrLf & _
             "          A.TP_CPRO_RETN_OPER_ATIV,      " & vbCrLf & _
             "          A.VA_OPER_ATIV,                " & vbCrLf & _
             "          A.NU_CTRL_MESG_SPB_ORIG,       " & vbCrLf & _
             "          A.DT_OPER_ATIV,                " & vbCrLf & _
             "          A.CO_CNTA_CUTD_SELIC_CNPT,     " & vbCrLf & _
             "          A.NU_PRTC_OPER_LG,             " & vbCrLf & _
             "          A.CO_PRAC,                     " & vbCrLf & _
             "          A.CO_MOED_ESTR,                " & vbCrLf & _
             "          A.PE_TAXA_NEGO,                " & vbCrLf & _
             "          A.VA_MOED_ESTR,                " & vbCrLf & _
             "          A.DT_LIQU_OPER_ATIV,           " & vbCrLf & _
             "          A.CO_SISB_COTR,                " & vbCrLf & _
             "          A.NU_COMD_OPER,                " & vbCrLf & _
             "          A.TP_CNAL_VEND,                " & vbCrLf & _
             "          A.CD_SUB_PROD,                 " & vbCrLf

    strSQL = strSQL & _
             "          B.CO_TEXT_XML,                 " & vbCrLf & _
             "          C.CO_MESG_SPB_REGT_OPER,       " & vbCrLf & _
             "          A.NU_ATIV_MERC,                 " & vbCrLf & _
             "          A.CD_ASSO_CAMB                 " & vbCrLf

    strSQL = strSQL & _
             " From     A8.TB_OPER_ATIV             A, " & vbCrLf & _
             "          A8.TB_OPER_ATIV_MESG_INTE   B, " & vbCrLf & _
             "          A8.TB_TIPO_OPER             C  " & vbCrLf & _
             " Where    A.TP_OPER = C.TP_OPER          " & vbCrLf

    If pvntSequenciaOperacao > 0 Then
        
        If Not pblnSolicitacaoCancelamento Then
            
            strSQL = strSQL & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   " & pvntSequenciaOperacao & vbCrLf & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
                 "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
                 "                                      Where    NU_SEQU_OPER_ATIV   =   " & pvntSequenciaOperacao & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
                 
        Else
        
            strSQL = strSQL & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   " & pvntSequenciaOperacao & vbCrLf & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
                 "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
                 "                                      Where    NU_SEQU_OPER_ATIV   =   " & pvntSequenciaOperacao & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   =   " & enumTipoSolicitacao.CancelamentoComMensagem & ")" & vbCrLf
                 
        End If
        
    ElseIf Trim(pstrCodigoOperacao) <> "" Then
        If Trim(pstrIdentificadorLastro) <> vbNullString Then
            
            strSQL = strSQL & _
                 "   And    A.CO_OPER_ATIV          =   '" & pstrCodigoOperacao & "'" & vbCrLf & _
                 "   And    A.SG_SIST               =   '" & pstrSiglaSistema & "'" & vbCrLf & _
                 "   And    A.CO_IDEF_LAST          =   '" & pstrIdentificadorLastro & "'" & vbCrLf & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
                 "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
                 "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
                 
        ElseIf Trim(pstrIdentificadorTitulo) = vbNullString Then
            
            strSQL = strSQL & _
                 "   And    A.CO_OPER_ATIV          =   '" & pstrCodigoOperacao & "'" & vbCrLf & _
                 "   And    A.SG_SIST               =   '" & pstrSiglaSistema & "'" & vbCrLf & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
                 "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
                 "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
                 
        Else
            
            strSQL = strSQL & _
                 "   And    A.CO_OPER_ATIV          =   '" & pstrCodigoOperacao & "'" & vbCrLf & _
                 "   And    A.SG_SIST               =   '" & pstrSiglaSistema & "'" & vbCrLf & _
                 "   And    A.NU_ATIV_MERC          =   '" & pstrIdentificadorTitulo & "'" & vbCrLf & _
                 "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
                 "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
                 "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
                 "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
                 
        End If
        
    ElseIf pstrNumeroOperacaoSelic <> "" Then
        
        strSQL = strSQL & _
             "   And    A.NU_COMD_OPER          =   '" & pstrNumeroOperacaoSelic & "'" & vbCrLf & _
             "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
             "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
             "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
             
    ElseIf pstrComandoOperacaoRet <> "" Then
        
        strSQL = strSQL & _
             "   And    A.NU_COMD_OPER_RETN     =   '" & pstrComandoOperacaoRet & "'" & vbCrLf & _
             "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
             "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
             "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf

    ' Operacões BMC 130 (Balcao e Eletronica)
    ElseIf pvntNrIdefNegoBMC <> 0 Then
        
        strSQL = strSQL & _
             "   And    A.NR_IDEF_NEGO_BMC      =   " & pvntNrIdefNegoBMC & vbCrLf & _
             "   And    A.TP_OPER              IN   ('" & enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao & "', '" & enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica & "', '" & enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega & "') " & vbCrLf & _
             "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
             "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
             "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
    
    ' Operacões BMC 130 (Balcao e Eletronica) e Layout 249 - Operacao 234
    ElseIf pstrChACAM <> "" Then
        
        strSQL = strSQL & _
             "   And    A.CD_ASSO_CAMB          =   '" & pstrChACAM & "'" & vbCrLf & _
             "   And    A.TP_OPER              IN   ('" & enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao & "', '" & enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica & "', '" & enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega & "') " & vbCrLf & _
             "   And    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "   And    B.DH_MESG_INTE          =  (Select   MAX(DH_MESG_INTE) " & vbCrLf & _
             "                                      From     A8.TB_OPER_ATIV_MESG_INTE " & vbCrLf & _
             "                                      Where    NU_SEQU_OPER_ATIV   =   A.NU_SEQU_OPER_ATIV " & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.CancelamentoComMensagem & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.RetornoLegado & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.Alteracao & vbCrLf & _
             "                                      And      TP_SOLI_MESG_INTE   <>  " & enumTipoSolicitacao.LivreMovimentacao & ")" & vbCrLf
             
    End If

    If pblnForUpdate And Not fgEstacaoDesenvolvimento Then
        strSQL = strSQL & " For Update"
    End If

    Set rsOperacao = fgQuerySQL(strSQL)

    If rsOperacao.EOF Then
'        ' 3013 - Operação já processada por outro usuário.
'        lngCodigoErroNegocio = 3013
'        GoTo ErrorHandler
        Exit Function
    End If
    
    If Not pobjRs Is Nothing Then
        Set pobjRs = rsOperacao
    End If

    LerXMLOperacao.loadXML fgSelectVarchar4000(rsOperacao!CO_TEXT_XML)

    'Adiciona o tipo da operação, pois na mensagem original não existe a mesma.
    If Not IsNull(rsOperacao!HO_ENVI_MESG_SPB) Then
        If rsOperacao!HO_ENVI_MESG_SPB <> LerXMLOperacao.documentElement.selectSingleNode("HO_AGND").Text Then
            If LerXMLOperacao.documentElement.selectSingleNode("HO_AGND").Text = vbNullString Or LerXMLOperacao.documentElement.selectSingleNode("HO_AGND").Text = "0" Then
                LerXMLOperacao.documentElement.selectSingleNode("HO_AGND").Text = Format(rsOperacao!HO_ENVI_MESG_SPB, "HHMM")
            End If
        End If
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "TP_OPER", IIf(IsNull(rsOperacao!TP_OPER), 0, rsOperacao!TP_OPER)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("TP_OPER").Text = IIf(IsNull(rsOperacao!TP_OPER), 0, rsOperacao!TP_OPER)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER") Is Nothing Then
        If IsNull(rsOperacao!CO_MESG_SPB_REGT_OPER) Then
            fgAppendNode LerXMLOperacao, "MESG", "CO_MESG_SPB_REGT_OPER", vbNullString
        Else
            fgAppendNode LerXMLOperacao, "MESG", "CO_MESG_SPB_REGT_OPER", Trim$(rsOperacao!CO_MESG_SPB_REGT_OPER)
        End If
    Else
        If IsNull(rsOperacao!CO_MESG_SPB_REGT_OPER) Then
            LerXMLOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text = vbNullString
        Else
            LerXMLOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text = Trim$(rsOperacao!CO_MESG_SPB_REGT_OPER)
        End If
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "NU_COMD_OPER", IIf(IsNull(rsOperacao!NU_COMD_OPER), "", rsOperacao!NU_COMD_OPER)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text = IIf(IsNull(rsOperacao!NU_COMD_OPER), "", rsOperacao!NU_COMD_OPER)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        If Not IsNull(rsOperacao!DT_OPER_ATIV) Then
            fgAppendNode LerXMLOperacao, "MESG", "DT_OPER_ATIV", fgDt_To_Xml(rsOperacao!DT_OPER_ATIV)
        Else
            fgAppendNode LerXMLOperacao, "MESG", "DT_OPER_ATIV", ""
        End If
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "NU_SEQU_OPER_ATIV", IIf(IsNull(rsOperacao!NU_SEQU_OPER_ATIV), 0, rsOperacao!NU_SEQU_OPER_ATIV)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = IIf(IsNull(rsOperacao!NU_SEQU_OPER_ATIV), 0, rsOperacao!NU_SEQU_OPER_ATIV)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CO_ULTI_SITU_PROC", IIf(IsNull(rsOperacao!CO_ULTI_SITU_PROC), 0, rsOperacao!CO_ULTI_SITU_PROC)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = IIf(IsNull(rsOperacao!CO_ULTI_SITU_PROC), 0, rsOperacao!CO_ULTI_SITU_PROC)
    End If
    
    If LerXMLOperacao.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "IN_ENTR_MANU", IIf(IsNull(rsOperacao!IN_ENTR_MANU), enumIndicadorSimNao.Nao, rsOperacao!IN_ENTR_MANU)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("IN_ENTR_MANU").Text = IIf(IsNull(rsOperacao!IN_ENTR_MANU), enumIndicadorSimNao.Nao, rsOperacao!IN_ENTR_MANU)
    End If
    

    If LerXMLOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "TP_ACAO_OPER_ATIV_EXEC", IIf(IsNull(rsOperacao!TP_ACAO_OPER_ATIV_EXEC), 0, rsOperacao!TP_ACAO_OPER_ATIV_EXEC)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = IIf(IsNull(rsOperacao!TP_ACAO_OPER_ATIV_EXEC), 0, rsOperacao!TP_ACAO_OPER_ATIV_EXEC)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "NU_CTRL_MESG_SPB_ORIG", IIf(IsNull(rsOperacao!NU_CTRL_MESG_SPB_ORIG), 0, rsOperacao!NU_CTRL_MESG_SPB_ORIG)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text = IIf(IsNull(rsOperacao!NU_CTRL_MESG_SPB_ORIG), 0, rsOperacao!NU_CTRL_MESG_SPB_ORIG)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("NU_COMD_ACAO_EXEC") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "NU_COMD_ACAO_EXEC", IIf(IsNull(rsOperacao!NU_COMD_ACAO_EXEC), 0, rsOperacao!NU_COMD_ACAO_EXEC)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("NU_COMD_ACAO_EXEC").Text = IIf(IsNull(rsOperacao!NU_COMD_ACAO_EXEC), 0, rsOperacao!NU_COMD_ACAO_EXEC)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CO_CNTA_CUTD_SELIC_CNPT", IIf(IsNull(rsOperacao!CO_CNTA_CUTD_SELIC_CNPT), 0, rsOperacao!CO_CNTA_CUTD_SELIC_CNPT)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text = IIf(IsNull(rsOperacao!CO_CNTA_CUTD_SELIC_CNPT), 0, rsOperacao!CO_CNTA_CUTD_SELIC_CNPT)
    End If
    
    If Not IsNull(rsOperacao!dh_ulti_atlz) Then
        If LerXMLOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ") Is Nothing Then
            fgAppendNode LerXMLOperacao, "MESG", "DH_ULTI_ATLZ", fgDtHr_To_Xml(rsOperacao!dh_ulti_atlz)
        Else
            LerXMLOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text = fgDtHr_To_Xml(rsOperacao!dh_ulti_atlz)
        End If
    End If

    If Not IsNull(rsOperacao!NU_PRTC_OPER_LG) Then
        If LerXMLOperacao.documentElement.selectSingleNode("NU_PRTC_OPER_LG") Is Nothing Then
            fgAppendNode LerXMLOperacao, "MESG", "NU_PRTC_OPER_LG", rsOperacao!NU_PRTC_OPER_LG
        Else
            LerXMLOperacao.documentElement.selectSingleNode("NU_PRTC_OPER_LG").Text = rsOperacao!NU_PRTC_OPER_LG
        End If
    Else
        fgAppendNode LerXMLOperacao, "MESG", "NU_PRTC_OPER_LG", ""
    End If

    If Not IsNull(rsOperacao!va_oper_ativ) Then
        If Not LerXMLOperacao.documentElement.selectSingleNode("VA_OPER_ATIV") Is Nothing Then
            If LerXMLOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text <> fgVlr_To_Xml(rsOperacao!va_oper_ativ) Then
                LerXMLOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text = fgVlr_To_Xml(rsOperacao!va_oper_ativ)
            End If
        End If
    End If

    If Not LerXMLOperacao.documentElement.selectSingleNode("TP_CPRO_OPER_ATIV") Is Nothing Then
        If LerXMLOperacao.documentElement.selectSingleNode("TP_CPRO_OPER_ATIV").Text <> rsOperacao!TP_CPRO_OPER_ATIV And _
           Not IsNull(rsOperacao!TP_CPRO_OPER_ATIV) And _
           rsOperacao!TP_CPRO_OPER_ATIV <> vbNullString Then
           LerXMLOperacao.documentElement.selectSingleNode("TP_CPRO_OPER_ATIV").Text = rsOperacao!TP_CPRO_OPER_ATIV
        End If
    End If
    If Not LerXMLOperacao.documentElement.selectSingleNode("TP_CPRO_RETN_OPER_ATIV") Is Nothing Then
        If LerXMLOperacao.documentElement.selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text <> rsOperacao!TP_CPRO_RETN_OPER_ATIV And _
           Not IsNull(rsOperacao!TP_CPRO_RETN_OPER_ATIV) And _
           rsOperacao!TP_CPRO_RETN_OPER_ATIV <> vbNullString Then
           LerXMLOperacao.documentElement.selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text = rsOperacao!TP_CPRO_RETN_OPER_ATIV
        End If
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CO_PRAC") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CO_PRAC", IIf(IsNull(rsOperacao!CO_PRAC), "", rsOperacao!CO_PRAC)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CO_PRAC").Text = IIf(IsNull(rsOperacao!CO_PRAC), "", rsOperacao!CO_PRAC)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CO_MOED_ESTR") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CO_MOED_ESTR", IIf(IsNull(rsOperacao!CO_MOED_ESTR), "", rsOperacao!CO_MOED_ESTR)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CO_MOED_ESTR").Text = IIf(IsNull(rsOperacao!CO_MOED_ESTR), "", rsOperacao!CO_MOED_ESTR)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("PE_TAXA_NEGO") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "PE_TAXA_NEGO", IIf(IsNull(rsOperacao!PE_TAXA_NEGO), "", rsOperacao!PE_TAXA_NEGO)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("PE_TAXA_NEGO").Text = IIf(IsNull(rsOperacao!PE_TAXA_NEGO), "", rsOperacao!PE_TAXA_NEGO)
    End If
    
    'RATS 394
    LerXMLOperacao.documentElement.selectSingleNode("PE_TAXA_NEGO").Text = Replace(LerXMLOperacao.documentElement.selectSingleNode("PE_TAXA_NEGO").Text, ".", ",")

    If LerXMLOperacao.documentElement.selectSingleNode("VA_MOED_ESTR") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "VA_MOED_ESTR", IIf(IsNull(rsOperacao!VA_MOED_ESTR), "", rsOperacao!VA_MOED_ESTR)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("VA_MOED_ESTR").Text = IIf(IsNull(rsOperacao!VA_MOED_ESTR), "", rsOperacao!VA_MOED_ESTR)
    End If
    
    LerXMLOperacao.documentElement.selectSingleNode("VA_MOED_ESTR").Text = Replace(LerXMLOperacao.documentElement.selectSingleNode("VA_MOED_ESTR").Text, ".", ",")
    
    If LerXMLOperacao.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing Then
        If Not IsNull(rsOperacao!DT_LIQU_OPER_ATIV) Then
            fgAppendNode LerXMLOperacao, "MESG", "DT_LIQU_OPER_ATIV", fgDt_To_Xml(rsOperacao!DT_LIQU_OPER_ATIV)
        End If
    Else
        If Not IsNull(rsOperacao!DT_LIQU_OPER_ATIV) Then
             LerXMLOperacao.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text = fgDt_To_Xml(rsOperacao!DT_LIQU_OPER_ATIV)
        End If
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CO_SISB_COTR") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CO_SISB_COTR", IIf(IsNull(rsOperacao!CO_SISB_COTR), "", rsOperacao!CO_SISB_COTR)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CO_SISB_COTR").Text = IIf(IsNull(rsOperacao!CO_SISB_COTR), "", rsOperacao!CO_SISB_COTR)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("TP_CNAL_VEND") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "TP_CNAL_VEND", IIf(IsNull(rsOperacao!TP_CNAL_VEND), "", rsOperacao!TP_CNAL_VEND)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("TP_CNAL_VEND").Text = IIf(IsNull(rsOperacao!TP_CNAL_VEND), "", rsOperacao!TP_CNAL_VEND)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CD_SUB_PROD") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CD_SUB_PROD", IIf(IsNull(rsOperacao!CD_SUB_PROD), "", rsOperacao!CD_SUB_PROD)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CD_SUB_PROD").Text = IIf(IsNull(rsOperacao!CD_SUB_PROD), "", rsOperacao!CD_SUB_PROD)
    End If

    If LerXMLOperacao.documentElement.selectSingleNode("CD_ASSO_CAMB") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "CD_ASSO_CAMB", IIf(IsNull(rsOperacao!CD_ASSO_CAMB), "", rsOperacao!CD_ASSO_CAMB)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("CD_ASSO_CAMB").Text = IIf(IsNull(rsOperacao!CD_ASSO_CAMB), "", rsOperacao!CD_ASSO_CAMB)
    End If

    'RATS - 942
    If LerXMLOperacao.documentElement.selectSingleNode("NU_ATIV_MERC") Is Nothing Then
        fgAppendNode LerXMLOperacao, "MESG", "NU_ATIV_MERC", IIf(IsNull(rsOperacao!NU_ATIV_MERC), "", rsOperacao!NU_ATIV_MERC)
    Else
        LerXMLOperacao.documentElement.selectSingleNode("NU_ATIV_MERC").Text = IIf(IsNull(rsOperacao!NU_ATIV_MERC), "", rsOperacao!NU_ATIV_MERC)
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerXMLOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de liberação de uma operação
'-------------------------------------------------------------------------------------------------

Public Function Liberar(ByVal pvntSequenciaOperacao As Variant, _
                        ByVal pintStatusOperacao As Integer, _
                        ByVal pintTipoAcaoAnterior As Integer, _
                        ByVal plngNumeroComandoAcao As Long, _
                        ByVal pstrDHUltimaAtualizacao As String, _
                        ByVal pblnIgnoraGradeHorario As Boolean)

Dim intAcao                                 As Integer '<-- enumTipoAcao
Dim intStatusOperacao                       As Integer '<-- enumStatusOperacao
Dim strAux                                  As String
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim objAlcadasLimites                       As A6A7A8.clsPerfil

On Error GoTo ErrorHandler

    Select Case pintTipoAcaoAnterior
'        Case enumTipoAcao.CancelamentoSolicitado
'            intAcao = enumTipoAcao.CancelamentoEnviado
'            intStatusOperacao = pintStatusOperacao
'        Case enumTipoAcao.EstornoSolicitado
'            intAcao = enumTipoAcao.EstornoEnviado
'            intStatusOperacao = pintStatusOperacao
        Case 0
'            intAcao = 0
'            intStatusOperacao = enumStatusOperacao.Liberada
'            Call AlterarStatusAcao(pvntSequenciaOperacao, _
'                                   intStatusOperacao, _
'                                   pstrDHUltimaAtualizacao, _
'                                   , intAcao)
    End Select

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlOperacao.loadXML(Ler(pvntSequenciaOperacao)) Then
        Err.Raise 1, , "Operação não encontrada"
    End If
    
    'Se operação 'Leilao BMA', tem que passar um XML com campos extras para a montagem da LDL1002
    If fgSelectSingleNodeText(xmlOperacao, "//TP_OPER") = CStr(enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA) Then
        CamposExtrasLeilaoBMA pvntSequenciaOperacao, xmlRemessa
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    If xmlRemessa Is Nothing Then
        objProcessoOperacao.GerenciarChamadaViaTela pvntSequenciaOperacao, _
                                                    enumFuncionalidade.Liberacao, _
                                                    pblnIgnoraGradeHorario, _
                                                    , , , , _
                                                    pstrDHUltimaAtualizacao
    Else
        objProcessoOperacao.GerenciarChamadaViaTela pvntSequenciaOperacao, _
                                                    enumFuncionalidade.Liberacao, _
                                                    pblnIgnoraGradeHorario, _
                                                    , , xmlRemessa, , _
                                                    pstrDHUltimaAtualizacao
    End If
    
    Set objProcessoOperacao = Nothing

    'Pikachu - 30/09/2004
    'Validação da alçada - Perfil Liberador
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(pvntSequenciaOperacao)
    
    If Not xmlOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
        If xmlOperacao.selectSingleNode("//IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
            If Not xmlOperacao.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                Set objAlcadasLimites = CreateObject("A6A7A8.clsPerfil")
                Call objAlcadasLimites.AprovarAlcada(fgUsuarioRede, _
                                                     xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text, _
                                                     xmlOperacao.selectSingleNode("//NU_PRTC_OPER_LG").Text)
                Set objAlcadasLimites = Nothing
            End If
        End If
    End If

    '---------------------------------------------------------------------------------
    
    If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DespesasBMC Then
            
            Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Call xmlMensagem.loadXML(objConciliacao.ConsultaConciliacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
            Set objConciliacao = Nothing
            
            If xmlMensagem.xml <> vbNullString Then
                Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Call objMensagem.AlterarStatus(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                               xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                               enumStatusMensagem.PagamentoLib, _
                                               vbNullString, _
                                               , _
                                               False, _
                                               xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                Set objMensagem = Nothing
            End If
        
        End If
    End If
    
     Set objOperacao = Nothing
     Set xmlOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set objAlcadasLimites = Nothing
    Set objProcessoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Liberar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de liberação de um registro
'-------------------------------------------------------------------------------------------------

Public Function LiberarRegistro(ByVal plngLocalLiquidacao As Long, _
                                ByVal pvntSequenciaOperacao As Variant, _
                                ByVal pintTipoOperacao As Variant, _
                                ByVal pstrDHUltimaAtualizacao As String, _
                                ByVal pblnIgnoraGradeHorario As Boolean, _
                       Optional ByVal pblnFluxoAutomatico As Boolean = True)

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim vntStatusOperacao                       As Variant

    On Error GoTo ErrorHandler

    vntStatusOperacao = Null
    Select Case plngLocalLiquidacao
        Case enumLocalLiquidacao.BMA
            'DRU: BMA - Registro de Operacoes Definitivas
            If pintTipoOperacao = enumTipoOperacaoLQS.OperacaoTermoCobertaBMA Or _
               pintTipoOperacao = enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA Or _
               pintTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Then
                    
                vntStatusOperacao = IIf(pblnFluxoAutomatico, enumStatusOperacao.LiquidadaFisicamenteAutomatica, enumStatusOperacao.LiquidadaFisicamente)
            
            ElseIf pintTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaAVista _
                Or pintTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
                
                'Nao mudar status, a GerenciarChamadaViaTela fará isso
            
            Else
                vntStatusOperacao = IIf(pblnFluxoAutomatico, enumStatusOperacao.RegistradaAutomatica, enumStatusOperacao.Registrada)
            End If
        Case enumLocalLiquidacao.CETIP
            'DRU: CETIP - Registro de Operacoes
            
            'Nao mudar status, a GerenciarChamadaViaTela fará isso
            vntStatusOperacao = Null
    End Select

    If Not IsNull(vntStatusOperacao) Then
        Call AlterarStatusAcao(pvntSequenciaOperacao, _
                               vntStatusOperacao, _
                               pstrDHUltimaAtualizacao _
                                )
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.GerenciarChamadaViaTela pvntSequenciaOperacao, _
                                                enumFuncionalidade.Liberacao, _
                                                pblnIgnoraGradeHorario
    Set objProcessoOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set objProcessoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarRegistro Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter a propriedade da tabela de operação
'-------------------------------------------------------------------------------------------------

Private Function flObterPropriedades() As ADODB.Recordset

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Select   *   " & vbCrLf & _
             "   From   A8.TB_OPER_ATIV Where 1 = 2"

    Set flObterPropriedades = fgQuerySQL(strSQL)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Ler as operações por tipo de operação especifico
'-------------------------------------------------------------------------------------------------

Public Function ObterOperacoesPorTipoOperacao(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String
Dim strOwnerTabelaSegregada                 As String
Dim strOwnerTabelaSegregadaHist             As String
Dim strSegregacao                           As String

Dim xmlComplementoWhere                     As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strValores                              As String
Dim blnOcultarQuantidades                   As Boolean

Dim blnBaseA8                               As Boolean
Dim blnBaseHistorica                        As Boolean

Dim datDataCorte                            As Date
Dim datInicio                               As Date
Dim datFim                                  As Date

Dim intTipoBackOffice                       As Integer
                                                                
On Error GoTo ErrorHandler

    'Verifica se é necessário exibir as quantidades
    blnOcultarQuantidades = (Not _
        xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Quantidade/OcultarQuantidade") Is Nothing)
    
    If Not blnOcultarQuantidades Then
    
        datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
        datInicio = gstrDataVazia
        datFim = gstrDataVazia
    
        If Not xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
            datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
        End If
    
        If Not xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
            datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
        End If
        
        If datInicio <> gstrDataVazia Then
            blnBaseHistorica = datInicio < datDataCorte
        End If
    
        If datFim <> gstrDataVazia Then
            blnBaseA8 = datFim > datDataCorte
        End If

        'Envia dados para a segregação ----------------------------------------------------------------------------
        Set xmlComplementoWhere = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlComplementoWhere, "", "Repeat_Where", "")
        Call fgAppendNode(xmlComplementoWhere, "Repeat_Where", "Grupo_Where", "")
        
        If xmlFiltro.selectNodes("Repeat_Filtros/*").length > 0 Then
            'Filtro Grupo Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_GRUP_VEIC_LEGA      IN (" & strValores & ")")
                'Inclui o atributo ALIAS para filtro específico na tabela DOMÍNIO (Segregação de Dados)
                Call fgAppendAttribute(xmlComplementoWhere, "Where", "Alias", "TabDominio_Sub")
            End If
            
            'Filtro por Tipo Operação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
        
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_OPER      IN (" & strValores & ")")
            End If
        
            'Filtro por Status
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SituacaoOperacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_ULTI_SITU_PROC      IN (" & strValores & ")")
            End If
            
            'Filtro por Empresa
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_EMPR      IN (" & strValores & ")")
            End If
            
            'Filtro Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_VEIC_LEGA      IN (" & strValores & ")")
            End If
            
            'Filtro Local de Liquidação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_LOCA_LIQU      IN (" & strValores & ")")
            End If
            
            'Filtro Contraparte
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Contraparte/*")
                strValores = objDomNode.Text
            Next
            
            If Trim(strValores) <> "" Then
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "NO_CNPT     LIKE '" & strValores & "%'")
            End If
            
            'Filtro Número do Comando
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroComando/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "NU_COMD_OPER      IN (" & strValores & ")")
            End If
            
            'Filtro por Ação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Acoes/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_ACAO_OPER_ATIV_EXEC  IN (" & strValores & ")")
            End If
            
            
            'KIDA - SGC
            'Filtro Grupo Canal Venda
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                'strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_CNAL_VEND  IN (" & strValores & ")")
            End If
            
            'Filtro Número Controle LTR / Cod Reembolso CCR
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroControleLTR/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "NU_CTRL_MESG_SPB_ORIG IN (" & strValores & ")")
            End If
            
            
            'Filtro por Data
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", _
                                    "DT_OPER_ATIV BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                                                    " AND " & objDomNode.selectSingleNode("DataFim").Text)
            Next objDomNode
        End If
        
        Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "IN_DISP_CONS = " & enumIndicadorSimNao.Sim)
        
        intTipoBackOffice = 99
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(objDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                If blnBaseA8 Then
                    strOwnerTabelaSegregada = fgSegregaDados("A8.TB_OPER_ATIV", , , , , , , , , False, xmlComplementoWhere)
                End If
                
                If blnBaseHistorica Then
                    strOwnerTabelaSegregadaHist = fgSegregaDados("A8HIST.TB_OPER_ATIV", , , , , , , , , False, xmlComplementoWhere)
                    If blnBaseA8 Then
                        strSegregacao = strOwnerTabelaSegregada & vbNewLine & " UNION ALL " & strOwnerTabelaSegregadaHist & vbNewLine
                    Else
                        strSegregacao = strOwnerTabelaSegregadaHist
                    End If
                Else
                    strSegregacao = strOwnerTabelaSegregada
                End If
        Set xmlComplementoWhere = Nothing
            Case 0
                If blnBaseA8 Then
                    strOwnerTabelaSegregada = fgSegregaDados("A8.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                End If
                
                If blnBaseHistorica Then
                    strOwnerTabelaSegregadaHist = fgSegregaDados("A8HIST.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                    If blnBaseA8 Then
                        strSegregacao = strOwnerTabelaSegregada & vbNewLine & " UNION ALL " & strOwnerTabelaSegregadaHist & vbNewLine
                    Else
                        strSegregacao = strOwnerTabelaSegregadaHist
                    End If
                Else
                    strSegregacao = strOwnerTabelaSegregada
                End If
            Case Is > 0
                If blnBaseA8 Then
                    strOwnerTabelaSegregada = fgSegregaDados("A8.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                    strOwnerTabelaSegregada = strOwnerTabelaSegregada & vbNewLine & " And TabDominio_Sub.TP_BKOF = " & intTipoBackOffice
                End If
                
                If blnBaseHistorica Then
                    strOwnerTabelaSegregadaHist = fgSegregaDados("A8HIST.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                    strOwnerTabelaSegregadaHist = strOwnerTabelaSegregada & vbNewLine & " And TabDominio_Sub.TP_BKOF = " & intTipoBackOffice
                    If blnBaseA8 Then
                        strSegregacao = strOwnerTabelaSegregada & vbNewLine & " UNION ALL " & strOwnerTabelaSegregadaHist & vbNewLine
                    Else
                        strSegregacao = strOwnerTabelaSegregadaHist
                    End If
                Else
                    strSegregacao = strOwnerTabelaSegregada
                End If
        End Select
    
        '----------------------------------------------------------------------------------------------------------
    End If
    
    If blnOcultarQuantidades Then
        strSQL = " Select    A.TP_OPER,                                               " & vbNewLine & _
                 "           A.NO_TIPO_OPER,                                          " & vbNewLine & _
                 "           A.CO_LOCA_LIQU,                                          " & vbNewLine & _
                 "           A.TP_MESG_RECB_INTE,                                     " & vbNewLine & _
                 "           A.TP_MESG_RETN_INTE,                                     " & vbNewLine & _
                 "           A.CO_MESG_SPB_REGT_OPER,                                 " & vbNewLine & _
                 "           A.DT_INIC_VIGE,                                          " & vbNewLine & _
                 "           A.DT_FIM_VIGE,                                           " & vbNewLine & _
                 "           A.CO_USUA_ULTI_ATLZ,                                     " & vbNewLine & _
                 "           A.CO_ETCA_TRAB_ULTI_ATLZ,                                " & vbNewLine & _
                 "           A.DH_ULTI_ATLZ,                                          " & vbNewLine & _
                 "           L.SG_LOCA_LIQU                                           " & vbNewLine & _
                 " From      A8.TB_TIPO_OPER        A,                                " & vbNewLine & _
                 "           A8.TB_LOCAL_LIQUIDACAO L                                 " & vbNewLine & _
                 " Where    (A.DT_FIM_VIGE          IS  NULL                          " & vbNewLine & _
                 " Or        A.DT_FIM_VIGE          >=  SYSDATE)                      " & vbNewLine & _
                 " And       A.CO_LOCA_LIQU         = L.CO_LOCA_LIQU (+)              " & vbNewLine & _
                 " And       1 /* Fixo Santander*/  = L.CO_EMPR_FUSI (+)              " & vbNewLine & _
                 " Order By  L.SG_LOCA_LIQU, A.NO_TIPO_OPER ASC"
    Else
        strSQL = " Select    A.CO_LOCA_LIQU,                                          " & vbNewLine & _
                 "           A.NO_TIPO_OPER,                                          " & vbNewLine & _
                 "           A.TP_OPER,                                               " & vbNewLine & _
                 "           A.TP_MESG_RECB_INTE,                                     " & vbNewLine & _
                 "           A.TP_MESG_RETN_INTE,                                     " & vbNewLine & _
                 "           A.CO_MESG_SPB_REGT_OPER,                                 " & vbNewLine & _
                 "           A.DT_INIC_VIGE,                                          " & vbNewLine & _
                 "           A.DT_FIM_VIGE,                                           " & vbNewLine & _
                 "           A.CO_USUA_ULTI_ATLZ,                                     " & vbNewLine & _
                 "           A.CO_ETCA_TRAB_ULTI_ATLZ,                                " & vbNewLine & _
                 "           A.DH_ULTI_ATLZ,                                          " & vbNewLine & _
                 "           L.SG_LOCA_LIQU,                                          " & vbNewLine & _
                 "           COUNT(B.TP_OPER)       NU_QTD                            " & vbNewLine & _
                 " From      A8.TB_TIPO_OPER           A,                             " & vbNewLine & _
                 "         (" & strSegregacao & ")     B,                             " & vbNewLine & _
                 "           A8.TB_LOCAL_LIQUIDACAO    L                              " & vbNewLine & _
                 " Where     A.TP_OPER              = B.TP_OPER             (+)       " & vbNewLine & _
                 " And      (A.DT_FIM_VIGE          IS  NULL                          " & vbNewLine & _
                 " Or        A.DT_FIM_VIGE          >=  SYSDATE)                      " & vbNewLine & _
                 " And       A.CO_LOCA_LIQU         = L.CO_LOCA_LIQU (+)              " & vbNewLine & _
                 " And       1 /* Fixo Santander*/  = L.CO_EMPR_FUSI (+)              " & vbNewLine

        strSQL = strSQL & _
                 " Group By  A.CO_LOCA_LIQU,                                          " & vbNewLine & _
                 "           A.NO_TIPO_OPER,                                          " & vbNewLine & _
                 "           A.TP_OPER,                                               " & vbNewLine & _
                 "           A.TP_MESG_RECB_INTE,                                     " & vbNewLine & _
                 "           A.TP_MESG_RETN_INTE,                                     " & vbNewLine & _
                 "           A.CO_MESG_SPB_REGT_OPER,                                 " & vbNewLine & _
                 "           A.DT_INIC_VIGE,                                          " & vbNewLine & _
                 "           A.DT_FIM_VIGE,                                           " & vbNewLine & _
                 "           A.CO_USUA_ULTI_ATLZ,                                     " & vbNewLine & _
                 "           A.CO_ETCA_TRAB_ULTI_ATLZ,                                " & vbNewLine & _
                 "           A.DH_ULTI_ATLZ,                                          " & vbNewLine & _
                 "           L.SG_LOCA_LIQU                                           " & vbNewLine & _
                 " Order By  L.SG_LOCA_LIQU, A.NO_TIPO_OPER ASC"
    End If

    ObterOperacoesPorTipoOperacao = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacoesPorTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Ler as operações pelo status epecificado no filtro
'-------------------------------------------------------------------------------------------------

Public Function ObterOperacoesPorStatus(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String
Dim strOwnerTabelaSegregada                 As String
Dim strOwnerTabelaSegregadaHist             As String
Dim strSegregacao                           As String

Dim xmlComplementoWhere                     As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strValores                              As String
Dim blnOcultarQuantidades                   As Boolean

Dim blnBaseA8                               As Boolean
Dim blnBaseHistorica                        As Boolean

Dim datDataCorte                            As Date
Dim datInicio                               As Date
Dim datFim                                  As Date

Dim intTipoBackOffice                       As Integer
                                                                
    On Error GoTo ErrorHandler

    'Verifica se é necessário exibir as quantidades
    blnOcultarQuantidades = (Not _
        xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Quantidade/OcultarQuantidade") Is Nothing)
    
    If Not blnOcultarQuantidades Then
    
        datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
        datInicio = gstrDataVazia
        datFim = gstrDataVazia
    
        If Not xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
            datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
        End If
    
        If Not xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
            datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
        End If
        
        If datInicio <> gstrDataVazia Then
            blnBaseHistorica = datInicio < datDataCorte
        End If
    
        If datFim <> gstrDataVazia Then
            blnBaseA8 = datFim > datDataCorte
        End If
            
        'Envia dados para a segregação ----------------------------------------------------------------------------
        Set xmlComplementoWhere = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlComplementoWhere, "", "Repeat_Where", "")
        Call fgAppendNode(xmlComplementoWhere, "Repeat_Where", "Grupo_Where", "")
        
        If xmlFiltro.selectNodes("Repeat_Filtros/*").length > 0 Then
            'Filtro Grupo Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_GRUP_VEIC_LEGA      IN (" & strValores & ")")
                'Inclui o atributo ALIAS para filtro específico na tabela DOMÍNIO (Segregação de Dados)
                Call fgAppendAttribute(xmlComplementoWhere, "Where", "Alias", "TabDominio_Sub")
            End If
            
            'Filtro por Tipo Operação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
        
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_OPER      IN (" & strValores & ")")
            End If
        
            'Filtro por Status
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SituacaoOperacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_ULTI_SITU_PROC      IN (" & strValores & ")")
            End If
            
            'Filtro por Empresa
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_EMPR      IN (" & strValores & ")")
            End If
            
            'Filtro Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_VEIC_LEGA      IN (" & strValores & ")")
            End If
            
            'Filtro Local de Liquidação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_LOCA_LIQU      IN (" & strValores & ")")
            End If
            
            'Filtro Contraparte
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Contraparte/*")
                strValores = objDomNode.Text
            Next
            
            If Trim(strValores) <> "" Then
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "NO_CNPT     LIKE '" & strValores & "%'")
            End If
            
            'Filtro Número do Comando
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroComando/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "NU_COMD_OPER      IN (" & strValores & ")")
            End If
            
            'Filtro por Ação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Acoes/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_ACAO_OPER_ATIV_EXEC  IN (" & strValores & ")")
            End If
            
            'Filtro Número Controle LTR / Cod Reembolso CCR
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroControleLTR/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "NU_CTRL_MESG_SPB_ORIG  IN (" & strValores & ")")
            End If
            
            'KIDA - SGC
            'Filtro Grupo Canal Venda
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                'strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_CNAL_VEND  IN (" & strValores & ")")
            End If
            
            'Filtro por Data
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", _
                                    "DT_OPER_ATIV BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                                                    " AND " & objDomNode.selectSingleNode("DataFim").Text)
            Next objDomNode
        End If
        
        Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "IN_DISP_CONS = " & enumIndicadorSimNao.Sim)
        
        intTipoBackOffice = 99
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(objDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                If blnBaseA8 Then
                    strOwnerTabelaSegregada = fgSegregaDados("A8.TB_OPER_ATIV", , , , , , , , , False, xmlComplementoWhere)
                End If
                
                If blnBaseHistorica Then
                    strOwnerTabelaSegregadaHist = fgSegregaDados("A8HIST.TB_OPER_ATIV", , , , , , , , , False, xmlComplementoWhere)
                    If blnBaseA8 Then
                        strSegregacao = strOwnerTabelaSegregada & vbNewLine & " UNION ALL " & strOwnerTabelaSegregadaHist & vbNewLine
                    Else
                        strSegregacao = strOwnerTabelaSegregadaHist
                    End If
                Else
                    strSegregacao = strOwnerTabelaSegregada
                End If
            Case 0
                If blnBaseA8 Then
                    strOwnerTabelaSegregada = fgSegregaDados("A8.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                End If
                
                If blnBaseHistorica Then
                    strOwnerTabelaSegregadaHist = fgSegregaDados("A8HIST.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                    If blnBaseA8 Then
                        strSegregacao = strOwnerTabelaSegregada & vbNewLine & " UNION ALL " & strOwnerTabelaSegregadaHist & vbNewLine
                    Else
                        strSegregacao = strOwnerTabelaSegregadaHist
                    End If
                Else
                    strSegregacao = strOwnerTabelaSegregada
                End If
            Case Is > 0
                If blnBaseA8 Then
                    strOwnerTabelaSegregada = fgSegregaDados("A8.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                    strOwnerTabelaSegregada = strOwnerTabelaSegregada & vbNewLine & " And TabDominio_Sub.TP_BKOF = " & intTipoBackOffice
                End If
                
                If blnBaseHistorica Then
                    strOwnerTabelaSegregadaHist = fgSegregaDados("A8HIST.TB_OPER_ATIV", , , , , , False, False, False, False, xmlComplementoWhere)
                    strOwnerTabelaSegregadaHist = strOwnerTabelaSegregadaHist & vbNewLine & " And TabDominio_Sub.TP_BKOF = " & intTipoBackOffice
                    If blnBaseA8 Then
                        strSegregacao = strOwnerTabelaSegregada & vbNewLine & " UNION ALL " & strOwnerTabelaSegregadaHist & vbNewLine
                    Else
                        strSegregacao = strOwnerTabelaSegregadaHist
                    End If
                Else
                    strSegregacao = strOwnerTabelaSegregada
                End If
        End Select
    
        Set xmlComplementoWhere = Nothing
        '----------------------------------------------------------------------------------------------------------
    End If
    
    If blnOcultarQuantidades Then
        strSQL = " SELECT   A.CO_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.DE_SITU_PROC                                                         " & vbNewLine & _
                 " FROM     A8.TB_SITU_PROC     A                                                  " & vbNewLine & _
                 " WHERE    A.TP_SITU_PROC      = " & enumTipoSituacaoSistema.OperacaoComAtivo & " " & vbNewLine & _
                 " ORDER BY A.DE_SITU_PROC      ASC"
    Else
        strSQL = " SELECT   'A8' AS OWNER,                                                         " & vbNewLine & _
                 "          A.CO_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.DE_SITU_PROC,                                                        " & vbNewLine & _
                 "          COUNT(B.CO_ULTI_SITU_PROC)   AS NU_QTD                                 " & vbNewLine & _
                 " FROM     A8.TB_SITU_PROC     A,                                                 " & vbNewLine & _
                 "         (" & strSegregacao & ")              B                                 " & vbNewLine & _
                 " WHERE    A.CO_SITU_PROC      = B.CO_ULTI_SITU_PROC    (+)                       " & vbNewLine & _
                 " AND      A.TP_SITU_PROC      = " & enumTipoSituacaoSistema.OperacaoComAtivo & " " & vbNewLine & _
                 " GROUP BY A.CO_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.DE_SITU_PROC                                                         " & vbNewLine & _
                 " ORDER BY A.DE_SITU_PROC      ASC"
                 
    End If
    
    ObterOperacoesPorStatus = fgQueryXMLLerTodos("SituacaoSistema", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set xmlComplementoWhere = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacoesPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter as operações conforme filtro especificado
'-------------------------------------------------------------------------------------------------

Public Function ObterDetalheOperacao(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                            Optional ByVal pblnAutomatico As Boolean = False) As String

Dim datInicio                               As Date
Dim datFim                                  As Date
Dim datDataCorte                            As Date

Dim blnChecaBaseHistórica                   As Boolean

Dim blnBaseA8                               As Boolean
Dim blnBaseHistorica                        As Boolean

Dim strSqlA8                                As String
Dim strSQLHist                              As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    datInicio = gstrDataVazia
    datFim = gstrDataVazia
    datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo

    blnChecaBaseHistórica = Not xmlFiltro.selectSingleNode("/Repeat_Filtros/BaseHistorica") Is Nothing

    If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data") Is Nothing Then
    
        If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
            datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
        End If

        If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
            datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
        End If

    End If

    If datInicio <> gstrDataVazia Then
        blnBaseHistorica = datInicio < datDataCorte And blnChecaBaseHistórica
    End If

    If datFim <> gstrDataVazia Then
        blnBaseA8 = datFim > datDataCorte
    End If

    strSqlA8 = ObterQueryDetalheOperacao(xmlFiltro, pblnAutomatico)
    strSQLHist = Replace$(strSqlA8, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
    strSQLHist = Replace$(strSQLHist, "'A8' AS OWNER", "'A8HIST' AS OWNER ")

    If blnBaseHistorica And Not blnBaseA8 Then
        strSQL = strSQLHist
    ElseIf blnBaseHistorica And blnBaseA8 Then
        strSQL = strSQLHist & _
                 " UNION ALL " & _
                 strSqlA8
    Else
        strSQL = strSqlA8
    End If
 
    ObterDetalheOperacao = fgQueryXMLLerTodos("DetalheOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    'ObterDetalheOperacaoXX(xmlFiltro, datinicio, datFim, Owner)

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "ObterDetalheOperacao", 0

End Function

'-------------------------------------------------------------------------------------------------
'Obter as operações conforme filtro especificado
'-------------------------------------------------------------------------------------------------

Public Function ObterDadosConsultaOperacao(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                                  Optional ByVal pblnAutomatico As Boolean = False) As String

Dim datInicio                               As Date
Dim datFim                                  As Date
Dim datDataCorte                            As Date

Dim blnChecaBaseHistórica                   As Boolean

Dim blnBaseA8                               As Boolean
Dim blnBaseHistorica                        As Boolean

Dim strSqlA8                                As String
Dim strSQLHist                              As String
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    datInicio = gstrDataVazia
    datFim = gstrDataVazia
    datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo

    blnChecaBaseHistórica = Not xmlFiltro.selectSingleNode("/Repeat_Filtros/BaseHistorica") Is Nothing

    If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data") Is Nothing Then
    
        If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
            datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
        End If

        If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
            datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
        End If

    End If

    If datInicio <> gstrDataVazia Then
        blnBaseHistorica = datInicio < datDataCorte And blnChecaBaseHistórica
    End If

    If datFim <> gstrDataVazia Then
        blnBaseA8 = datFim > datDataCorte
    End If

    strSqlA8 = ObterQueryConsultaOperacao(xmlFiltro, pblnAutomatico)
    strSQLHist = Replace$(strSqlA8, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
    strSQLHist = Replace$(strSQLHist, "'A8' AS OWNER", "'A8HIST' AS OWNER ")

    If blnBaseHistorica And Not blnBaseA8 Then
        strSQL = strSQLHist
    ElseIf blnBaseHistorica And blnBaseA8 Then
        strSQL = strSQLHist & _
                 " UNION ALL " & _
                 strSqlA8
    Else
        strSQL = strSqlA8
    End If
 
    ObterDadosConsultaOperacao = fgQueryXMLLerTodos("DetalheOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "ObterDadosConsultaOperacao", 0

End Function

'-------------------------------------------------------------------------------------------------
'Montar dinamicamente o select de operações conforme filtro especificado
'-------------------------------------------------------------------------------------------------

Private Function ObterQueryDetalheOperacao(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                                  Optional ByVal pblnAutomatico As Boolean = False) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomAttrs                             As MSXML2.IXMLDOMNamedNodeMap

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strOperador                             As String
Dim lngContador                             As Long
Dim blnOperacaoCetipConciliacao             As Boolean
Dim intIndicDebCred                         As enumTipoDebitoCredito
Dim intTipoLiquidacao                       As enumTipoLiquidacao
Dim blnLiquidacaoCorretora                  As Boolean

Dim blnSegregaBackOffice                    As Boolean
Dim intTipoBackOffice                       As Integer
                                                                
    On Error GoTo ErrorHandler
    
    intTipoBackOffice = 99
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(objDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            blnSegregaBackOffice = True
            If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SegregaBackOffice/*").length <> 0 Then
                Set objDomNode = xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_SegregaBackOffice/*")
                blnSegregaBackOffice = CBool(objDomNode.Text)
            End If
        
            If pblnAutomatico = False Then
                strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "B", False, False, blnSegregaBackOffice, True, True, False)
            End If
        Case 0
            strWhereSegregacao = vbNullString
        Case Is > 0
            strWhereSegregacao = vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    'Captura Filtro ------------------------------------------------------------------------------------------
    
    'Filtro por Canal Operação Interna
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalOperacaoInternaOperacao/*")
        If InStr(1, objDomNode.Text, "CONDIÇÃO") = 0 Then
            strValores = strValores & "'" & objDomNode.Text & "', "
        Else
            strValores = strValores & objDomNode.Text
            Exit For
        End If
    Next
    
    If Trim(strValores) <> "" Then
        If InStr(1, strValores, "CONDIÇÃO") = 0 Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.CO_CNAL_OPER_INTE  IN (" & strValores & ")"
        Else
            strValores = Replace$(strValores, "CONDIÇÃO", vbNullString)
            strWhereFiltro = strWhereFiltro & " AND " & strValores
        End If
    End If
    
    'Filtro por Numero Operacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_SEQU_OPER_ATIV  IN (" & strValores & ")"
    End If
    
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Status - 2ª forma de utilização
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SituacaoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Tipo Ação
    'Atributo utilizado devido a uma necessidade da tela frmLiberacaoOperacaoMensagem,
    'que seleciona todas as operações com Situação 4,5 ou 23 OU Ação 3 ou 5
    If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_TipoAcao/@Operador") Is Nothing Then
        strOperador = "OR"
        strWhereFiltro = Mid$(strWhereFiltro, 8)
        strWhereFiltro = " AND (" & strWhereFiltro
    Else
        strOperador = "AND"
    End If
    strValores = vbNullString

    'Carlos 31/05/2004
    'Tem dois pois tem telas que utilizam de maneiras diferentes
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoAcao/*")
        If objDomNode.Text = "0" Then
            strWhereFiltro = strWhereFiltro & " " & strOperador & " A.TP_ACAO_OPER_ATIV_EXEC IS NULL " & vbCrLf
        Else
            strValores = strValores & objDomNode.Text & ", "
        End If
    Next

    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Acoes/*")
        If objDomNode.Text = "0" Then
            strWhereFiltro = strWhereFiltro & " " & strOperador & " A.TP_ACAO_OPER_ATIV_EXEC IS NULL " & vbCrLf
        Else
            strValores = strValores & objDomNode.Text & ", "
        End If
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " " & strOperador & " A.TP_ACAO_OPER_ATIV_EXEC      IN (" & strValores & ") "
    End If
    strWhereFiltro = IIf(strOperador = "OR", strWhereFiltro & " ) " & vbCrLf, strWhereFiltro & vbCrLf)

    'Filtro por Tipo Operação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER       IN (" & strValores & ")"
    End If

    'Filtro por Tipo Operação (Excludente = NOT IN)
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoExc/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER   NOT IN (" & strValores & ")"
    End If

    'Tipo de Liquidacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
        intTipoLiquidacao = Val(objDomNode.Text)
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_LIQU_OPER_ATIV       IN (" & strValores & ")"
    End If

    'Filtro Tipo de Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoContraparte/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNPT      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro por Data Liquidacao
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataLiquidacao")
        strWhereFiltro = strWhereFiltro & "AND A.DT_LIQU_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                         objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode

    'Filtro por LOCAL LIQUIDACAO / CONJUNTO DE TIPOS DE OPERACAO
    lngContador = 0
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacaoTipoOperacao/*")
        lngContador = lngContador + 1
        If lngContador = 1 Then
            strWhereFiltro = strWhereFiltro & vbNewLine & "AND ( "
        Else
            strWhereFiltro = strWhereFiltro & vbNewLine & " OR "
        End If
        Set objDomAttrs = objDomNode.attributes
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " (A.CO_LOCA_LIQU = " & IIf(objDomAttrs.getNamedItem("Local").Text = "", _
                                                    "A.CO_LOCA_LIQU", _
                                                    objDomAttrs.getNamedItem("Local").Text) & "" & _
                     " AND A.TP_OPER IN (" & IIf(objDomAttrs.getNamedItem("Tipos").Text = "", _
                                                    "A.TP_OPER ", _
                                                    objDomAttrs.getNamedItem("Tipos").Text) & ") )"
    Next
    If lngContador > 0 Then
        strWhereFiltro = strWhereFiltro & ")"
    End If

    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      B.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Veiculo Legal / Sistema
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If
    
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.SG_SIST      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Nome Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Contraparte/*")
        strValores = objDomNode.Text
    Next
    
    If Trim(strValores) <> "" Then
        strWhereFiltro = strWhereFiltro & " AND      UPPER(A.NO_CNPT)     LIKE '" & strValores & "%'" & vbCrLf
    End If
    
    'Filtro CNPJ Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CNPJContraparte/*")
        strValores = objDomNode.Text
    Next
    
    If Trim(strValores) <> "" Then
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNPJ_CNPT      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro ISPB Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ISPBContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If strValores = "''" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.CO_ISPB_BANC_LIQU_CNPT      IN ('') " & vbCrLf
            strWhereFiltro = strWhereFiltro & " OR       A.CO_ISPB_BANC_LIQU_CNPT      IS NULL)" & vbCrLf
        ElseIf strValores = "'NotNull'" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.CO_ISPB_BANC_LIQU_CNPT  NOT IN ('') " & vbCrLf
            strWhereFiltro = strWhereFiltro & " OR       A.CO_ISPB_BANC_LIQU_CNPT  IS NOT NULL)" & vbCrLf
        Else
            strWhereFiltro = strWhereFiltro & " AND      A.CO_ISPB_BANC_LIQU_CNPT      IN (" & strValores & ")" & vbCrLf
        End If
    End If
    
    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Número do Comando
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroComando/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_COMD_OPER      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Número do Contrato SISBACEN
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ContratacaoSISBACEN/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNTR_SISB     IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro ChACAM
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ChACAM/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CD_ASSO_CAMB     IN ('" & strValores & "')" & vbCrLf
    End If

    'Filtro ChACAM ou Número do Contrato SISBACEN
    If Not xmlFiltro.selectSingleNode("//Grupo_ChACAMContratacaoSISBACEN") Is Nothing Then
        strWhereFiltro = strWhereFiltro & " AND     (A.CO_CNTR_SISB     IN (" & xmlFiltro.selectSingleNode("//CntrSISBACEN").Text & ")" & vbCrLf
        strWhereFiltro = strWhereFiltro & " OR       A.NU_COMD_OPER     IN (" & xmlFiltro.selectSingleNode("//CntrSISBACEN").Text & ")" & vbCrLf
        strWhereFiltro = strWhereFiltro & " OR       A.CD_ASSO_CAMB     IN ('" & xmlFiltro.selectSingleNode("//ChACAM").Text & "'))" & vbCrLf
    End If

    'Filtro Numero Identificacao Negociacao BMC
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentNegocioBMC/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & " AND      A.NR_IDEF_NEGO_BMC     IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Canal SISBACEN Corretora
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalSISBACENCorretora/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_SISB_COTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Código Praça
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoPraca/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_PRAC     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Código Moeda Estrangeira
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoMoedaEstrangeira/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_MOED_ESTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Valor Moeda Estrangeira
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ValorMoedaEstrangeira/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.VA_MOED_ESTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Taxa Negociação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TaxaNegociacao/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.PE_TAXA_NEGO     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Valor Financeiro
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ValorFinanceiro/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.VA_OPER_ATIV     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Data Liquidação Moeda Estrangeira
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataLiquidacaoMoedaEstr").length <> 0 Then
        Set objDomNode = xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataLiquidacaoMoedaEstr").Item(0)

        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.DT_LIQU_OPER_ATIV_MOED_ESTR      = " & objDomNode.Text & ""
    End If

    'Filtro Título
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Titulo/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_ATIV_MERC      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por CNPJ do Veículo Legal da Operacao (não é da contraparte)
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CNPJ/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      B.CO_CNPJ_VEIC_LEGA  IN (" & strValores & ")"
    End If
    
    'Filtro Conta SELIC
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ContaSELIC/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNTA_CUTD_SELIC_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If
    
    blnLiquidacaoCorretora = False
    
    'Filtro Layout de Entrada da Operacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LayoutEntrada/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
        
        'RATS 1212 - Operação Corretora
        If objDomNode.Text = enumTipoMensagemBUS.OperacoesCorretoras Then
            blnLiquidacaoCorretora = True
        End If
        
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      D.TP_MESG_RECB_INTE      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Data de Vencimento
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataVencimento")
        strWhereFiltro = strWhereFiltro & "AND A.DT_VENC_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                          objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode
    
    'Filtro Indicador de entrada manual
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_EntradaManual/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.IN_ENTR_MANU      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Classificacao do Veiculo Legal na BMA
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoTitularBMA/*")
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      B.TP_TITL_BMA "
                        
        'Monta o OPERADOR sql de acordo com o Nome da Tag que for passada
        strWhereFiltro = strWhereFiltro & fgDECODE(LCase(objDomNode.baseName), _
                                                        "igual", "=", _
                                                        "menor", "<", _
                                                        "maior", ">", _
                                                        "menorouigual", "<=", _
                                                        "maiorouigual", ">=", _
                                                        "diferente", "<>" _
                                                        )
        
        strWhereFiltro = strWhereFiltro & "'" & objDomNode.Text & "'"
    Next

    'Filtro Indicador Contingencia de sistema
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SituacaoContingencia/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      F.IN_SIST_SITU_CNTG      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Tipo Operação Rotina Abertura
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoRotinaAbertura/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      D.CO_OPER_SELIC      IS NOT NULL" & vbCrLf
    End If
    
    'Filtro Número Controle LTR
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroControleLTR/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If UCase$(strValores) = "'NULL'" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.NU_CTRL_MESG_SPB_ORIG      IN ('')"
            strWhereFiltro = strWhereFiltro & " OR       A.NU_CTRL_MESG_SPB_ORIG      IS NULL)"
        Else
            strWhereFiltro = strWhereFiltro & " AND      A.NU_CTRL_MESG_SPB_ORIG      IN (" & strValores & ")"
        End If
    End If
    
    'Filtro Identificador Participante Câmara
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentificadorPartCamara/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_PARP_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro Identificador Participante Câmara Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentificadorPartCamaraContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNPT_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro Grupo Lançamento Financeiro
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoLancamentoFinanceiro/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_GRUP_LANC_FINC      IN (" & strValores & ")"
    End If
    
    'Filtro da Ação da mensagem SPB
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_AcaoMensagemSPB/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND  A.TP_OPER IN( SELECT TP_OPER FROM A8.TB_TIPO_OPER_ACAO_MESG_SPB WHERE TP_ACAO_OPER_ATIV_EXEC IN ( " & strValores & " ) )"
    End If
    
    'Filtro por Data
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                  objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_ATIV) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_Data/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro por Data - 2ª forma de utilização
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                  objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_ATIV) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_DataOperacao/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro por Data de Retorno
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataRetorno").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataRetorno/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataRetorno")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV_RETN between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                       objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_ATIV_RETN) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_DataRetorno/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro para trazer somente 'OPERACOES CETIP COM CONCILIACAO'
    If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_OperacaoCetipConciliacao") Is Nothing Then
        blnOperacaoCetipConciliacao = True
    End If

    'Filtro por Indicador de Débito e Crédito
    If intTipoLiquidacao = enumTipoLiquidacao.Bruta Then
        strValores = vbNullString
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NaturezaMovimento/*")
            intIndicDebCred = Val(objDomNode.Text)

            intIndicDebCred = IIf(intIndicDebCred = enumTipoDebitoCredito.Debito, _
                                                    enumTipoDebitoCredito.Credito, _
                                                    enumTipoDebitoCredito.Debito)
            
            strValores = strValores & intIndicDebCred & ", "
        Next
        
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.IN_OPER_DEBT_CRED  IN (" & strValores & ")"
        End If
    End If
        
    'KIDA - SGC
    'Filtro Grupo Canal Venda
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
    End If
    
    
    '---------------------------------------------------------------------------------------------------------

    strSQL = " SELECT   'A8' AS OWNER,                                                     " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                                               " & vbNewLine & _
             "          A.TP_OPER,                                                         " & vbNewLine & _
             "          A.DT_OPER_ATIV,                                                    " & vbNewLine & _
             "          A.DT_OPER_ATIV_RETN,                                               " & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          A.NU_COMD_OPER_RETN,                                               " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          A.SG_SIST,                                                         " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          B.CO_CNPJ_VEIC_LEGA,                                               " & vbNewLine & _
             "          B.TP_BKOF,                                                         " & vbNewLine & _
             "          G.DE_BKOF,                                                         " & vbNewLine & _
             "          A.TP_LIQU_OPER_ATIV,                                               " & vbNewLine & _
             "          E.NO_TIPO_LIQU_OPER_ATIV,                                          " & vbNewLine & _
             "          A.CO_CNPJ_CNPT,                                                    " & vbNewLine & _
             "          A.NO_CNPT,                                                         " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          A.TP_ACAO_OPER_ATIV_EXEC,                                          " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                    " & vbNewLine & _
             "          A.IN_OPER_DEBT_CRED                 AS  CAMPO_IN_OPER_DEBT_CRED,   " & vbNewLine

    strSQL = strSQL & _
             "          DECODE(A.IN_OPER_DEBT_CRED,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito')    " & _
                                              "IN_OPER_DEBT_CRED, " & vbNewLine & _
             "          DECODE(A.IN_ENTR_SAID_RECU_FINC,                                   " & vbNewLine & _
             "                 " & enumTipoEntradaSaida.ENTRADA & ", 'Entrada', " & _
                                   enumTipoEntradaSaida.Saida & ", 'Saída')   " & _
                                   "IN_ENTR_SAID_RECU_FINC,                                " & vbNewLine & _
             "          A.NU_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.DE_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.PU_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.QT_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.VA_OPER_ATIV,                                                    " & vbNewLine & _
             "          A.VA_OPER_ATIV        AS VA_OPER_ATIV_VLRXML /* para ficar com '.' separando os decimais, no xml */," & vbNewLine & _
             "          A.VA_OPER_ATIV_REAJ,                                               " & vbNewLine & _
             "          A.VA_OPER_ATIV_REAJ   AS VA_OPER_ATIV_REAJ_VLRXML,                 " & vbNewLine & _
             "          DECODE(A.VA_OPER_ATIV_REAJ, NULL, A.VA_OPER_ATIV, A.VA_OPER_ATIV_REAJ) AS VA_OPER_ATIV_ORIG," & vbNewLine & _
             "          DECODE(A.VA_OPER_ATIV_REAJ, NULL, A.VA_OPER_ATIV, A.VA_OPER_ATIV_REAJ) AS VA_OPER_ATIV_ORIG_VLRXML," & vbNewLine & _
             "          A.CO_TITL_CUTD,                                                    " & vbNewLine
             
    strSQL = strSQL & _
             "          A.DT_VENC_ATIV,                                                    " & vbNewLine & _
             "          A.PE_TAXA_NEGO,                                                    " & vbNewLine & _
             "          A.CO_OPER_ATIV,                                                    " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
             "          A.DT_LIQU_OPER_ATIV,                                               " & vbNewLine & _
             "          A.TP_CPRO_OPER_ATIV,                                               " & vbNewLine & _
             "          A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.HO_ENVI_MESG_SPB,                                                " & vbNewLine & _
             "          A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,                                    " & vbNewLine & _
             "          A.CO_CNTA_CUTD_SELIC_CNPT,                                         " & vbNewLine & _
             "          A.CO_OPER_CETIP,                                                   " & vbNewLine
             
    strSQL = strSQL & _
             "          DECODE(A.TP_OPER,                                                  " & vbNewLine & _
             "                 " & enumTipoOperacaoLQS.Vinculo & " , A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,  " & vbNewLine & _
             "                 " & enumTipoOperacaoLQS.Desvinculo & " , A.CO_CNTA_CUTD_SELIC_CNPT,    " & vbNewLine & _
             "                 DECODE(A.IN_ENTR_SAID_RECU_FINC,                                   " & vbNewLine & _
             "                       " & enumTipoEntradaSaida.ENTRADA & ", A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbNewLine & _
                                         enumIndicadorEntradaSaida.Saida & ", A.CO_CNTA_CUTD_SELIC_CNPT)) " & _
                                "CO_CNTA_CEDT, " & vbNewLine & _
             "          DECODE(A.TP_OPER,                                                  " & vbNewLine & _
             "              " & enumTipoOperacaoLQS.Vinculo & " , A.CO_CNTA_CUTD_SELIC_CNPT, " & vbNewLine & _
             "              " & enumTipoOperacaoLQS.Desvinculo & " , A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbNewLine & _
             "              DECODE(A.IN_ENTR_SAID_RECU_FINC,                                   " & vbNewLine & _
             "                    " & enumTipoEntradaSaida.ENTRADA & ", A.CO_CNTA_CUTD_SELIC_CNPT, " & vbNewLine & _
                                    enumIndicadorEntradaSaida.Saida & ", A.CO_CNTA_CUTD_SELIC_VEIC_LEGA) )" & _
                                "CO_CNTA_CESS, " & vbNewLine
                                
    strSQL = strSQL & _
             "          D.TP_MESG_RECB_INTE,                                               " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.TP_CPRO_RETN_OPER_ATIV,                                          " & vbNewLine & _
             "          A.NU_COMD_ACAO_EXEC,                                               " & vbNewLine & _
             "          A.NU_PRTC_OPER_LG,                                                 " & vbNewLine & _
             "          A.CO_ISPB_BANC_LIQU_CNPT,                                          " & vbNewLine & _
             "          H.SG_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.TP_IF_CRED_DEBT,                                                 " & vbNewLine & _
             "          DECODE(A.TP_IF_CRED_DEBT,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito')    " & _
                                              "DE_IF_CRED_DEBT, " & vbNewLine & _
             "          A.CO_AGEN_COTR,                                                    " & vbNewLine & _
             "          A.NU_CC_COTR,                                                      " & vbNewLine & _
             "          A.PZ_DIAS_RETN_OPER_ATIV,                                          " & vbNewLine & _
             "          A.VA_OPER_ATIV_RETN,                                               " & vbNewLine & _
             "          A.VA_OPER_ATIV_RETN         AS VA_OPER_ATIV_RETN_VLRXML,           " & vbNewLine & _
             "          A.TP_CNPT,                                                         " & vbNewLine & _
             "          A.NU_CTRL_MESG_SPB_ORIG,                                           " & vbNewLine & _
             "          B.ID_PART_CAMR_CETIP,                                              " & vbNewLine & _
             "          A.CO_CNPT_CAMR,                                                    " & vbNewLine

    strSQL = strSQL & _
             "          A.TP_PGTO_LDL,                                                     " & vbNewLine & _
             "          A.CO_PARP_CAMR,                                                    " & vbNewLine & _
             "          A.CO_GRUP_LANC_FINC,                                               " & vbNewLine & _
             "          I.DE_GRUP_LANC_FINC,                                               " & vbNewLine & _
             "          J.DS_PARP_CAMR,                                                    " & vbNewLine & _
             "          A.CO_MOED_ESTR,                                                    " & vbNewLine & _
             "          A.CO_CNTR_SISB,                                                    " & vbNewLine & _
             "          A.CO_ISPB_IF_CNPT,                                                 " & vbNewLine & _
             "          A.CO_PRAC,                                                         " & vbNewLine & _
             "          A.VA_MOED_ESTR,                                                    " & vbNewLine & _
             "          A.VA_MOED_ESTR             AS VA_MOED_ESTR_VLRXML,                 " & vbNewLine & _
             "          A.DT_LIQU_OPER_ATIV_MOED_ESTR,                                     " & vbNewLine & _
             "          A.CO_SISB_COTR,                                                    " & vbNewLine & _
             "          A.TP_CNAL_VEND,                                                    " & vbNewLine & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,                                     " & vbNewLine & _
             "          A.CO_IDEF_LAST,                                                    " & vbNewLine & _
             "          A.CD_OPER_ETRT,                                                    " & vbNewLine & _
             "          A.NR_IDEF_NEGO_BMC,                                                " & vbNewLine & _
             "          A.TP_NEGO,                                                         " & vbNewLine & _
             "          A.CD_ASSO_CAMB,                                                    " & vbNewLine & _
             "          A.NR_CNPJ_CPF,                                                     " & vbNewLine & _
             "          A.CO_USUA_CADR_OPER,                                               " & vbNewLine & _
             "          A.NR_OPER_CAMB_2,                                                  " & vbNewLine & _
             "          A.CD_LOTE,                                                         " & vbNewLine
    strSQL = strSQL & _
             "          DECODE(L.TP_DEBT_CRED_LOTE,                                         " & vbNewLine & _
             "                   NULL, '', 1, 'Débito', 2, 'Crédito')    TP_DEBT_CRED_LOTE," & vbNewLine & _
             "          L.VL_TOTL_LOTE,                                                    " & vbNewLine & _
             "          L.QT_TOTL_LOTE,                                                    " & vbNewLine & _
             "          L.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          L.IN_LIQU_LOTE                                                     " & vbNewLine
             

'ESTÁ IMPLEMENTAÇÃO ESTÁ EM STAND-BY, AGUARDANDO PRIORIZAÇÃO PARA SER IMPLANTADA
'campos incluídos por solicitação dos usuário do Comex, devido projeto Sisbacen
'    strSQL = strSQL & _
'             "          A.NO_CLIE,                                                         " & vbNewLine & _
'             "          A.CD_MOED_ISO,                                                     " & vbNewLine & _
'             "          A.VA_MOED_ESTR,                                                    " & vbNewLine & _
'             "          A.NR_PERC_TAXA_CAMB,                                               " & vbNewLine & _
'             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
'             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
'             "          A.TP_OPER_CAMB                                                     " & vbNewLine

    strSQL = strSQL & _
             " FROM     " & vbNewLine
    
    If blnOperacaoCetipConciliacao Then
        strSQL = strSQL & _
             "          (SELECT DISTINCT TP_OPER FROM A8.TB_TIPO_OPER_CNCL_CETIP)   OCC,   " & vbNewLine
    End If
    
    strSQL = strSQL & _
             "          A8.TB_OPER_ATIV                 A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C,                                 " & vbNewLine & _
             "          A8.TB_TIPO_OPER                 D,                                 " & vbNewLine & _
             "          A8.TB_TIPO_LIQU_OPER_ATIV       E,                                 " & vbNewLine & _
             "          A8.TB_SITU_CNTG_SIST            F,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 G,                                 " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO          H,                                 " & vbNewLine & _
             "          A8.TB_GRUP_LANC_FINC            I,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF_PARP_CAMR       J,                                 " & vbNewLine & _
             "          A8.TB_LOTE_OPER_ATIV            L                                  " & vbNewLine

    strSQL = strSQL & _
             " WHERE    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                           " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                " & vbNewLine & _
             " AND      B.TP_BKOF               = G.TP_BKOF             (+)                " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                           " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER                                " & vbNewLine & _
             " AND      A.TP_LIQU_OPER_ATIV     = E.TP_LIQU_OPER_ATIV   (+)                " & vbNewLine & _
             " AND      A.IN_DISP_CONS          = " & enumIndicadorSimNao.Sim & "          " & vbNewLine & _
             " AND      A.SG_SIST               = F.SG_SIST             (+)                " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU          = H.CO_LOCA_LIQU                           " & vbNewLine & _
             " AND      H.CO_EMPR_FUSI          = 1                                        " & vbNewLine & _
             " AND      A.CO_GRUP_LANC_FINC     = I.CO_GRUP_LANC_FINC   (+)                " & vbNewLine & _
             " AND      A.CO_PARP_CAMR          = J.CO_PARP_CAMR        (+)                " & vbNewLine & _
             " AND      A.CD_LOTE               = L.CD_LOTE             (+)                " & vbNewLine & _
             " AND      A.SG_SIST               = L.SG_SIST             (+)                " & vbNewLine & _
             " AND      A.DT_OPER_ATIV          = L.DT_OPER_ATIV        (+)                " & vbNewLine

             
    If blnOperacaoCetipConciliacao Then
        strSQL = strSQL & _
             " AND      OCC.TP_OPER             = A.TP_OPER                                " & vbNewLine
    End If
    
    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If
    
    'Ordernação customizada pelo chamador
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_OrderBy/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strSQL = strSQL & vbNewLine & " ORDER BY " & strValores
    End If

    ObterQueryDetalheOperacao = strSQL
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQueryDetalheOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    Exit Function
    Resume

End Function

'-------------------------------------------------------------------------------------------------
'Montar dinamicamente o select de operações conforme filtro especificado
'-------------------------------------------------------------------------------------------------

Private Function ObterQueryConsultaOperacao(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                                  Optional ByVal pblnAutomatico As Boolean = False) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomAttrs                             As MSXML2.IXMLDOMNamedNodeMap

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strOperador                             As String
Dim lngContador                             As Long
Dim blnOperacaoCetipConciliacao             As Boolean
Dim intIndicDebCred                         As enumTipoDebitoCredito
Dim intTipoLiquidacao                       As enumTipoLiquidacao

Dim blnSegregaBackOffice                    As Boolean
Dim intTipoBackOffice                       As Integer
                                                                
    On Error GoTo ErrorHandler
    
    intTipoBackOffice = 99
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(objDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            blnSegregaBackOffice = True
            If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SegregaBackOffice/*").length <> 0 Then
                Set objDomNode = xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_SegregaBackOffice/*")
                blnSegregaBackOffice = CBool(objDomNode.Text)
            End If
        
            If pblnAutomatico = False Then
                strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "B", False, False, blnSegregaBackOffice, True, True, False)
            End If
        Case 0
            strWhereSegregacao = vbNullString
        Case Is > 0
            strWhereSegregacao = vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    'Captura Filtro ------------------------------------------------------------------------------------------
    
    'Filtro por Canal Operação Interna
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalOperacaoInternaOperacao/*")
        If InStr(1, objDomNode.Text, "CONDIÇÃO") = 0 Then
            strValores = strValores & "'" & objDomNode.Text & "', "
        Else
            strValores = strValores & objDomNode.Text
            Exit For
        End If
    Next
    
    If Trim(strValores) <> "" Then
        If InStr(1, strValores, "CONDIÇÃO") = 0 Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.CO_CNAL_OPER_INTE  IN (" & strValores & ")"
        Else
            strValores = Replace$(strValores, "CONDIÇÃO", vbNullString)
            strWhereFiltro = strWhereFiltro & " AND " & strValores
        End If
    End If
    
    'Filtro por Numero Operacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_SEQU_OPER_ATIV  IN (" & strValores & ")"
    End If
    
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Status - 2ª forma de utilização
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SituacaoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Tipo Ação
    'Atributo utilizado devido a uma necessidade da tela frmLiberacaoOperacaoMensagem,
    'que seleciona todas as operações com Situação 4,5 ou 23 OU Ação 3 ou 5
    If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_TipoAcao/@Operador") Is Nothing Then
        strOperador = "OR"
        strWhereFiltro = Mid$(strWhereFiltro, 8)
        strWhereFiltro = " AND (" & strWhereFiltro
    Else
        strOperador = "AND"
    End If
    strValores = vbNullString

    'Carlos 31/05/2004
    'Tem dois pois tem telas que utilizam de maneiras diferentes
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoAcao/*")
        If objDomNode.Text = "0" Then
            strWhereFiltro = strWhereFiltro & " " & strOperador & " A.TP_ACAO_OPER_ATIV_EXEC IS NULL " & vbCrLf
        Else
            strValores = strValores & objDomNode.Text & ", "
        End If
    Next

    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Acoes/*")
        If objDomNode.Text = "0" Then
            strWhereFiltro = strWhereFiltro & " " & strOperador & " A.TP_ACAO_OPER_ATIV_EXEC IS NULL " & vbCrLf
        Else
            strValores = strValores & objDomNode.Text & ", "
        End If
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " " & strOperador & " A.TP_ACAO_OPER_ATIV_EXEC      IN (" & strValores & ") "
    End If
    strWhereFiltro = IIf(strOperador = "OR", strWhereFiltro & " ) " & vbCrLf, strWhereFiltro & vbCrLf)

    'Filtro por Tipo Operação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER       IN (" & strValores & ")"
    End If

    'Filtro por Tipo Operação (Excludente = NOT IN)
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoExc/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER   NOT IN (" & strValores & ")"
    End If

    'Tipo de Liquidacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
        intTipoLiquidacao = Val(objDomNode.Text)
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_LIQU_OPER_ATIV       IN (" & strValores & ")"
    End If

    'Filtro Tipo de Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoContraparte/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNPT      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro por Data Liquidacao
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataLiquidacao")
        strWhereFiltro = strWhereFiltro & "AND A.DT_LIQU_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                         objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode

    'Filtro por LOCAL LIQUIDACAO / CONJUNTO DE TIPOS DE OPERACAO
    lngContador = 0
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacaoTipoOperacao/*")
        lngContador = lngContador + 1
        If lngContador = 1 Then
            strWhereFiltro = strWhereFiltro & vbNewLine & "AND ( "
        Else
            strWhereFiltro = strWhereFiltro & vbNewLine & " OR "
        End If
        Set objDomAttrs = objDomNode.attributes
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " (A.CO_LOCA_LIQU = " & IIf(objDomAttrs.getNamedItem("Local").Text = "", _
                                                    "A.CO_LOCA_LIQU", _
                                                    objDomAttrs.getNamedItem("Local").Text) & "" & _
                     " AND A.TP_OPER IN (" & IIf(objDomAttrs.getNamedItem("Tipos").Text = "", _
                                                    "A.TP_OPER ", _
                                                    objDomAttrs.getNamedItem("Tipos").Text) & ") )"
    Next
    If lngContador > 0 Then
        strWhereFiltro = strWhereFiltro & ")"
    End If

    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      B.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Veiculo Legal / Sistema
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If
    
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.SG_SIST      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Nome Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Contraparte/*")
        strValores = objDomNode.Text
    Next
    
    If Trim(strValores) <> "" Then
        strWhereFiltro = strWhereFiltro & " AND      UPPER(A.NO_CNPT)     LIKE '" & strValores & "%'" & vbCrLf
    End If
    
    'Filtro CNPJ Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CNPJContraparte/*")
        strValores = objDomNode.Text
    Next
    
    If Trim(strValores) <> "" Then
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNPJ_CNPT      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro ISPB Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ISPBContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If strValores = "''" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.CO_ISPB_BANC_LIQU_CNPT      IN ('') " & vbCrLf
            strWhereFiltro = strWhereFiltro & " OR       A.CO_ISPB_BANC_LIQU_CNPT      IS NULL)" & vbCrLf
        ElseIf strValores = "'NotNull'" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.CO_ISPB_BANC_LIQU_CNPT  NOT IN ('') " & vbCrLf
            strWhereFiltro = strWhereFiltro & " OR       A.CO_ISPB_BANC_LIQU_CNPT  IS NOT NULL)" & vbCrLf
        Else
            strWhereFiltro = strWhereFiltro & " AND      A.CO_ISPB_BANC_LIQU_CNPT      IN (" & strValores & ")" & vbCrLf
        End If
    End If
    
    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Número do Comando
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroComando/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_COMD_OPER      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Número do Contrato SISBACEN
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ContratacaoSISBACEN/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNTR_SISB     IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro ChACAM
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ChACAM/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CD_ASSO_CAMB     IN ('" & strValores & "')" & vbCrLf
    End If

    'Filtro ChACAM ou Número do Contrato SISBACEN
    If Not xmlFiltro.selectSingleNode("//Grupo_ChACAMContratacaoSISBACEN") Is Nothing Then
        strWhereFiltro = strWhereFiltro & " AND     (A.CO_CNTR_SISB     IN (" & xmlFiltro.selectSingleNode("//CntrSISBACEN").Text & ")" & vbCrLf
        strWhereFiltro = strWhereFiltro & " OR       A.CD_ASSO_CAMB     IN ('" & xmlFiltro.selectSingleNode("//ChACAM").Text & "'))" & vbCrLf
    End If

    'Filtro Numero Identificacao Negociacao BMC
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentNegocioBMC/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & " AND      A.NR_IDEF_NEGO_BMC     IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Canal SISBACEN Corretora
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalSISBACENCorretora/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_SISB_COTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Código Praça
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoPraca/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_PRAC     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Código Moeda Estrangeira
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoMoedaEstrangeira/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_MOED_ESTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Valor Moeda Estrangeira
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ValorMoedaEstrangeira/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.VA_MOED_ESTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Taxa Negociação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TaxaNegociacao/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.PE_TAXA_NEGO     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Valor Financeiro
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ValorFinanceiro/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.VA_OPER_ATIV     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Data Liquidação Moeda Estrangeira
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataLiquidacaoMoedaEstr").length <> 0 Then
        Set objDomNode = xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataLiquidacaoMoedaEstr").Item(0)

        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.DT_LIQU_OPER_ATIV_MOED_ESTR      = " & objDomNode.Text & ""
    End If

    'Filtro Título
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Titulo/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_ATIV_MERC      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por CNPJ do Veículo Legal da Operacao (não é da contraparte)
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CNPJ/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      B.CO_CNPJ_VEIC_LEGA  IN (" & strValores & ")"
    End If
    
    'Filtro Conta SELIC
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ContaSELIC/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNTA_CUTD_SELIC_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Layout de Entrada da Operacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LayoutEntrada/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      D.TP_MESG_RECB_INTE      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Data de Vencimento
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataVencimento")
        strWhereFiltro = strWhereFiltro & "AND A.DT_VENC_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                          objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode
    
    'Filtro Indicador de entrada manual
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_EntradaManual/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.IN_ENTR_MANU      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Classificacao do Veiculo Legal na BMA
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoTitularBMA/*")
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      B.TP_TITL_BMA "
                        
        'Monta o OPERADOR sql de acordo com o Nome da Tag que for passada
        strWhereFiltro = strWhereFiltro & fgDECODE(LCase(objDomNode.baseName), _
                                                        "igual", "=", _
                                                        "menor", "<", _
                                                        "maior", ">", _
                                                        "menorouigual", "<=", _
                                                        "maiorouigual", ">=", _
                                                        "diferente", "<>" _
                                                        )
        
        strWhereFiltro = strWhereFiltro & "'" & objDomNode.Text & "'"
    Next

    'Filtro Indicador Contingencia de sistema
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SituacaoContingencia/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      F.IN_SIST_SITU_CNTG      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Tipo Operação Rotina Abertura
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoRotinaAbertura/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      D.CO_OPER_SELIC      IS NOT NULL" & vbCrLf
    End If
    
    'Filtro Número Controle LTR / Cod Reembolso CCR
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroControleLTR/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If UCase$(strValores) = "'NULL'" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.NU_CTRL_MESG_SPB_ORIG      IN ('')"
            strWhereFiltro = strWhereFiltro & " OR       A.NU_CTRL_MESG_SPB_ORIG      IS NULL)"
        Else
            strWhereFiltro = strWhereFiltro & " AND      A.NU_CTRL_MESG_SPB_ORIG      IN (" & strValores & ")"
        End If
    End If
    
    'Filtro Identificador Participante Câmara
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentificadorPartCamara/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_PARP_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro Identificador Participante Câmara Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentificadorPartCamaraContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNPT_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro Grupo Lançamento Financeiro
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoLancamentoFinanceiro/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_GRUP_LANC_FINC      IN (" & strValores & ")"
    End If
    
    'Filtro da Ação da mensagem SPB
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_AcaoMensagemSPB/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND  A.TP_OPER IN( SELECT TP_OPER FROM A8.TB_TIPO_OPER_ACAO_MESG_SPB WHERE TP_ACAO_OPER_ATIV_EXEC IN ( " & strValores & " ) )"
    End If
    
    'Filtro por Data
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                  objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_ATIV) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_Data/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro por Data - 2ª forma de utilização
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                  objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_ATIV) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_DataOperacao/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro por Data de Retorno
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataRetorno").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataRetorno/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataRetorno")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV_RETN between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                       objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_ATIV_RETN) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_DataRetorno/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro para trazer somente 'OPERACOES CETIP COM CONCILIACAO'
    If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_OperacaoCetipConciliacao") Is Nothing Then
        blnOperacaoCetipConciliacao = True
    End If

    'Filtro por Indicador de Débito e Crédito
    If intTipoLiquidacao = enumTipoLiquidacao.Bruta Then
        strValores = vbNullString
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NaturezaMovimento/*")
            intIndicDebCred = Val(objDomNode.Text)

            intIndicDebCred = IIf(intIndicDebCred = enumTipoDebitoCredito.Debito, _
                                                    enumTipoDebitoCredito.Credito, _
                                                    enumTipoDebitoCredito.Debito)
            
            strValores = strValores & intIndicDebCred & ", "
        Next
        
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.IN_OPER_DEBT_CRED  IN (" & strValores & ")"
        End If
    End If
        
    'KIDA - SGC
    'Filtro Grupo Canal Venda
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
    End If
    
    
    'KIDA - Paginacao -------------------------------------------------------------------
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SeqOperPaginacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND  A.NU_SEQU_OPER_ATIV > " & strValores
    End If
    
    '---------------------------------------------------------------------------------
    
    
    
    '---------------------------------------------------------------------------------------------------------

    strSQL = " SELECT   'A8' AS OWNER,                                                     " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                                               " & vbNewLine & _
             "          A.TP_OPER,                                                         " & vbNewLine & _
             "          A.DT_OPER_ATIV,                                                    " & vbNewLine & _
             "          A.DT_OPER_ATIV_RETN,                                               " & vbNewLine & _
             "          LTRIM(A.NU_COMD_OPER, 0) AS NU_COMD_OPER,                          " & vbNewLine & _
             "          LTRIM(A.NU_COMD_OPER_RETN, 0) AS NU_COMD_OPER_RETN,                " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          A.SG_SIST,                                                         " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          B.CO_CNPJ_VEIC_LEGA,                                               " & vbNewLine & _
             "          B.TP_BKOF,                                                         " & vbNewLine & _
             "          G.DE_BKOF,                                                         " & vbNewLine & _
             "          A.TP_LIQU_OPER_ATIV,                                               " & vbNewLine & _
             "          E.NO_TIPO_LIQU_OPER_ATIV,                                          " & vbNewLine & _
             "          A.CO_CNPJ_CNPT,                                                    " & vbNewLine & _
             "          A.NO_CNPT,                                                         " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          A.TP_ACAO_OPER_ATIV_EXEC,                                          " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                    " & vbNewLine & _
             "          A.IN_OPER_DEBT_CRED                 AS  CAMPO_IN_OPER_DEBT_CRED,   " & vbNewLine

    strSQL = strSQL & _
             "          DECODE(A.IN_OPER_DEBT_CRED,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito')    " & _
                                              "IN_OPER_DEBT_CRED, " & vbNewLine & _
             "          DECODE(A.IN_ENTR_SAID_RECU_FINC,                                   " & vbNewLine & _
             "                 " & enumTipoEntradaSaida.ENTRADA & ", 'Entrada', " & _
                                   enumTipoEntradaSaida.Saida & ", 'Saída')   " & _
                                   "IN_ENTR_SAID_RECU_FINC,                                " & vbNewLine & _
             "          A.NU_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.DE_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.PU_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.QT_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.VA_OPER_ATIV,                                                    " & vbNewLine & _
             "          A.VA_OPER_ATIV        AS VA_OPER_ATIV_VLRXML /* para ficar com '.' separando os decimais, no xml */," & vbNewLine & _
             "          A.VA_OPER_ATIV_REAJ,                                               " & vbNewLine & _
             "          A.VA_OPER_ATIV_REAJ   AS VA_OPER_ATIV_REAJ_VLRXML,                 " & vbNewLine & _
             "          DECODE(A.VA_OPER_ATIV_REAJ, NULL, A.VA_OPER_ATIV, A.VA_OPER_ATIV_REAJ) AS VA_OPER_ATIV_ORIG," & vbNewLine & _
             "          DECODE(A.VA_OPER_ATIV_REAJ, NULL, A.VA_OPER_ATIV, A.VA_OPER_ATIV_REAJ) AS VA_OPER_ATIV_ORIG_VLRXML," & vbNewLine & _
             "          A.CO_TITL_CUTD,                                                    " & vbNewLine
             
    strSQL = strSQL & _
             "          A.DT_VENC_ATIV,                                                    " & vbNewLine & _
             "          A.PE_TAXA_NEGO,                                                    " & vbNewLine & _
             "          A.CO_OPER_ATIV,                                                    " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
             "          A.DT_LIQU_OPER_ATIV,                                               " & vbNewLine & _
             "          A.TP_CPRO_OPER_ATIV,                                               " & vbNewLine & _
             "          A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.HO_ENVI_MESG_SPB,                                                " & vbNewLine & _
             "          A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,                                    " & vbNewLine & _
             "          A.CO_CNTA_CUTD_SELIC_CNPT,                                         " & vbNewLine & _
             "          A.CO_OPER_CETIP,                                                   " & vbNewLine
             
    strSQL = strSQL & _
             "          DECODE(A.TP_OPER,                                                  " & vbNewLine & _
             "                 " & enumTipoOperacaoLQS.Vinculo & " , A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,  " & vbNewLine & _
             "                 " & enumTipoOperacaoLQS.Desvinculo & " , A.CO_CNTA_CUTD_SELIC_CNPT,    " & vbNewLine & _
             "                 DECODE(A.IN_ENTR_SAID_RECU_FINC,                                   " & vbNewLine & _
             "                       " & enumTipoEntradaSaida.ENTRADA & ", A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbNewLine & _
                                         enumIndicadorEntradaSaida.Saida & ", A.CO_CNTA_CUTD_SELIC_CNPT)) " & _
                                "CO_CNTA_CEDT, " & vbNewLine & _
             "          DECODE(A.TP_OPER,                                                  " & vbNewLine & _
             "              " & enumTipoOperacaoLQS.Vinculo & " , A.CO_CNTA_CUTD_SELIC_CNPT, " & vbNewLine & _
             "              " & enumTipoOperacaoLQS.Desvinculo & " , A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbNewLine & _
             "              DECODE(A.IN_ENTR_SAID_RECU_FINC,                                   " & vbNewLine & _
             "                    " & enumTipoEntradaSaida.ENTRADA & ", A.CO_CNTA_CUTD_SELIC_CNPT, " & vbNewLine & _
                                    enumIndicadorEntradaSaida.Saida & ", A.CO_CNTA_CUTD_SELIC_VEIC_LEGA) )" & _
                                "CO_CNTA_CESS, " & vbNewLine
                                
    strSQL = strSQL & _
             "          D.TP_MESG_RECB_INTE,                                               " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.TP_CPRO_RETN_OPER_ATIV,                                          " & vbNewLine & _
             "          A.NU_COMD_ACAO_EXEC,                                               " & vbNewLine & _
             "          A.NU_PRTC_OPER_LG,                                                 " & vbNewLine & _
             "          A.CO_ISPB_BANC_LIQU_CNPT,                                          " & vbNewLine & _
             "          H.SG_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.TP_IF_CRED_DEBT,                                                 " & vbNewLine & _
             "          DECODE(A.TP_IF_CRED_DEBT,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito')    " & _
                                              "DE_IF_CRED_DEBT, " & vbNewLine & _
             "          A.CO_AGEN_COTR,                                                    " & vbNewLine & _
             "          A.NU_CC_COTR,                                                      " & vbNewLine & _
             "          A.PZ_DIAS_RETN_OPER_ATIV,                                          " & vbNewLine & _
             "          A.VA_OPER_ATIV_RETN,                                               " & vbNewLine & _
             "          A.VA_OPER_ATIV_RETN         AS VA_OPER_ATIV_RETN_VLRXML,           " & vbNewLine & _
             "          A.TP_CNPT,                                                         " & vbNewLine & _
             "          A.NU_CTRL_MESG_SPB_ORIG,                                           " & vbNewLine & _
             "          B.ID_PART_CAMR_CETIP,                                              " & vbNewLine & _
             "          A.CO_CNPT_CAMR,                                                    " & vbNewLine

    strSQL = strSQL & _
             "          A.TP_PGTO_LDL,                                                     " & vbNewLine & _
             "          A.CO_PARP_CAMR,                                                    " & vbNewLine & _
             "          A.CO_GRUP_LANC_FINC,                                               " & vbNewLine & _
             "          I.DE_GRUP_LANC_FINC,                                               " & vbNewLine & _
             "          J.DS_PARP_CAMR,                                                    " & vbNewLine & _
             "          A.CO_MOED_ESTR,                                                    " & vbNewLine & _
             "          A.CO_CNTR_SISB,                                                    " & vbNewLine & _
             "          A.CO_ISPB_IF_CNPT,                                                 " & vbNewLine & _
             "          A.CO_PRAC,                                                         " & vbNewLine & _
             "          A.VA_MOED_ESTR,                                                    " & vbNewLine & _
             "          A.VA_MOED_ESTR             AS VA_MOED_ESTR_VLRXML,                 " & vbNewLine & _
             "          A.DT_LIQU_OPER_ATIV_MOED_ESTR,                                     " & vbNewLine & _
             "          A.CO_SISB_COTR,                                                    " & vbNewLine & _
             "          A.TP_CNAL_VEND,                                                    " & vbNewLine & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,                                     " & vbNewLine & _
             "          A.CO_IDEF_LAST,                                                    " & vbNewLine & _
             "          A.CD_OPER_ETRT,                                                    " & vbNewLine & _
             "          A.NR_IDEF_NEGO_BMC,                                                " & vbNewLine & _
             "          A.TP_NEGO,                                                         " & vbNewLine & _
             "          A.CD_ASSO_CAMB,                                                    " & vbNewLine & _
             "          A.NR_CNPJ_CPF,                                                     " & vbNewLine & _
             "          K.DH_MESG_INTE,                                                    " & vbNewLine
    
'    strSQL = strSQL & _
'             "          A.NO_CLIE,                                                         " & vbNewLine & _
'             "          A.CD_MOED_ISO,                                                     " & vbNewLine
    
    strSQL = strSQL & _
             "          row_number() over(ORDER BY A.NU_SEQU_OPER_ATIV ASC) ID             " & vbNewLine & _
             " FROM     " & vbNewLine
    
    If blnOperacaoCetipConciliacao Then
        strSQL = strSQL & _
             "          (SELECT DISTINCT TP_OPER FROM A8.TB_TIPO_OPER_CNCL_CETIP)   OCC,   " & vbNewLine
    End If
    
    strSQL = strSQL & _
             "          A8.TB_OPER_ATIV                 A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C,                                 " & vbNewLine & _
             "          A8.TB_TIPO_OPER                 D,                                 " & vbNewLine & _
             "          A8.TB_TIPO_LIQU_OPER_ATIV       E,                                 " & vbNewLine & _
             "          A8.TB_SITU_CNTG_SIST            F,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 G,                                 " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO          H,                                 " & vbNewLine & _
             "          A8.TB_GRUP_LANC_FINC            I,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF_PARP_CAMR       J,                                 " & vbNewLine & _
             "          A8.TB_OPER_ATIV_MESG_INTE       K                                  " & vbNewLine

    strSQL = strSQL & _
             " WHERE    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                                   " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                        " & vbNewLine & _
             " AND      B.TP_BKOF               = G.TP_BKOF             (+)                        " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                                   " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER                                        " & vbNewLine & _
             " AND      A.TP_LIQU_OPER_ATIV     = E.TP_LIQU_OPER_ATIV   (+)                        " & vbNewLine & _
             " AND      A.IN_DISP_CONS          = " & enumIndicadorSimNao.Sim & "                  " & vbNewLine & _
             " AND      A.SG_SIST               = F.SG_SIST             (+)                        " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU          = H.CO_LOCA_LIQU                                   " & vbNewLine & _
             " AND      H.CO_EMPR_FUSI          = 1                                                " & vbNewLine & _
             " AND      A.CO_GRUP_LANC_FINC     = I.CO_GRUP_LANC_FINC   (+)                        " & vbNewLine & _
             " AND      A.CO_PARP_CAMR          = J.CO_PARP_CAMR        (+)                        " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV     = K.NU_SEQU_OPER_ATIV                              " & vbNewLine & _
             " AND      K.DH_MESG_INTE          = (SELECT MIN(DH_MESG_INTE)                        " & vbNewLine & _
             "                                     FROM   A8.TB_OPER_ATIV_MESG_INTE                " & vbNewLine & _
             "                                     WHERE  NU_SEQU_OPER_ATIV = A.NU_SEQU_OPER_ATIV) " & vbNewLine
             
    If blnOperacaoCetipConciliacao Then
        strSQL = strSQL & _
             " AND      OCC.TP_OPER             = A.TP_OPER                                " & vbNewLine
    End If
    
    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If
    
    'KIDA - Ordernação customizada pelo chamador
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_OrderBy/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strSQL = strSQL & vbNewLine & " ORDER BY " & strValores
    End If

    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_RowNun/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        
        ObterQueryConsultaOperacao = "SELECT * FROM (" & strSQL & ") where id <= " & strValores
        
    Else
    
        ObterQueryConsultaOperacao = strSQL
    End If

    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQueryConsultaOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    Exit Function
    Resume

End Function

'-------------------------------------------------------------------------------------------------
'Obter o Net de operações para conciliação financeira
'-------------------------------------------------------------------------------------------------

Public Function ObterNetOperacaoConciliacaoFinanceira(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                                                      ByVal strCampoIdentContraparte As String) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strAgrupaCodigoStatus                   As String
Dim strAgrupaDescricaoStatus                As String
Dim intTipoLiquidacao                       As enumTipoLiquidacao
Dim intTipoContraparte                      As enumTipoContraparte
Dim intIndicDebCred                         As enumTipoDebitoCredito
Dim blnNegacao                              As Boolean

Dim objRS                                   As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    strAgrupaCodigoStatus = _
        "DECODE(A.CO_ULTI_SITU_PROC, " & enumStatusOperacao.RegistradaAutomatica & ", " & _
                                         enumStatusOperacao.Registrada & ", " & _
        "DECODE(A.CO_ULTI_SITU_PROC, " & enumStatusOperacao.ConcordanciaBackofficeAutomatico & ", " & _
                                         enumStatusOperacao.ConcordanciaBackoffice & ", " & _
        "DECODE(A.CO_ULTI_SITU_PROC, " & enumStatusOperacao.PagamentoBackofficeAutomatico & ", " & _
                                         enumStatusOperacao.PagamentoBackoffice & ", A.CO_ULTI_SITU_PROC)))"
    
    strAgrupaDescricaoStatus = _
        "DECODE(D.DE_SITU_PROC, 'Registrada Automática'     , 'Registrada'             , " & _
        "DECODE(D.DE_SITU_PROC, 'Concordância BO Automática', 'Concordância Backoffice', " & _
        "DECODE(D.DE_SITU_PROC, 'Pagamento BO Automático'   , 'Pagamento Backoffice'   , D.DE_SITU_PROC)))"
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_StatusOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Data
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
        strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV between " & objDomNode.selectSingleNode("DataIni").Text & " AND " & vbNewLine & _
                                                                          objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode

    'Filtro ISPB Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ISPBContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If strValores = "''" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.CO_ISPB_BANC_LIQU_CNPT      IN ('') " & vbCrLf
            strWhereFiltro = strWhereFiltro & " OR       A.CO_ISPB_BANC_LIQU_CNPT      IS NULL)" & vbCrLf
        ElseIf strValores = "'NotNull'" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.CO_ISPB_BANC_LIQU_CNPT  NOT IN ('') " & vbCrLf
            strWhereFiltro = strWhereFiltro & " OR       A.CO_ISPB_BANC_LIQU_CNPT  IS NOT NULL)" & vbCrLf
        Else
            strWhereFiltro = strWhereFiltro & " AND      A.CO_ISPB_BANC_LIQU_CNPT      IN (" & strValores & ")" & vbCrLf
        End If
    End If
    
    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro Tipo de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
        intTipoLiquidacao = Val(objDomNode.Text)
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_LIQU_OPER_ATIV      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro Tipo de Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoContraparte/*")
        strValores = strValores & objDomNode.Text & ", "
        intTipoContraparte = Val(objDomNode.Text)
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNPT      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro por Indicador de Débito e Crédito
    If intTipoLiquidacao = enumTipoLiquidacao.Bruta Then
        intIndicDebCred = 0
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NaturezaMovimento/*")
            intIndicDebCred = Val(objDomNode.Text)
        Next
        
        If intIndicDebCred <> 0 Then
            intIndicDebCred = IIf(intIndicDebCred = enumTipoDebitoCredito.Debito, _
                                                    enumTipoDebitoCredito.Credito, _
                                                    enumTipoDebitoCredito.Debito)
            
            strWhereFiltro = strWhereFiltro & " AND      A.IN_OPER_DEBT_CRED  IN (" & intIndicDebCred & ")"
        End If
    End If
    
    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Canal Operação Interna
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalOperacaoInterna/*")
        If InStr(1, objDomNode.Text, "CONDIÇÃO") = 0 Then
            strValores = strValores & "'" & objDomNode.Text & "', "
        Else
            strValores = strValores & objDomNode.Text
            Exit For
        End If
    Next
    
    If Trim(strValores) <> "" Then
        If InStr(1, strValores, "CONDIÇÃO") = 0 Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.CO_CNAL_OPER_INTE  IN (" & strValores & ")"
        Else
            strValores = Replace$(strValores, "CONDIÇÃO", vbNullString)
            strWhereFiltro = strWhereFiltro & " AND " & strValores
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    
    strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "B", False, False, True, True, True, False)
    
    If intTipoLiquidacao = enumTipoLiquidacao.Bruta Then
        strSQL = " SELECT   A.CO_EMPR,                                                         " & vbNewLine & _
                 "          A.CO_ISPB_BANC_LIQU_CNPT,                                          " & vbNewLine & _
                 "          A.CO_CNPJ_CNPT,                                                    " & vbNewLine & _
                 "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
                 "          A.NO_CNPT,                                                         " & vbNewLine & _
                 "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
                 "          A.NU_CTRL_MESG_SPB_ORIG,                                           " & vbNewLine & _
                 "          A.IN_OPER_DEBT_CRED,                                               " & vbNewLine & _
                 "          A.VA_OPER_ATIV,                                                    " & vbNewLine & _
                 "          A.TP_OPER,                                                         " & vbNewLine & _
                 "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
                 "          B.CO_CNPJ_VEIC_LEGA,                                               " & vbNewLine & _
                 "          A.CO_PARP_CAMR,                                                    " & vbNewLine & _
                 "          C.NO_ISPB,                                                         " & vbNewLine & _
                 "          D.DE_SITU_PROC,                                                    " & vbNewLine & _
                 "          A.NR_CNPJ_CPF                                                      " & vbNewLine
             
    Else
    
        strSQL = " SELECT   A.CO_EMPR,                                                         " & vbNewLine & _
                 "          B.CO_CNPJ_VEIC_LEGA,                                               " & vbNewLine

        If intTipoContraparte = enumTipoContraparte.Interno Then
            'RATS 586/001
            'Problema produção - Cassiano - 07/11/2006
            strSQL = strSQL & _
                 "          A." & strCampoIdentContraparte & ",                                " & vbNewLine
        Else
            strSQL = strSQL & _
                 "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
                 "          A.CO_PARP_CAMR,                                                    " & vbNewLine & _
                 "          A.CO_CNPT_CAMR,                                                    " & vbNewLine
        
        End If

        strSQL = strSQL & _
                 "          A.CO_ISPB_BANC_LIQU_CNPT,                                          " & vbNewLine & _
                 "          " & strAgrupaCodigoStatus & "    AS CO_ULTI_SITU_PROC,             " & vbNewLine & _
                 "          C.NO_ISPB,                                                         " & vbNewLine & _
                 "          '' AS NO_CNPT,                                                     " & vbNewLine & _
                 "          " & strAgrupaDescricaoStatus & " AS DE_SITU_PROC,                  " & vbNewLine & _
                 "          NVL(SUM(DECODE(A.IN_OPER_DEBT_CRED, 1, A.VA_OPER_ATIV, 0)),0) -    " & vbNewLine & _
                 "          NVL(SUM(DECODE(A.IN_OPER_DEBT_CRED, 2, A.VA_OPER_ATIV, 0)),0)      " & vbNewLine & _
                 "          AS VA_OPER_ATIV                                                    " & vbNewLine

    End If
    
    strSQL = strSQL & _
             " FROM     A8.TB_OPER_ATIV                 A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_INSTITUICAO_SPB           C,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 D                                  " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA                = B.CO_VEIC_LEGA                     " & vbNewLine & _
             " AND      A.SG_SIST                     = B.SG_SIST                          " & vbNewLine & _
             " AND      A.CO_ISPB_BANC_LIQU_CNPT      = C.CO_ISPB             (+)          " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC           = D.CO_SITU_PROC                     " & vbNewLine
    
    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & strWhereSegregacao
    End If
    
    If intTipoLiquidacao = enumTipoLiquidacao.Bilateral Then
        strSQL = strSQL & vbNewLine & _
                 " GROUP BY A.CO_EMPR,                          " & vbNewLine & _
                 "          B.CO_CNPJ_VEIC_LEGA,                " & vbNewLine

        If intTipoContraparte = enumTipoContraparte.Interno Then
            'RATS 586/001
            'Problema produção - Cassiano - 07/11/2006
            strSQL = strSQL & _
                 "          A." & strCampoIdentContraparte & ", " & vbNewLine
        Else
            strSQL = strSQL & _
                 "          B.NO_VEIC_LEGA,                     " & vbNewLine & _
                 "          A.CO_PARP_CAMR,                     " & vbNewLine & _
                 "          A.CO_CNPT_CAMR,                     " & vbNewLine
        End If

        strSQL = strSQL & _
                 "          A.CO_ISPB_BANC_LIQU_CNPT,           " & vbNewLine & _
                 "          " & strAgrupaCodigoStatus & ",      " & vbNewLine & _
                 "          C.NO_ISPB,                          " & vbNewLine & _
                 "          '',                                 " & vbNewLine & _
                 "          " & strAgrupaDescricaoStatus & "    " & vbNewLine
    End If
    
    If intTipoContraparte <> enumTipoContraparte.Interno Or intTipoLiquidacao <> enumTipoLiquidacao.Bilateral Then
        strSQL = strSQL & vbNewLine & _
                 " ORDER BY B.NO_VEIC_LEGA "
    Else
        strSQL = strSQL & vbNewLine & _
                 " ORDER BY B.CO_CNPJ_VEIC_LEGA "
    End If
    
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLeitura.loadXML(fgQueryXMLLerTodos("NetOperacao", strSQL, App.EXEName & "." & TypeName(Me)))
    
    If intTipoLiquidacao = enumTipoLiquidacao.Bilateral And intTipoContraparte = enumTipoContraparte.Externo Then
        For Each objDomNode In xmlLeitura.selectNodes("Repeat_NetOperacao/*")
            strSQL = " SELECT A.NO_CNPT         " & vbNewLine & _
                     " FROM   A8.TB_OPER_ATIV A " & vbNewLine & _
                     " WHERE  A.CO_EMPR      =  " & objDomNode.selectSingleNode("CO_EMPR").Text & "       " & vbNewLine & _
                     " AND    A.CO_CNPT_CAMR = '" & objDomNode.selectSingleNode("CO_CNPT_CAMR").Text & "' " & vbNewLine
                     
            If Trim(strWhereFiltro) <> "" Then
                strSQL = strSQL & vbNewLine & strWhereFiltro
            End If
            
            Set objRS = fgQuerySQL(strSQL)
            
            If objRS.RecordCount > 0 Then
                objDomNode.selectSingleNode("NO_CNPT").Text = objRS!NO_CNPT
            End If
        Next
    End If
    
    ObterNetOperacaoConciliacaoFinanceira = xmlLeitura.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlLeitura = Nothing
    
    Exit Function
    
ErrorHandler:
    Set xmlLeitura = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetOperacaoConciliacaoFinanceira Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter o Net de operações para conciliação física
'-------------------------------------------------------------------------------------------------

Public Function ObterNetOperacaoConciliacaoFisica(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String

    On Error GoTo ErrorHandler
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_StatusOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      B.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro Tipo de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_LIQU_OPER_ATIV      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro por Data
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacao")
        strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_ATIV between " & fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataIni").Text) & " AND" & vbNewLine & _
                                                                          fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataFim").Text)
    Next objDomNode

    '---------------------------------------------------------------------------------------------------------
    
    strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "B", False, False, True, True, True, False)
    
    strSQL = " SELECT   A.NU_ATIV_MERC,                                                    " & vbNewLine & _
             "          A.DT_VENC_ATIV,                                                    " & vbNewLine & _
             "          A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,                                    " & vbNewLine & _
             "          NVL(SUM(DECODE(A.IN_OPER_DEBT_CRED, 2, A.QT_ATIV_MERC, 0)),0) -    " & vbNewLine & _
             "          NVL(SUM(DECODE(A.IN_OPER_DEBT_CRED, 1, A.QT_ATIV_MERC, 0)),0)      " & vbNewLine & _
             "          AS QT_ATIV_MERC,                                                   " & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,                                    " & vbNewLine & _
             "          B.TP_BKOF,                                                         " & vbNewLine & _
             "          C.DE_SITU_PROC                                                     " & vbNewLine & _
             " FROM     A8.TB_OPER_ATIV                 A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C                                  " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                           " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                           " & vbNewLine
    
    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & strWhereSegregacao
    End If
    
    strSQL = strSQL & _
             " GROUP BY A.NU_ATIV_MERC, A.DT_VENC_ATIV, A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, A.NU_COMD_OPER, A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, B.TP_BKOF,C.DE_SITU_PROC"
    
    ObterNetOperacaoConciliacaoFisica = fgQueryXMLLerTodos("NetOperacao", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetOperacaoConciliacaoFisica Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter uma operação pelo numero de controle RDC original para complementar a operação recebida do leagdo
'-------------------------------------------------------------------------------------------------

Public Function ComplOperacaoComRDCOriginal(ByRef xmlOperacao As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String
Dim strOperacao                             As String
Dim xmlOperacaoOriginal                     As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlOperacaoOriginal = CreateObject("MSXML2.DOMDocument.4.0")

    strOperacao = LerTodos(, , , , , , xmlOperacao.documentElement.selectSingleNode("NU_CTRL_RDSC_ORIG").Text)
    If strOperacao = vbNullString Then
        ' 3028 - Número de controle RDC original inválido.
        lngCodigoErroNegocio = 3028
        GoTo ErrorHandler
    End If
    xmlOperacaoOriginal.loadXML strOperacao

    fgAppendNode xmlOperacao, "MESG", "CO_VEIC_LEGA", xmlOperacaoOriginal.documentElement.selectSingleNode("Grupo_Operacao/CO_VEIC_LEGA").Text
    xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text = xmlOperacaoOriginal.documentElement.selectSingleNode("Grupo_Operacao/SG_SIST").Text

    'Incluido 22/04 para aceitar pagto por lastro - Carlos
    fgAppendNode xmlOperacao, "MESG", "DT_VENC_ATIV", xmlOperacaoOriginal.documentElement.selectSingleNode("Grupo_Operacao/DT_VENC_ATIV").Text
    fgAppendNode xmlOperacao, "MESG", "NU_ATIV_MERC", xmlOperacaoOriginal.documentElement.selectSingleNode("Grupo_Operacao/NU_ATIV_MERC").Text
    fgAppendNode xmlOperacao, "MESG", "CO_IDEF_LAST", xmlOperacaoOriginal.documentElement.selectSingleNode("Grupo_Operacao/CO_IDEF_LAST").Text

    If xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Then
        strOperacao = LerTodos(, , , , , enumTipoOperacaoLQS.ConversaoRedesconto, _
                               xmlOperacao.documentElement.selectSingleNode("NU_CTRL_RDSC_ORIG").Text)
        If strOperacao = vbNullString Then
            fgAppendNode xmlOperacao, "MESG", "PREV_PJ_ENV", enumIndicadorSimNao.Nao
            Exit Function
        End If
        xmlOperacaoOriginal.loadXML strOperacao

        If CLng(xmlOperacaoOriginal.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
            'Marcar que a previsão do PJ  já foi enviada
            fgAppendNode xmlOperacao, "MESG", "PREV_PJ_ENV", enumIndicadorSimNao.Sim
        Else
            fgAppendNode xmlOperacao, "MESG", "PREV_PJ_ENV", enumIndicadorSimNao.Nao
        End If
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ComplOperacaoComRDCOriginal Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Alterar o agendamento de envio de uma operação
'-------------------------------------------------------------------------------------------------

Public Function AlterarAgendamento(ByVal pvntSequenciaOperacao As Variant, _
                                   ByVal pstrHorarioAgendamento As String, _
                                   ByVal pstrDHUltimaAtualizacao As String, _
                                   ByVal plngLocalLiquidacao As Long, _
                                   ByVal pstrCodigoMensagem As String, _
                                   ByVal pintStatusOperacao As enumStatusOperacao)

Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim xmlGradeHorario                         As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim lngSequencialMensagem                   As Long
Dim lngSequencialISPBDestino                As Long
Dim strHorarioAbertura                      As String
Dim strHorarioEncerramento                  As String

    On Error GoTo ErrorHandler
    
    If pstrHorarioAgendamento = vbNullString Then pstrHorarioAgendamento = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) & "000000"
    
    Set xmlGradeHorario = CreateObject("MSXML2.DOMDocument.4.0")
    
    '-------------------------------------------------------------------------------------------------------
    '>>> Verificar Grade de Horário
    '-------------------------------------------------------------------------------------------------------
    '    Captura ISPB Destino
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    lngSequencialISPBDestino = objMensagem.ObterSequencialISPBDestino(pstrCodigoMensagem, _
                                                                      , _
                                                                      plngLocalLiquidacao)
    Set objMensagem = Nothing
    '-------------------------------------------------------------------------------------------------------
    '    Captura a Grade de Horário
    Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
    Call xmlGradeHorario.loadXML(objGradeHorario.ObterGradeHorario(pstrCodigoMensagem, lngSequencialISPBDestino))
    Set objGradeHorario = Nothing
    
    strHorarioAbertura = Mid(xmlGradeHorario.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioAbertura").Text, 9, 6)
    strHorarioEncerramento = Mid(xmlGradeHorario.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioEncerramento").Text, 9, 6)
    '-------------------------------------------------------------------------------------------------------
    'Verifica se o horário está fora da Grade
    If TimeSerial(Mid(pstrHorarioAgendamento, 9, 2), Mid(pstrHorarioAgendamento, 11, 2), Mid(pstrHorarioAgendamento, 13, 2)) < _
       TimeSerial(Mid(strHorarioAbertura, 1, 2), Mid(strHorarioAbertura, 3, 2), Mid(strHorarioAbertura, 5, 2)) Or _
       TimeSerial(Mid(pstrHorarioAgendamento, 9, 2), Mid(pstrHorarioAgendamento, 11, 2), Mid(pstrHorarioAgendamento, 13, 2)) > _
       TimeSerial(Mid(strHorarioEncerramento, 1, 2), Mid(strHorarioEncerramento, 3, 2), Mid(strHorarioEncerramento, 5, 2)) Then
       
        'Agendamento fora da Grade de Horário estipulada.
        lngCodigoErroNegocio = 3020
        GoTo ErrorHandler
    End If
    
    Set xmlGradeHorario = Nothing
    
    strSQL = " Update   A8.TB_OPER_ATIV                                                               " & vbNewLine & _
             " Set      HO_ENVI_MESG_SPB        = " & fgDtHrXML_To_Oracle(pstrHorarioAgendamento) & "," & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "',                            " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "',                 " & vbNewLine & _
             "          DH_ULTI_ATLZ            = Sysdate                                             " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV       = " & pvntSequenciaOperacao & "                       " & vbNewLine & _
             " And      DH_ULTI_ATLZ            = " & fgDtHrXML_To_Oracle(pstrDHUltimaAtualizacao)

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    'Armazenar o Histórico do Agendamento da Operação
    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    objHistoricoOperacao.Incluir pvntSequenciaOperacao, _
                                 pintStatusOperacao, _
                                 , _
                                 enumTipoAcao.AlteracaoHorarioAgendamento, _
                                 pstrHorarioAgendamento
    Set objHistoricoOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objHistoricoOperacao = Nothing
    Set objGradeHorario = Nothing
    Set objMensagem = Nothing
    Set xmlGradeHorario = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter uma operação pelo numero de controle IF da mensagem SPB
'-------------------------------------------------------------------------------------------------

Public Function ObterOperacaoPorControleIF(ByVal pstrNumeroControleIF As String, _
                                  Optional ByVal pblnRetornarErro As Boolean = True, _
                                  Optional ByVal pstrCodigoMensagemSPB As String = vbNullString, _
                                  Optional ByVal pblnForUpdate As Boolean = True) As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim rsOperacao                              As ADODB.Recordset

    On Error GoTo ErrorHandler

    Set rsOperacao = CreateObject("ADODB.Recordset")
    Set ObterOperacaoPorControleIF = CreateObject("MSXML2.DOMDocument.4.0")

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV, " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC, " & vbCrLf & _
             "          A.DH_ULTI_ATLZ, " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          B.CO_TEXT_XML " & vbCrLf & _
             " From     A8.TB_OPER_ATIV             A, " & vbCrLf & _
             "          A8.TB_MESG_RECB_ENVI_SPB    B  " & vbCrLf & _
             " Where    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
             " And      B.NU_CTRL_IF            =   '" & pstrNumeroControleIF & "'" & vbCrLf

    If pstrCodigoMensagemSPB = vbNullString Then
        strSQL = strSQL & _
             " And      B.DH_REGT_MESG_SPB      =   (Select Min(DH_REGT_MESG_SPB) " & vbCrLf & _
             "                                       From   A8.TB_MESG_RECB_ENVI_SPB " & vbCrLf & _
             "                                       Where  NU_CTRL_IF      =   '" & pstrNumeroControleIF & "')" & vbCrLf
    Else
        strSQL = strSQL & _
             " And      B.CO_MESG_SPB            =   '" & pstrCodigoMensagemSPB & "'" & vbCrLf
    End If

   If pblnForUpdate And Not fgEstacaoDesenvolvimento Then
        strSQL = strSQL & _
                 " For Update Of A.NU_SEQU_OPER_ATIV "
    End If
    
    Set rsOperacao = fgQuerySQL(strSQL)

    If rsOperacao.EOF Then
        If pblnRetornarErro Then
            ' 3022 - Número de Controle IF da Mensagem R1 inválido.
            lngCodigoErroNegocio = 3022
            GoTo ErrorHandler
        Else
            Exit Function
        End If
    End If

    ObterOperacaoPorControleIF.loadXML fgSelectVarchar4000(rsOperacao!CO_TEXT_XML)

    'Adiciona o tipo da operação, pois na mensagem original não existe a mesma.
    If ObterOperacaoPorControleIF.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        fgAppendNode ObterOperacaoPorControleIF, "SISMSG", "TP_OPER", rsOperacao!TP_OPER
    Else
        ObterOperacaoPorControleIF.documentElement.selectSingleNode("TP_OPER").Text = rsOperacao!TP_OPER
    End If
    If ObterOperacaoPorControleIF.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode ObterOperacaoPorControleIF, "SISMSG", "NU_SEQU_OPER_ATIV", rsOperacao!NU_SEQU_OPER_ATIV
    Else
        ObterOperacaoPorControleIF.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = rsOperacao!NU_SEQU_OPER_ATIV
    End If
    If ObterOperacaoPorControleIF.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode ObterOperacaoPorControleIF, "SISMSG", "CO_ULTI_SITU_PROC", rsOperacao!CO_ULTI_SITU_PROC
    Else
        ObterOperacaoPorControleIF.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = rsOperacao!CO_ULTI_SITU_PROC
    End If
    If ObterOperacaoPorControleIF.documentElement.selectSingleNode("DH_ULTI_ATLZ") Is Nothing Then
        fgAppendNode ObterOperacaoPorControleIF, "SISMSG", "DH_ULTI_ATLZ", fgDtHr_To_Xml(rsOperacao!dh_ulti_atlz)
    Else
        ObterOperacaoPorControleIF.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text = rsOperacao!dh_ulti_atlz
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacaoPorControleIF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Verificar a pendencia de Repasse Financeiro
'-------------------------------------------------------------------------------------------------

Public Function VerificarRepassePendente(ByVal plngSequenciaConciliacao As Long) As Boolean

Dim strSQL                                  As String
Dim rsOperacao                              As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsOperacao = CreateObject("ADODB.Recordset")

    strSQL = "  SELECT  NU_SEQU_OPER_ATIV,  " & vbCrLf & _
             "          CO_ULTI_SITU_PROC   " & vbCrLf & _
             "   FROM   A8.TB_OPER_ATIV     " & vbCrLf & _
             "  WHERE   NU_SEQU_CNCL_OPER_ATIV_MESG     =   " & plngSequenciaConciliacao & vbCrLf & _
             "    AND   CO_ULTI_SITU_PROC               <>  " & enumStatusOperacao.Liquidada

    Set rsOperacao = fgQuerySQL(strSQL)

    If rsOperacao.RecordCount > 0 Then
        VerificarRepassePendente = True
    Else
        VerificarRepassePendente = False
    End If

    Set rsOperacao = Nothing

    If Not objContext Is Nothing Then
       'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set rsOperacao = Nothing

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarRepassePendente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler tipos de operação conforme filtro especificado.
Private Function flLerTodosTiposOperacao(ByVal pxmlFiltro As MSXML2.DOMDocument40) As String

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
    
    strSQL = " SELECT  A.TP_OPER,          " & vbNewLine & _
             "         A.NO_TIPO_OPER,     " & vbNewLine & _
             "         A.TP_MESG_RECB_INTE " & vbNewLine & _
             " FROM    A8.TB_TIPO_OPER   A " & vbNewLine

    If strWhereFiltro <> vbNullString Then
        strWhereFiltro = " WHERE" & strWhereFiltro
        strWhereFiltro = Replace$(strWhereFiltro, "WHERE AND", "WHERE")
        strSQL = strSQL & strWhereFiltro
    End If
    
    strSQL = strSQL & " ORDER BY A.NO_TIPO_OPER "
                 
    flLerTodosTiposOperacao = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
       objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLerTodosTiposOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter valores das operações por Grupo de Veiculo Legal
'-------------------------------------------------------------------------------------------------

Public Function ObterValoresPorGrupoVeiculoLegal(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim xmlDomRetorno                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strRetorno                              As String
Dim strRetornoAux                           As String

    On Error GoTo ErrorHandler
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Tipo Operação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER       IN (" & strValores & ")"
    End If

    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Tipo Operação Rotina Abertura
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoRotinaAbertura/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      D.CO_OPER_SELIC      IS NOT NULL" & vbCrLf
    End If
    
    'Filtro por Data
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data")
        strWhereFiltro = strWhereFiltro & " AND      A.DT_OPER_ATIV between " & _
                        objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                        objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode
    '---------------------------------------------------------------------------------------------------------
    
    strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "B", False, False, False, False, True, False)
    
    '1º Captura os valores por Grupo de Veículo Legal...
    strSQL = " SELECT   B.CO_GRUP_VEIC_LEGA,                            " & vbNewLine & _
             "          C.NO_GRUP_VEIC_LEGA,                            " & vbNewLine & _
             "          SUM(A.VA_OPER_ATIV)     VA_OPER_ATIV            " & vbNewLine & _
             " FROM     A8.TB_OPER_ATIV         A,                      " & vbNewLine & _
             "          A8.TB_VEIC_LEGA         B,                      " & vbNewLine & _
             "          A8.TB_GRUP_VEIC_LEGA    C,                      " & vbNewLine & _
             "          A8.TB_TIPO_OPER         D                       " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA        " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST             " & vbNewLine & _
             " AND      B.CO_GRUP_VEIC_LEGA     = C.CO_GRUP_VEIC_LEGA   " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER             "
             
    If Trim(strWhereSegregacao) <> "" Then
        strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If
    
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                strWhereFiltro
    End If
    
    strSQL = strSQL & vbNewLine & _
            " GROUP BY  B.CO_GRUP_VEIC_LEGA,                            " & vbNewLine & _
            "           C.NO_GRUP_VEIC_LEGA"

    strRetorno = fgQueryXMLLerTodos("ValoresPorGrupoVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))
    
    '...depois, captura as operações que compõe o total acima
    strSQL = " SELECT   A.NU_SEQU_OPER_ATIV,                            " & vbNewLine & _
             "          A.DH_ULTI_ATLZ                                  " & vbNewLine & _
             " FROM     A8.TB_OPER_ATIV         A,                      " & vbNewLine & _
             "          A8.TB_VEIC_LEGA         B,                      " & vbNewLine & _
             "          A8.TB_GRUP_VEIC_LEGA    C,                      " & vbNewLine & _
             "          A8.TB_TIPO_OPER         D                       " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA        " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST             " & vbNewLine & _
             " AND      B.CO_GRUP_VEIC_LEGA     = C.CO_GRUP_VEIC_LEGA   " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER             "
             
    If Trim(strWhereSegregacao) <> "" Then
        strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If
    
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                strWhereFiltro
    End If
    
    strRetornoAux = fgQueryXMLLerTodos("CodigoOperacao", strSQL, App.EXEName & "." & TypeName(Me))
    
    'Monta o XML híbrido de retorno para a interface
    Set xmlDomRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomRetorno, "", "ConciliacaoOperacao", "")
    If Trim(strRetorno) <> vbNullString Then Call fgAppendXML(xmlDomRetorno, "ConciliacaoOperacao", strRetorno)
    If Trim(strRetornoAux) <> vbNullString Then Call fgAppendXML(xmlDomRetorno, "ConciliacaoOperacao", strRetornoAux)
    
    ObterValoresPorGrupoVeiculoLegal = xmlDomRetorno.xml
    
    Set xmlDomRetorno = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set xmlDomRetorno = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresPorGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Tratar a inclusão específica da Liquidação Fisica de uma operação
'-------------------------------------------------------------------------------------------------

Private Function flIncluirLiquidacaoFisicaOper(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                               ByRef plngStatusOperacao As Long) As Boolean

Dim strSQL                                  As String
Dim strValoresVariaveis                     As String
Dim strValoresVariaveis1                     As String
Dim strValoresFixos                         As String
Dim strValoresNulos                         As String
Dim strColunas                              As String
Dim lngTipoOperacao                         As Long
Dim vntSequencialOperacao                   As Variant
Dim strCodigoMensagem                       As String
Dim strOperacao                             As String
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodeAux                           As MSXML2.IXMLDOMNode
Dim xmlTipoLiquidacao                       As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objHistoricoSituacaoOperacao            As A8LQS.clsHistSituacaoOperacao
Dim objTipoLiquidacao                       As A6A7A8.clsTipoLiquidacao
Dim strIdentificadorOperacao                As String
Dim strMensagemRetornoLegado                As String
Dim strTipoLiquidacao                       As String

Dim intEntradaManual                        As enumIndicadorSimNao

On Error GoTo ErrorHandler

    If Not xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
        intEntradaManual = Val(xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU").Text)
    Else
        intEntradaManual = enumIndicadorSimNao.Nao
    End If

    'Irá segurar a inclusão, pois o Codigo da Operção e a Sigla do Sistema não podem duplicar e o indice
    ' não pode ser único pois temos a compromissada de ida que vem com estes iguais modificando somente
    ' o identificado do título
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    objTipoOperacao.BloquearTipoOperacaoParaInclusao

    If intEntradaManual = enumIndicadorSimNao.Nao Then
        'testar a duplicidade pelo identificador do titulo
        'Verificar se existe o Identificador da Operação para o Sistema Origem e o Identificador do Titulo
        strOperacao = LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

        If strOperacao <> vbNullString Then
            '3037 - Identificador da Operação já existe.
            lngCodigoErroNegocio = 3037
            GoTo ErrorHandler
        End If
    Else
        strOperacao = LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

        If strOperacao <> vbNullString Then
            '3037 - Identificador da Operação já existe.
            lngCodigoErroNegocio = 3037
            GoTo ErrorHandler
        End If
    End If

    'Obter a Operação correspondente a Mensagem do BUS
    fgAppendNode xmlRemessa, "MESG", "CO_LOCA_LIQU", enumLocalLiquidacao.BMA
    objTipoOperacao.ObterTipoOperacaoPorMsgBUS xmlRemessa, lngTipoOperacao, strCodigoMensagem, strMensagemRetornoLegado

    If strCodigoMensagem = vbNullString Then
        If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            Set objTipoLiquidacao = CreateObject("A6A7A8.clsTipoLiquidacao")
            strTipoLiquidacao = objTipoLiquidacao.Ler(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
            If strTipoLiquidacao <> vbNullString Then
                Set xmlTipoLiquidacao = CreateObject("MSXML2.DOMDocument.4.0")
                xmlTipoLiquidacao.loadXML strTipoLiquidacao
                strCodigoMensagem = xmlTipoLiquidacao.documentElement.selectSingleNode("CO_MESG_SPB_LIQU").Text
                Set xmlTipoLiquidacao = Nothing
            End If
            Set objTipoLiquidacao = Nothing
        Else
            fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral
        End If
    End If

    strColunas = " Insert Into  A8.TB_OPER_ATIV (" & vbCrLf & _
             " TP_OPER, " & vbCrLf & _
             " CO_EMPR, " & vbCrLf & _
             " CO_USUA_CADR_OPER, " & vbCrLf & _
             " SG_SIST, " & vbCrLf & _
             " CO_OPER_ATIV, " & vbCrLf & _
             " CO_LOCA_LIQU, " & vbCrLf & _
             " NU_COMD_OPER, " & vbCrLf & _
             " DT_OPER_ATIV, " & vbCrLf & _
             " CO_ULTI_SITU_PROC, " & vbCrLf & _
             " CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             " DH_ULTI_ATLZ, " & vbCrLf & _
             " IN_DISP_CONS, " & vbCrLf & _
             " IN_ENTR_MANU, " & vbCrLf

    strColunas = strColunas & _
             " NU_ATIV_MERC, " & vbCrLf & _
             " DE_ATIV_MERC, " & vbCrLf & _
             " DT_VENC_ATIV, " & vbCrLf & _
             " NU_SEQU_OPER_ATIV, " & vbCrLf & _
             " CO_VEIC_LEGA, " & vbCrLf & _
             " CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbCrLf & _
             " CO_CNPJ_CNPT, " & vbCrLf & _
             " NO_CNPT, " & vbCrLf & _
             " CO_CNTA_CUTD_SELIC_CNPT, " & vbCrLf & _
             " IN_OPER_DEBT_CRED, " & vbCrLf & _
             " IN_ENTR_SAID_RECU_FINC, " & vbCrLf & _
             " QT_ATIV_MERC, " & vbCrLf & _
             " PU_ATIV_MERC, " & vbCrLf & _
             " VA_OPER_ATIV, " & vbCrLf

    strColunas = strColunas & _
             " TP_LIQU_OPER_ATIV,    HO_ENVI_MESG_SPB,            NU_COMD_OPER_RETN,      DT_OPER_ATIV_RETN,    DT_LIQU_OPER_ATIV, " & vbCrLf & _
             " TP_CPRO_OPER_ATIV,    TP_CPRO_RETN_OPER_ATIV,      IN_ENVI_PREV_SIST_PJ,   IN_ENVI_RELZ_SIST_PJ, " & vbCrLf & _
             " IN_ENVI_PREV_SIST_A6, IN_ENVI_RELZ_SIST_A6,        TP_ACAO_OPER_ATIV_EXEC, NU_COMD_ACAO_EXEC, " & vbCrLf & _
             " NU_PRTC_OPER_LG,      NU_SEQU_CNCL_OPER_ATIV_MESG, NU_CTRL_MESG_SPB_ORIG) " & vbCrLf

    strValoresNulos = " 1, Null, Null, Null, Null, " & vbCrLf & _
                      " Null, Null, Null, Null, Null, " & vbCrLf & _
                      " Null, Null, Null, Null, Null, " & vbCrLf & _
                      " Null) " & vbCrLf

    strValoresFixos = " Values (" & vbCrLf & _
                     lngTipoOperacao & ", " & vbCrLf & _
                     xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text & ", " & vbCrLf & _
                     "'" & xmlRemessa.documentElement.selectSingleNode("CO_USUA_CADR_OPER").Text & "', " & vbCrLf & _
                     "'" & xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "', " & vbCrLf & _
                     "'" & xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text & "', " & vbCrLf & _
                     xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text & ", " & vbCrLf & _
                     "Null, " & vbCrLf & _
                     fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text) & ", " & vbCrLf & _
                     plngStatusOperacao & ", " & vbCrLf & _
                     "'" & fgUsuarioRede & "', " & vbCrLf & _
                     "'" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
                     "SYSDATE, " & vbCrLf & _
                     enumIndicadorSimNao.Sim & ", " & vbCrLf & _
                     intEntradaManual & ", " & vbCrLf

    For Each objDomNode In xmlRemessa.selectNodes("MESG/REPE_TITL/GRUPO_TITULO")
        strValoresVariaveis1 = vbNullString
        strValoresVariaveis1 = "'" & objDomNode.selectSingleNode("NU_ATIV_MERC").Text & "', " & vbCrLf & _
                               "'" & objDomNode.selectSingleNode("DE_ATIV_MERC").Text & "', " & vbCrLf & _
                               fgDtXML_To_Oracle(objDomNode.selectSingleNode("DT_VENC_ATIV").Text) & ", " & vbCrLf
        For Each objDomNodeAux In objDomNode.selectNodes("//GRUPO_TITULO/REPE_CNTA/GRUPO_CONTA")
            strSQL = vbNullString
            strValoresVariaveis = vbNullString
            vntSequencialOperacao = fgExecuteSequence("A8.SQ_A8_NU_SEQU_OPER_ATIV")
            strValoresVariaveis = strValoresVariaveis & vntSequencialOperacao & ", " & vbCrLf & _
                                  "'" & objDomNodeAux.selectSingleNode("CO_VEIC_LEGA").Text & "', " & vbCrLf & _
                                  objDomNodeAux.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text & ", " & vbCrLf & _
                                  "Null, " & vbCrLf & _
                                  "Null, " & vbCrLf & _
                                  "Null, " & vbCrLf & _
                                  objDomNodeAux.selectSingleNode("IN_OPER_DEBT_CRED").Text & ", " & vbCrLf

            If CLng(objDomNodeAux.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                strValoresVariaveis = strValoresVariaveis & enumIndicadorEntradaSaida.ENTRADA & ", " & vbCrLf
            Else
                strValoresVariaveis = strValoresVariaveis & enumIndicadorEntradaSaida.Saida & ", " & vbCrLf
            End If

            strValoresVariaveis = strValoresVariaveis & fgVlrToDBServer(objDomNodeAux.selectSingleNode("QT_ATIV_MERC").Text) & ", " & vbCrLf & _
                                  fgVlrToDBServer(objDomNodeAux.selectSingleNode("PU_ATIV_MERC").Text) & ", " & vbCrLf & _
                                  fgVlrToDBServer(objDomNodeAux.selectSingleNode("VA_OPER_ATIV").Text) & ", " & vbCrLf

            strSQL = strColunas & strValoresFixos & strValoresVariaveis1 & strValoresVariaveis & strValoresNulos

            fgExecuteSQL strSQL

            'Armazenar a mensagem original
            Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
            objOperacaoMensagemInterna.Incluir vntSequencialOperacao, _
                                               CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                               CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                               xmlRemessa
            Set objOperacaoMensagemInterna = Nothing

            'Armazenar o Histórico da Situacao da Operação
            Set objHistoricoSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistoricoSituacaoOperacao.Incluir vntSequencialOperacao, _
                                                 plngStatusOperacao
            Set objHistoricoSituacaoOperacao = Nothing

            If xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", vntSequencialOperacao
            Else
                xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vntSequencialOperacao
            End If
            
            If xmlRemessa.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "TP_OPER", lngTipoOperacao
            Else
                xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = lngTipoOperacao
            End If

            If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", plngStatusOperacao
            Else
                xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao
            End If
            
            If xmlRemessa.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "CO_MESG", strCodigoMensagem
            Else
                xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text = strCodigoMensagem
            End If
            
            If xmlRemessa.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "TP_ACAO_OPER_ATIV_EXEC", "0"
            Else
                xmlRemessa.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = "0"
            End If
            
            If xmlRemessa.documentElement.selectSingleNode("TP_MESG_RETN_INTE") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "TP_MESG_RETN_INTE", strMensagemRetornoLegado
            Else
                xmlRemessa.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text = strMensagemRetornoLegado
            End If

        Next
    Next

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirLiquidacaoFisicaOper Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Ler operação pelo tipo de solicitação
'-------------------------------------------------------------------------------------------------

Public Function LerPorSolicitacao(ByVal pstrCodigoOperacao As String, _
                                  ByVal pstrSiglaSistema As String, _
                                  ByVal plngTipoSolicitacao As Long, _
                         Optional ByVal pstrIdentificadorTitulo As String, _
                         Optional ByVal pstrDataVencimento As String = "", _
                         Optional ByVal plngTipoOperacao As Long = 0, _
                         Optional ByVal pstrIdentificadorLastro As String = "", _
                         Optional ByVal pblnForUpdate As Boolean = True) As ADODB.Recordset

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Select    A.NU_SEQU_OPER_ATIV,            " & vbNewLine & _
             "          A.DH_MESG_INTE,                 " & vbNewLine & _
             "          A.TP_MESG_INTE,                 " & vbNewLine & _
             "          A.TP_SOLI_MESG_INTE,            " & vbNewLine & _
             "          A.CO_TEXT_XML,                  " & vbNewLine & _
             "          B.TP_OPER,                      " & vbNewLine & _
             "          B.TP_CNAL_VEND                  " & vbNewLine & _
             " From     A8.TB_OPER_ATIV_MESG_INTE   A,  " & vbNewLine & _
             "          A8.TB_OPER_ATIV             B   " & vbNewLine & _
             " Where    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV         " & vbNewLine & _
             " And      B.CO_OPER_ATIV          =   '" & pstrCodigoOperacao & "'" & vbNewLine & _
             " And      B.SG_SIST               =   '" & pstrSiglaSistema & "'  " & vbNewLine & _
             " And      A.TP_SOLI_MESG_INTE     =    " & plngTipoSolicitacao & vbNewLine

    If Trim(pstrIdentificadorTitulo) <> vbNullString Then
        strSQL = strSQL & _
                 " And      B.NU_ATIV_MERC          =   '" & pstrIdentificadorTitulo & "'" & vbNewLine
    End If

    If Val(pstrDataVencimento) > 0 Then
        strSQL = strSQL & _
                 " And      B.DT_VENC_ATIV          =    " & fgDtXML_To_Oracle(pstrDataVencimento) & vbNewLine
    End If

    If plngTipoOperacao > 0 Then
        strSQL = strSQL & _
                 " And      B.TP_OPER               =    " & plngTipoOperacao & vbNewLine
    End If

    If Trim(pstrIdentificadorLastro) <> vbNullString Then
        strSQL = strSQL & _
                 " And      B.CO_IDEF_LAST          =   '" & pstrIdentificadorLastro & "'" & vbNewLine
    End If

    strSQL = strSQL & _
             " Order By A.DH_MESG_INTE, A.TP_SOLI_MESG_INTE " & vbNewLine
             
    If pblnForUpdate And Not fgEstacaoDesenvolvimento Then
        strSQL = strSQL & _
                 " For Update "
    End If

    Set LerPorSolicitacao = fgQuerySQL(strSQL)

    If Not objContext Is Nothing Then
        'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerPorSolicitacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Obter o status de uma operação invertida
'-------------------------------------------------------------------------------------------------

Public Function ObterStatusOperInvertida(ByVal pstrCodigoEmpresa As String, _
                                         ByVal pstrDataOperacao As String, _
                                         ByVal pstrComandoOperacao As String, _
                                         ByVal pstrIdentificadorTitulo As String, _
                                         ByVal pstrDataVencimento As String, _
                                         ByVal pvntSequenciaOperacao As Variant, _
                                         ByRef pvntSequenciaOperacaoP2 As Variant) As Long

Dim strSQL                                  As String
Dim rsOperacao                              As ADODB.Recordset

    On Error GoTo ErrorHandler

    strSQL = "Select    NU_SEQU_OPER_ATIV, CO_ULTI_SITU_PROC  " & vbNewLine & _
             " From     A8.TB_OPER_ATIV             " & vbNewLine & _
             " Where    CO_EMPR               =     " & pstrCodigoEmpresa & vbNewLine & _
             " And      DT_OPER_ATIV          =     " & fgDtXML_To_Oracle(pstrDataOperacao) & vbNewLine & _
             " And      NU_COMD_OPER          =    '" & pstrComandoOperacao & "'" & vbNewLine & _
             " And      NU_ATIV_MERC          =    '" & pstrIdentificadorTitulo & "'" & vbNewLine & _
             " And      DT_VENC_ATIV          =     " & fgDtXML_To_Oracle(pstrDataVencimento) & vbNewLine & _
             " And      NU_SEQU_OPER_ATIV    <>     " & pvntSequenciaOperacao

    Set rsOperacao = fgQuerySQL(strSQL)

    If rsOperacao.RecordCount > 0 Then
        ObterStatusOperInvertida = rsOperacao!CO_ULTI_SITU_PROC
        pvntSequenciaOperacaoP2 = rsOperacao!NU_SEQU_OPER_ATIV
    Else
        ObterStatusOperInvertida = 0
        pvntSequenciaOperacaoP2 = 0
    End If

    If Not objContext Is Nothing Then
        'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterStatusOperInvertida Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Alterar o valor da operção
'-------------------------------------------------------------------------------------------------

Public Function AjustarValor(ByVal pvntSequenciaOperacao As Variant, _
                             ByVal pstrDataUltimaAtualizacao As String, _
                             ByVal pvntValorNovo As Variant, _
                             ByVal pstrJustificativa As String) As Boolean

    On Error GoTo ErrorHandler
    
    '///////////////////////////////////////////////////////////////////////////////////////////////
    '//
    '// ALTERA o valor da operação.
    '//
    '// OBS: o campo VA_OPER_ATIV_REAJ recebe o valor 'original' da operação, e o valor ajustado
    '//      fica no campo VA_OPER_ATIV mesmo.
    '//
    '///////////////////////////////////////////////////////////////////////////////////////////////

    Dim xmlOp                       As MSXML2.DOMDocument40
    Dim strSQL                      As String
    Dim objHistoricoOperacao        As A8LQS.clsHistSituacaoOperacao
    Dim objOperacao                 As A8LQS.clsOperacao
    
    Set xmlOp = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    xmlOp.loadXML objOperacao.Ler(pvntSequenciaOperacao)
    
    'Altera os campos
    strSQL = " Update   A8.TB_OPER_ATIV " & vbCrLf & _
             " Set      VA_OPER_ATIV        = " & fgVlrToDBServer(pvntValorNovo) & "," & vbCrLf & _
             "          VA_OPER_ATIV_REAJ   = DECODE(VA_OPER_ATIV_REAJ, NULL, VA_OPER_ATIV, VA_OPER_ATIV_REAJ) " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV   = " & pvntSequenciaOperacao & vbCrLf & _
             " AND      DH_ULTI_ATLZ        = " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao)

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    
    Call objHistoricoOperacao.Incluir(pvntSequenciaOperacao, _
                                        xmlOp.selectSingleNode("//CO_ULTI_SITU_PROC").Text, _
                                        , _
                                        enumTipoAcao.AjusteValor, _
                                        ("Valor anterior: " & fgVlrXml_To_Interface(xmlOp.selectSingleNode("//VA_OPER_ATIV").Text) & ", Justif: " & Replace(pstrJustificativa, "'", "")))

    Set objOperacao = Nothing
    Set objHistoricoOperacao = Nothing
    
    AjustarValor = True

Exit Function
ErrorHandler:
    
    Set objOperacao = Nothing
    Set objHistoricoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AjustarValor Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Atualizar Codigo Associacao Cambio
'-------------------------------------------------------------------------------------------------

Public Function AtualizarChACAM(ByVal pvntSequenciaOperacao As Variant, _
                                ByVal pvntChACAM As Variant) As Boolean

    On Error GoTo ErrorHandler

    Dim strSQL                      As String


    'Altera campo ChACAM
    strSQL = " Update   A8.TB_OPER_ATIV " & vbCrLf & _
             " Set      CD_ASSO_CAMB        = '" & Trim(pvntChACAM) & "'" & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV   = " & pvntSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operacao original nao localizada no SLCC.
        lngCodigoErroNegocio = 3120
        GoTo ErrorHandler
    End If

    AtualizarChACAM = True

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AtualizarChACAM Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Alterar o valor das operações para liquidação multilateral BMC

Public Function AlterarValorMoedasBMC(ByVal plngSequenciaOperacao As Variant, _
                                      ByVal pstrTipoAlteracao As String, _
                                      ByVal pstrNovoValor As String) As Boolean

Dim strSQL                                  As String
Dim strCampoValor                           As String

    On Error GoTo ErrorHandler

    strCampoValor = IIf(pstrTipoAlteracao = "MN", "VA_OPER_ATIV", "VA_MOED_ESTR")
    
    strSQL = " UPDATE   A8.TB_OPER_ATIV              " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          DH_ULTI_ATLZ            =   Sysdate, " & vbCrLf & _
             "          " & strCampoValor & "   =    " & fgVlrToDBServer(pstrNovoValor) & vbCrLf & _
             " WHERE    NU_SEQU_OPER_ATIV       =    " & plngSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    AlterarValorMoedasBMC = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarValorMoedasBMC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar o numero de comando da operação

Public Function AlterarComandoOperacao(ByVal plngSequenciaOperacao As Variant, _
                                       ByVal pstrNroComandoOperacao As String, _
                              Optional ByVal plngCodigoOperacaoCetip As Long = 0, _
                              Optional ByVal plngTipoLiquidacao As Long = 0) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " UPDATE   A8.TB_OPER_ATIV           " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          DH_ULTI_ATLZ            =   Sysdate, " & vbCrLf & _
             "          NU_COMD_OPER            =    '" & Trim(pstrNroComandoOperacao) & "'" & vbCrLf

    If plngCodigoOperacaoCetip > 0 Then
        strSQL = strSQL & " , CO_OPER_CETIP      =    " & plngCodigoOperacaoCetip & vbCrLf
    End If

    If plngTipoLiquidacao > 0 Then
        strSQL = strSQL & " , TP_LIQU_OPER_ATIV  =    " & plngTipoLiquidacao & vbCrLf
    End If

    strSQL = strSQL & " Where    NU_SEQU_OPER_ATIV       =    " & plngSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    AlterarComandoOperacao = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarComandoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Criar no xml da operação algumas TAGs epscíficas para operação Leilão BMA
'-------------------------------------------------------------------------------------------------

Public Function CamposExtrasLeilaoBMA(ByVal pvntSequenciaOperacao As Variant, _
                                      ByRef pxmlCampos As MSXML2.DOMDocument40) As String

    On Error GoTo ErrorHandler

    'Leilao BMA - Alguns campos que não estão na remessa devem ser usados para o preenchimento da LDL1002
    'Esta rotina carrega um xml com estes campos
    
    Dim strAux
    Dim objVL                                   As A6A7A8.clsVeiculoLegal
    Dim xmlVL                                   As MSXML2.DOMDocument40
    Dim xmlOperacao                             As MSXML2.DOMDocument40
    
    If pxmlCampos Is Nothing Then
        Set pxmlCampos = CreateObject("MSXML2.DOMDocument.4.0")
    End If
    
    Set objVL = CreateObject("A6A7A8.clsVeiculoLegal")
    Set xmlVL = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlOperacao.loadXML Ler(pvntSequenciaOperacao)
    
    strAux = objVL.Ler(fgSelectSingleNodeText(xmlOperacao, "//CO_VEIC_LEGA"), fgSelectSingleNodeText(xmlOperacao, "//SG_SIST"))
                            
    xmlVL.loadXML strAux
    
    pxmlCampos.loadXML ""
    fgAppendNode pxmlCampos, "", "Campos", ""
    fgAppendNode pxmlCampos, "Campos", "TP_ATIV", "4"        'TpAtv - 4 - Leilao
    fgAppendNode pxmlCampos, "Campos", "TP_REQU", "2"        'TpReq - 2 - Voluntaria"
    fgAppendNode pxmlCampos, "Campos", "TP_TITL_ORIG", fgSelectSingleNodeText(xmlVL, "//TP_TITL_BMA")
    fgAppendNode pxmlCampos, "Campos", "CO_TITL_ORIG", fgSelectSingleNodeText(xmlVL, "//CO_TITL_BMA")
    fgAppendNode pxmlCampos, "Campos", "CO_FIND_COBE", "14"

    CamposExtrasLeilaoBMA = pxmlCampos.xml

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "CamposExtrasLeilaoBMA Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Obter o número de comando da especificação
'-------------------------------------------------------------------------------------------------

Public Function ObterComandoEspecificacao(ByRef xmlOperacao As MSXML2.DOMDocument40) As Long

Dim strSQL                                  As String
Dim strWhereTipoOper                        As String
Dim rsQuery                                 As ADODB.Recordset
Dim strDHInicio                             As String
Dim strDHFim                                As String
                                                         
On Error GoTo ErrorHandler
    
    strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")

    If CLng(xmlOperacao.documentElement.selectSingleNode("//TP_ESFC").Text) = enumTipoEspecificacao.Cobertura Or _
       CLng(xmlOperacao.documentElement.selectSingleNode("//TP_ESFC").Text) = enumTipoEspecificacao.Intermediacao Then

        strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
                 "        TP_OPER,               " & _
                 "        CO_ULTI_SITU_PROC,     " & _
                 "        QT_ATIV_MERC           " & vbCrLf & _
                 " FROM   A8.TB_OPER_ATIV        " & vbCrLf & _
                 " WHERE  NU_COMD_OPER = '" & xmlOperacao.documentElement.selectSingleNode("//NU_COMD_OPER").Text & "'" & vbCrLf & _
                 "   AND  CO_EMPR      =  " & xmlOperacao.documentElement.selectSingleNode("//CO_EMPR").Text & vbCrLf & _
                 "   AND  CO_VEIC_LEGA = '" & xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text & "'" & vbCrLf & _
                 "   AND  DT_OPER_ATIV BETWEEN " & strDHInicio & _
                 "   AND " & strDHFim
        
        strWhereTipoOper = " AND TP_OPER IN ( " & enumTipoOperacaoLQS.DefinitivaCobertaBMA & " , " & _
                                                  enumTipoOperacaoLQS.DefinitivaDescobertaBMA & " , " & _
                                                  enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA & " , " & _
                                                  enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA & " , " & _
                                                  enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA & " , " & _
                                                  enumTipoOperacaoLQS.OperacaoTermoCobertaBMA & ", " & _
                                                  enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA & ", " & _
                                                  enumTipoOperacaoLQS.CompromissadaGenericaAVista & ") "
        
        strSQL = strSQL & strWhereTipoOper
        
        strSQL = strSQL & " ORDER BY DH_ULTI_ATLZ DESC"
        
        Set rsQuery = fgQuerySQL(strSQL)

        If Not rsQuery.EOF Then
            If Trim(xmlOperacao.documentElement.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text) = vbNullString Then
                xmlOperacao.documentElement.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text = rsQuery!NU_CTRL_MESG_SPB_ORIG
            End If
        End If

    End If

    Set rsQuery = Nothing

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterComandoEspecificacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Atualiza o número de protocolo retornado da verificação de alçada.

Public Function AlterarNumeroProtocoloLG(ByVal plngSequenciaOperacao As Variant, _
                                         ByVal pvntProtocoloLG As Variant) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " UPDATE   A8.TB_OPER_ATIV            " & vbCrLf & _
             " SET      CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "', " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "          NU_PRTC_OPER_LG         = '" & Trim(pvntProtocoloLG) & "'" & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV       =  " & plngSequenciaOperacao

    If fgExecuteSQL(strSQL) = 0 Then
        'Operação já processada por outro usuário.
        lngCodigoErroNegocio = 3013
        GoTo ErrorHandler
    End If

    AlterarNumeroProtocoloLG = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarNumeroProtocoloLG Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Verificar se existe um operação de especificação
Public Function ExisteOperacaoEspecificacao(ByRef pxmlremessa As MSXML2.DOMDocument40) As Boolean
                                                         
Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim strWhere                                As String
Dim strDHInicio                             As String
Dim strDHFim                                As String
Dim strNumeroOperNegoBMA                    As String
Dim lngCodigoEmpresa                        As Long
Dim strCodigoVeiculoLegal                   As String
                                                         
On Error GoTo ErrorHandler
    
    strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
    
    strNumeroOperNegoBMA = pxmlremessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text
    lngCodigoEmpresa = pxmlremessa.documentElement.selectSingleNode("//CO_EMPR").Text
    strCodigoVeiculoLegal = pxmlremessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text
    
    strSQL = " SELECT   NU_CTRL_MESG_SPB_ORIG " & vbCrLf & _
             "   FROM   A8.TB_OPER_ATIV " & vbCrLf & _
             "  WHERE   NU_COMD_OPER = '" & strNumeroOperNegoBMA & "'" & vbCrLf & _
             "    AND   CO_EMPR      =  " & lngCodigoEmpresa & vbCrLf & _
             "    AND   CO_VEIC_LEGA = '" & strCodigoVeiculoLegal & "'" & vbCrLf & _
             "    AND   TP_OPER IN ( " & enumTipoOperacaoLQS.EspecDefinitivaIntermediacao & ", " & vbCrLf & _
                                         enumTipoOperacaoLQS.EspecDefinitivaCobertura & ", " & vbCrLf & _
                                         enumTipoOperacaoLQS.EspecTermoIntermediacao & ", " & vbCrLf & _
                                         enumTipoOperacaoLQS.EspecTermoCobertura & ", " & vbCrLf & _
                                         enumTipoOperacaoLQS.EspecCompromissadaIntermediacao & ", " & vbCrLf & _
                                         enumTipoOperacaoLQS.EspecCompromissadaCobertura & ")" & vbCrLf & _
            "  AND  CO_ULTI_SITU_PROC NOT IN (" & enumStatusOperacao.CanceladaOrigem & ")" & vbCrLf & _
            "  AND  DT_OPER_ATIV BETWEEN " & strDHInicio & vbCrLf & _
            "  AND " & strDHFim

    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
       ExisteOperacaoEspecificacao = False
    Else
       ExisteOperacaoEspecificacao = True
    End If

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExisteOperacaoEspecificacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Obter as operações que compuseram o lançamento consolidado de conta corrente
'-------------------------------------------------------------------------------------------------

Public Function flObterComposicaoNet(ByVal pvntSequenciaOperacao As Variant) As String

Dim strSQL                                  As String
Dim xmlRetorno                              As MSXML2.DOMDocument40
    
    On Error GoTo ErrorHandler
    
    strSQL = " SELECT   DISTINCT A.NU_SEQU_OPER_ATIV,                       " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                     " & vbNewLine & _
             "          DECODE(A.IN_ENTR_SAID_RECU_FINC, " & enumTipoEntradaSaida.ENTRADA & ", 'Entrada', " & _
             "                                           " & enumTipoEntradaSaida.Saida & ", 'Saída') " & _
             "          IN_ENTR_SAID_RECU_FINC,                             " & vbNewLine & _
             "          A.VA_OPER_ATIV,                                       " & vbNewLine & _
             "          A.DT_OPER_ATIV,                                     " & vbNewLine & _
             "          C.SG_LOCA_LIQU                                      " & vbNewLine & _
             " FROM     A8.TB_OPER_ATIV A,                                  " & vbNewLine & _
             "          A8.TB_VEIC_LEGA B,                                  " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO C,                           " & vbNewLine & _
             "          A8.TB_CNCL_OPER_LANC_CNTA_CRRT D                    " & vbNewLine & _
             " WHERE    D.NR_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & " " & vbNewLine & _
             " AND     (A.NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & " " & vbNewLine & _
             " OR       A.NU_SEQU_OPER_ATIV = D.NR_SEQU_LANC_CNTA_CRRT)     " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA      = B.CO_VEIC_LEGA                " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU      = C.CO_LOCA_LIQU                " & vbNewLine & _
             " AND      C.CO_EMPR_FUSI      = 1                             " & vbNewLine & _
             " ORDER BY A.NU_SEQU_OPER_ATIV "

    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlRetorno.loadXML(fgQueryXMLLerTodos("DetalheOperacao", strSQL, App.EXEName & "." & TypeName(Me)))

    If xmlRetorno.selectNodes("//Repeat_DetalheOperacao/*").length > 1 Then
        flObterComposicaoNet = xmlRetorno.xml
    Else
        flObterComposicaoNet = vbNullString
    End If
    
    Set xmlRetorno = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlRetorno = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterComposicaoNet Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
