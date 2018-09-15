VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluxoProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pela disponibização do fluxo das mensagens

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

'--------------------------------------------------------------------------------------------------
'Disponibilizar lancamento conta corrente
'--------------------------------------------------------------------------------------------------

Public Function DisponibilizarLancamentoCC(ByVal xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objContaCorrente                        As A8LQS.clsContaCorrente

    On Error GoTo ErrorHandler

    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    Call objContaCorrente.Disponibilizar(xmlOperacao)
    Set objContaCorrente = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objContaCorrente = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "DisponibilizarLancamentoCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Enviar Mensagem retorno sistema legado
'--------------------------------------------------------------------------------------------------

Public Function EnviarMensagemRetorno(ByVal xmlOperacao As MSXML2.DOMDocument40, _
                             Optional ByVal pblnForcarLeitura As Boolean = True) As Boolean

Dim objLegado                               As A8LQS.clsLegado
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacaoAtual                        As MSXML2.DOMDocument40

Dim strMensagem                             As String

    On Error GoTo ErrorHandler

    Set objLegado = CreateObject("A8LQS.clsLegado")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacaoAtual = CreateObject("MSXML2.DOMDocument.4.0")
    
    If pblnForcarLeitura Then
        Set xmlOperacaoAtual = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, , , , , , , , , , , False)
    Else
        Call xmlOperacaoAtual.loadXML(xmlOperacao.xml)
    End If
    
    strMensagem = objLegado.EnviarRetorno(xmlOperacaoAtual)
    Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, False, False)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacaoAtual = Nothing
    
    Exit Function

ErrorHandler:
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacaoAtual = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMensagemRetorno Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Enviar Movimento PJ
'--------------------------------------------------------------------------------------------------

Public Function EnviarMovimentoPJ(ByVal xmlOperacao As MSXML2.DOMDocument40, _
                                  ByVal intTipoMovimento As enumTipoMovimentoPJ) As Boolean

Dim strMensagem                             As String
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

    On Error GoTo ErrorHandler

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    If intTipoMovimento = enumTipoMovimentoPJ.Realizado Then
        strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
    
    ElseIf intTipoMovimento = enumTipoMovimentoPJ.EstornoRealizado Then
        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
    
    ElseIf intTipoMovimento = enumTipoMovimentoPJ.EstornoPrevisto Then
        strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
    
    ElseIf intTipoMovimento = enumTipoMovimentoPJ.Previsto Then
        strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
    
    End If
    
    Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, True, False)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objGestaoCaixa = Nothing
    Set objProcessoOperacao = Nothing
    
    Exit Function

ErrorHandler:
    Set objGestaoCaixa = Nothing
    Set objProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMovimentoPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Enviar Movimento PJ - Moeda Estrangeira
'--------------------------------------------------------------------------------------------------

Public Function EnviarMovimentoPJME(ByVal xmlOperacao As MSXML2.DOMDocument40, _
                                    ByVal intTipoMovimento As enumTipoMovimentoPJ) As Boolean

Dim strMensagem                             As String
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

    On Error GoTo ErrorHandler

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        
    If intTipoMovimento = enumTipoMovimentoPJ.Previsto Then
        strMensagem = objGestaoCaixa.EnviarPrevisaoME(xmlOperacao)
        
    ElseIf intTipoMovimento = enumTipoMovimentoPJ.EstornoPrevisto Then
        strMensagem = objGestaoCaixa.EstornarPrevisaoME(xmlOperacao)
    
    End If
    
    Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, True, False)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objGestaoCaixa = Nothing
    Set objProcessoOperacao = Nothing
    
    Exit Function

ErrorHandler:
    Set objGestaoCaixa = Nothing
    Set objProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMovimentoPJME Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Enviar Movimento A6 Sub Reserva
'--------------------------------------------------------------------------------------------------

Public Function EnviarMovimentoSubReserva(ByVal xmlOperacao As MSXML2.DOMDocument40, _
                                          ByVal intTipoMovimento As enumTipoMovimento) As Boolean

Dim strMensagem                             As String
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

    On Error GoTo ErrorHandler

    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    If intTipoMovimento = enumTipoMovimento.RealizadoConfirmado Then
        strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlOperacao)
    
    ElseIf intTipoMovimento = enumTipoMovimento.RealizadoSolicitado Then
        strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlOperacao)
    
    ElseIf intTipoMovimento = enumTipoMovimento.EstornoPrevisto Then
        strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
    
    ElseIf intTipoMovimento = enumTipoMovimento.Previsto Then
        strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
    
    End If
    
    Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.SubReserva, True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objSubReserva = Nothing
    Set objProcessoOperacao = Nothing
    
    Exit Function

ErrorHandler:
    Set objSubReserva = Nothing
    Set objProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMovimentoSubReserva Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar registros do Fluxo de processamento de mensagens
Public Function LerTodos(Optional ByVal pstrCodigoMensagemRec As String = vbNullString, _
                         Optional ByVal pstrCodigoMensagemEnv As String = vbNullString, _
                         Optional ByVal plngFuncaoSistema As Long = 0) As String

Dim strSQL                                  As String
Dim blnWhere                                As Boolean

    On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_MESG_SPB_RECB, CO_MESG_SPB_ENVI, CO_FCAO_SIST " & vbNewLine & _
             " FROM     A8.TB_FLUX_PROC " & vbNewLine

    blnWhere = False
    If pstrCodigoMensagemRec <> vbNullString Then
        strSQL = strSQL & " WHERE   CO_MESG_SPB_RECB    =   '" & pstrCodigoMensagemRec & "'" & vbNewLine
        blnWhere = True
    End If

    If pstrCodigoMensagemEnv <> vbNullString Then
        If blnWhere Then
            strSQL = strSQL & " AND     CO_MESG_SPB_ENVI    =   '" & pstrCodigoMensagemEnv & "'" & vbNewLine
        Else
            strSQL = strSQL & " WHERE   CO_MESG_SPB_ENVI    =   '" & pstrCodigoMensagemEnv & "'" & vbNewLine
            blnWhere = True
        End If
    End If

    If plngFuncaoSistema > 0 Then
        If blnWhere Then
            strSQL = strSQL & " AND     CO_FCAO_SIST    =   " & plngFuncaoSistema & vbNewLine
        Else
            strSQL = strSQL & " WHERE   CO_FCAO_SIST    =   " & plngFuncaoSistema & vbNewLine
            blnWhere = True
        End If
    End If

    LerTodos = fgQueryXMLLerTodos("HistSituacaoIntegracao", strSQL, App.EXEName & "." & TypeName(Me))

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
