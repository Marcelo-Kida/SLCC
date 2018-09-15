VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessamentoDiretoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
Option Explicit

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

Private Sub flAjustarValoresTagsEntrada(ByRef xmlRemessa As MSXML2.DOMDocument40)

    On Error GoTo ErrorHandler

    If Not xmlRemessa.selectSingleNode("//NU_COMD_OPER") Is Nothing Then
        If IsNumeric(xmlRemessa.selectSingleNode("//NU_COMD_OPER").Text) Then
            xmlRemessa.selectSingleNode("//NU_COMD_OPER").Text = Val(xmlRemessa.selectSingleNode("//NU_COMD_OPER").Text)
        End If
    End If

    If Not xmlRemessa.selectSingleNode("//TP_TRAF_LDL") Is Nothing Then
        If IsNumeric(xmlRemessa.selectSingleNode("//TP_TRAF_LDL").Text) Then
            xmlRemessa.selectSingleNode("//TP_TRAF_LDL").Text = Format$(xmlRemessa.selectSingleNode("//TP_TRAF_LDL").Text, "00")
        End If
    End If
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAjustarValoresTagsEntrada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub flGerenciarEnvioMensagens(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                      ByVal intFuncionalidade As enumFuncionalidade)

Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    
    Select Case intFuncionalidade
        Case enumFuncionalidade.RecebimentoBoleto
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlRemessa, enumTipoMovimentoPJ.Previsto)
            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlRemessa, enumTipoMovimento.Previsto)
            
    End Select
    
    Set objFluxoProcessamento = Nothing
    
    Exit Sub

ErrorHandler:
    Set objFluxoProcessamento = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarEnvioMensagens Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub flIncluirOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40)

Dim objOperacao                             As A8LQS.clsOperacao

    Call fgAppendNode(xmlRemessa, "MESG", "PROCESSAMENTO_DIRETO", vbNullString)
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.IncluirOperacao(xmlRemessa, enumStatusOperacao.EmSer)
    Set objOperacao = Nothing
    
    Exit Sub

ErrorHandler:
    Set objOperacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub flJustificarProcessoNaoAutomatico(ByRef xmlRemessa As MSXML2.DOMDocument40)

Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao

    Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    Call objHistSituacaoOperacao.AlterarJustificativa(xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                      xmlRemessa.selectSingleNode("//CO_ULTI_SITU_PROC").Text, _
                                                      enumJustificativa.RegraWorkflow)
    Set objHistSituacaoOperacao = Nothing
    
    Exit Sub

ErrorHandler:
    Set objHistSituacaoOperacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flJustificarProcessoNaoAutomatico Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub flObterDadosTipoOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40)

    On Error GoTo ErrorHandler

    If Not xmlRemessa.selectSingleNode("//TP_VINC_DSVN_TRAF") Is Nothing Then
        If xmlRemessa.selectSingleNode("//TP_VINC_DSVN_TRAF").Text = "4" Then
            Call fgAppendNode(xmlRemessa, "MESG", "TP_OPER", enumTipoOperacaoLQS.TransferenciaSemFinanceiro)
            Call fgAppendNode(xmlRemessa, "MESG", "CO_MESG", "SEL1063")
            Call fgAppendNode(xmlRemessa, "MESG", "TP_MESG_RETN_INTE", "14")
        End If
    End If
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDadosTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Public Function GerenciarRecebimentoOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                    Optional ByRef strError As String) As Boolean

    On Error GoTo ErrorHandler

    strError = vbNullString
    
    Call flAjustarValoresTagsEntrada(xmlRemessa)
    Call flObterDadosTipoOperacao(xmlRemessa)
    Call flIncluirOperacao(xmlRemessa)
    Call flGerenciarEnvioMensagens(xmlRemessa, enumFuncionalidade.RecebimentoBoleto)
    Call flJustificarProcessoNaoAutomatico(xmlRemessa)
    
    GerenciarRecebimentoOperacao = True

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
