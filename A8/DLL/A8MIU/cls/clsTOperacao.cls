VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao tratamento de Operações de Sistemas Legados.
' É responsável também, pela ligação da camada intermediária, à camada de negócios.

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

' Este método aciona a confirmação de operações, na camada de negócios.

Public Function Confirmar(ByVal pvntSequenciaOperacao As Variant, _
                          ByVal pintStatusOperacao As enumStatusOperacao, _
                          ByVal pstrDHUltimaAtualizacao As String, _
                          ByVal pvntProtocolo As Variant) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.Confirmar(pvntSequenciaOperacao, pintStatusOperacao, pstrDHUltimaAtualizacao, pvntProtocolo)
    Set objOperacao = Nothing

    Confirmar = vbNullString        '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    Confirmar = Err.Description
    
End Function

' Este método aciona a liquidação de operações via contingência, na camada de negócios.

Public Function LiquidarViaContingencia(ByVal pvntSequenciaOperacao As Variant, _
                                        ByVal pintStatusOperacao As enumStatusOperacao, _
                                        ByVal pstrDHUltimaAtualizacao As String) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.LiquidarViaContingencia(pvntSequenciaOperacao, pstrDHUltimaAtualizacao)
    Set objOperacao = Nothing

    LiquidarViaContingencia = vbNullString       '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarViaContingencia = Err.Description
    
End Function

' Este método aciona a baixa de operações via contingência, na camada de negócios.

Public Function BaixarViaContingencia(ByVal pvntSequenciaOperacao As Variant, _
                                      ByVal pintStatusOperacao As enumStatusOperacao, _
                                      ByVal pstrDHUltimaAtualizacao As String) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.BaixarViaContingencia(pvntSequenciaOperacao, pstrDHUltimaAtualizacao)
    Set objOperacao = Nothing

    BaixarViaContingencia = vbNullString       '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    BaixarViaContingencia = Err.Description
    
End Function

' Este método aciona a liberação de operações, na camada de negócios.

Public Function Liberar(ByVal pvntSequenciaOperacao As Variant, _
                        ByVal pintStatusOperacao As enumStatusOperacao, _
                        ByVal pintAcao As enumTipoAcao, _
                        ByVal plngNumeroComandoAcao As Long, _
                        ByVal pstrDHUltimaAtualizacao As String, _
                        ByVal pblnIgnoraGradeHorario As Boolean) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.Liberar(pvntSequenciaOperacao, _
                             pintStatusOperacao, _
                             pintAcao, _
                             plngNumeroComandoAcao, _
                             pstrDHUltimaAtualizacao, _
                             pblnIgnoraGradeHorario)
    Set objOperacao = Nothing

    Liberar = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    Liberar = Err.Description
    
End Function

' Este método aciona o cancelamento de operações enviadas manualmente, na camada de negócios.

Public Function CancelarEntradaManual(ByVal pvntSequenciaOperacao As Variant, _
                                      ByVal pintStatus As Integer, _
                                      ByVal pstrDHUltimaAtualizacao As String) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.CancelarEntradaManual(pvntSequenciaOperacao, pintStatus, pstrDHUltimaAtualizacao)
    Set objOperacao = Nothing

    CancelarEntradaManual = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    CancelarEntradaManual = Err.Description
    
End Function

' Este método aciona a alteração de agendamento de operações, na camada de negócios.

Public Function AlterarAgendamento(ByVal pvntSequenciaOperacao As Variant, _
                                   ByVal pstrHorarioAgendamento As String, _
                                   ByVal pstrDHUltimaAtualizacao As String, _
                                   ByVal plngLocalLiquidacao As Long, _
                                   ByVal pstrCodigoMensagem As String, _
                                   ByVal pintStatusOperacao As enumStatusOperacao)

Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.AlterarAgendamento(pvntSequenciaOperacao, _
                                        pstrHorarioAgendamento, _
                                        pstrDHUltimaAtualizacao, _
                                        plngLocalLiquidacao, _
                                        pstrCodigoMensagem, _
                                        pintStatusOperacao)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a alteração de tipos de compromissos, na camada de negócios.

Public Function AlterarTipoCompromisso(ByVal plngSequenciaOperacao As Long, _
                                       ByVal pstrDataUltimaAtualizacao As String, _
                                       ByVal pstrTipoCompromisso As String, _
                                       ByVal pstrTipoCompromissoRet As String, _
                                       ByVal pintStatusOperacao As String, _
                                       ByVal pstrAcaoAnterior As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.AlterarTipoCompromisso(plngSequenciaOperacao, _
                                            pstrDataUltimaAtualizacao, _
                                            pstrTipoCompromisso, _
                                            pstrTipoCompromissoRet, _
                                            pintStatusOperacao, _
                                            pstrAcaoAnterior)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarTipoCompromisso Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona o cancelamento de operações, na camada de negócios.

Public Function Cancelar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pintStatusOperacao As Integer, _
                         ByVal plngNumeroComando As Long, _
                         ByVal pstrDataUltimaAtualizacao As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Cancelar = objOperacao.Cancelar(pvntSequenciaOperacao, pintStatusOperacao, plngNumeroComando, pstrDataUltimaAtualizacao)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Cancelar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona o estorno de operações, na camada de negócios.

Public Function Estornar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pintStatusOperacao As Integer, _
                         ByVal plngNumeroComando As Long, _
                         ByVal pstrDataUltimaAtualizacao As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Estornar = objOperacao.Estornar(pvntSequenciaOperacao, pintStatusOperacao, plngNumeroComando, pstrDataUltimaAtualizacao)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Estornar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona o reenvio de operações, na camada de negócios.

Public Function Reenviar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pstrDHUltimaAtualizacao As String) As Boolean
                         
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Reenviar = objOperacao.Reenviar(pvntSequenciaOperacao, pstrDHUltimaAtualizacao)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Reenviar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a liberação de registro de operações, na camada de negócios.

Public Function LiberarRegistro(ByVal plngLocalLiquidacao As Long, _
                                ByVal pvntSequenciaOperacao As Variant, _
                                ByVal pintTipoOperacao As Long, _
                                ByVal pstrDHUltimaAtualizacao As String, _
                                ByVal pblnIgnoraGradeHorario As Boolean) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Call objOperacao.LiberarRegistro(plngLocalLiquidacao, pvntSequenciaOperacao, pintTipoOperacao, _
                             pstrDHUltimaAtualizacao, _
                             pblnIgnoraGradeHorario, False)
    Set objOperacao = Nothing

    LiberarRegistro = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiberarRegistro = Err.Description
    
End Function

' Este método aciona o ajuste do valor de operações, na camada de negócios.

Public Function AjustarValor(ByVal pvntSequenciaOperacao As Variant, _
                             ByVal pstrDataUltimaAtualizacao As String, _
                             ByVal pvntValorNovo As Variant, _
                             ByVal pstrJustificativa As String) As Boolean

    Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    AjustarValor = objOperacao.AjustarValor(pvntSequenciaOperacao, _
                                            pstrDataUltimaAtualizacao, _
                                            pvntValorNovo, _
                                            pstrJustificativa)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AjustarValor Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

