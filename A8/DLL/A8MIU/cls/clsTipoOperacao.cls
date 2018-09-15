VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao Cadastro de Tipos de Operação Existentes.
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

' Este método aciona a leitura de tipos de operação para conciliação, de acordo com o local de liquidação,
' na camada de negócios, e retorna para a camada de interface.

Public Function ObterTiposOperacaoConciliacao(ByVal plngLocalLiquidacao As Long, _
                                              ByRef pvntCodErro As Variant, _
                                              ByRef pvntMensagemErro As Variant) As String

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    ObterTiposOperacaoConciliacao = objTipoOperacao.ObterTiposOperacaoConciliacao(plngLocalLiquidacao)
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
Exit Function
ErrorHandler:
    Set objTipoOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTiposOperacaoConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de controles de previsões
' na camada de negócios, e retorna para a camada de interface.

Public Function ObterControlePrevisoes() As String

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
    
    On Error GoTo ErrorHandler
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    ObterControlePrevisoes = objTipoOperacao.ObterControlePrevisoes()
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterControlePrevisoes Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a leitura dos tipos de operação existentes,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerTodos() As String

    Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
    
    On Error GoTo ErrorHandler
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    LerTodos = objTipoOperacao.LerTodos
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a leitura de processos de operação,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerProcessosOperacao() As String

    Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
    
    On Error GoTo ErrorHandler
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    LerProcessosOperacao = objTipoOperacao.LerProcessosOperacao
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerProcessosOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
