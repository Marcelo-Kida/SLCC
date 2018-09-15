VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCadastroWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao Cadastro de Parametrização de Workflow.

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

' Este método aciona a leitura de um registro do cadastro de parametrização de workflow na camada de negócios,
' e retorna para a camada de interface.

Public Function Ler(ByVal pintTipoOperacao As Integer, _
                    ByVal pintFuncaoSistema As Integer, _
                    ByVal pintBancoLiquidante As Integer, _
                    ByRef pvntCodErro As Variant, _
                    ByRef pvntMensagemErro As Variant) As String

Dim objCadastroWorkflow                     As A8LQS.clsCadastroWorkflow

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objCadastroWorkflow = CreateObject("A8LQS.clsCadastroWorkflow")
    Ler = objCadastroWorkflow.Ler(pintTipoOperacao, _
                                  pintFuncaoSistema, _
                                  pintBancoLiquidante)
    Set objCadastroWorkflow = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objCadastroWorkflow = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura geral do cadastro de parametrização de workflow na camada de negócios,
' e retorna para a camada de interface.

Public Function LerTodos(ByVal pintBancoLiquidante As Integer, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String

Dim objCadastroWorkflow                     As A8LQS.clsCadastroWorkflow

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objCadastroWorkflow = CreateObject("A8LQS.clsCadastroWorkflow")
    LerTodos = objCadastroWorkflow.LerTodos(pintBancoLiquidante)
    Set objCadastroWorkflow = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objCadastroWorkflow = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura das excessões às regras de parametrização de workflow na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterExcecoesParametrizacao(ByVal pintTipoOperacao As Integer, _
                                            ByVal pintFuncaoSistema As Integer, _
                                            ByVal pintBancoLiquidante As Integer, _
                                            ByRef pvntCodErro As Variant, _
                                            ByRef pvntMensagemErro As Variant) As String

Dim objCadastroWorkflow                     As A8LQS.clsCadastroWorkflow

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objCadastroWorkflow = CreateObject("A8LQS.clsCadastroWorkflow")
    ObterExcecoesParametrizacao = objCadastroWorkflow.ObterExcecoesParametrizacao(pintTipoOperacao, _
                                                                                  pintFuncaoSistema, _
                                                                                  pintBancoLiquidante)
    Set objCadastroWorkflow = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objCadastroWorkflow = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterExcecoesParametrizacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
