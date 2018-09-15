VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorReprocCCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao Controle e Fluxo de Conta Corrente / Contábil.
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

' Este método aciona a gerenciamento geral de lançamentos de conta corrente, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function Gerenciar(ByVal pvntSequenciaOperacao As Variant, _
                          ByVal pstrInLancDebCred As String, _
                 Optional ByVal pNumSeqLancamento As Long = 0) As String
    
Dim objContaCorrente                        As A8LQS.clsContaCorrente
    
On Error GoTo ErrorHandler
    
    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    
    Call objContaCorrente.Integrar(pvntSequenciaOperacao, _
                                   enumIntegrarCCContabilidade.ContaCorrente, _
                                   True, _
                                   enumTipoLancamentoIntegracao.Normal, _
                                   vbNullString, _
                                   vbNullString, _
                                   pstrInLancDebCred, _
                                   vbNullString, _
                                   pNumSeqLancamento, _
                                   True)
    
    Set objContaCorrente = Nothing

    Gerenciar = vbNullString        '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    
    Set objContaCorrente = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    Gerenciar = Err.Description
    
End Function


