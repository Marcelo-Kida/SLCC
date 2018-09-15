VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTContaCorrente"
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
                          ByVal pintTipoLancamentoIntegracao As Integer, _
                          ByVal pintStatusIntegracao As enumStatusIntegracao, _
                          ByVal pstrDHUltimaAtualizacao As String, _
                          ByVal pstrJustificativaCancel As String, _
                          ByVal pblnIgnoraGradeHorario As Boolean, _
                 Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = Normal, _
                 Optional ByVal pstrJustificativa As String = vbNullString, _
                 Optional ByVal pstrVaLancCC As String = vbNullString, _
                 Optional ByVal pstrInLancDebCred As String = vbNullString, _
                 Optional ByVal pstrNetOperacoes As String = vbNullString, _
                 Optional ByVal pNumSeqLancamento As Long = 0) As String
    
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument
    
    On Error GoTo ErrorHandler
    
    If pintStatusIntegracao = enumStatusIntegracao.IntegradoCC Then
        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
        Call objContaCorrente.Integrar(pvntSequenciaOperacao, _
                                       enumIntegrarCCContabilidade.ContaCorrente, _
                                       pblnIgnoraGradeHorario, _
                                       plngTipoLancamentoIntegracao, _
                                       pstrJustificativa, _
                                       pstrVaLancCC, _
                                       pstrInLancDebCred, _
                                       pstrNetOperacoes, _
                                       pNumSeqLancamento)
    
    ElseIf pintStatusIntegracao = enumStatusIntegracao.Integrado Then
        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
        Call objContaCorrente.Integrar(pvntSequenciaOperacao, _
                                       enumIntegrarCCContabilidade.Contabilidade, _
                                       pblnIgnoraGradeHorario, _
                                       plngTipoLancamentoIntegracao, _
                                       pstrJustificativa, _
                                       pstrVaLancCC, _
                                       pstrInLancDebCred, _
                                       pstrNetOperacoes, _
                                       pNumSeqLancamento)
    
    Else
        Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
        Call objLancamentoCCContabil.AlterarStatus(pvntSequenciaOperacao, _
                                                   pintTipoLancamentoIntegracao, _
                                                   pintStatusIntegracao, _
                                                   pstrDHUltimaAtualizacao, _
                                                   pstrJustificativaCancel)
                                                   
                                                   
        If pintStatusIntegracao = enumStatusIntegracao.Cancelado Then
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = New MSXML2.DOMDocument
            
            xmlOperacao.Load objOperacao.LerXMLOperacao(pvntSequenciaOperacao)
            
            '********************************************************************************************
            '  RATS 1008 - 02/08/2010 - Enviar o retorno de lancamento CC cancelado para o legado
            '********************************************************************************************
            Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
        
            If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
                Call fgAppendNode(xmlOperacao, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
            End If
        
            xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.Cancelado
        
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao, False)
            '********************************************************************************************

        End If

    End If
    
    Set objContaCorrente = Nothing
    Set objLancamentoCCContabil = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing

    Gerenciar = vbNullString        '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objContaCorrente = Nothing
    Set objLancamentoCCContabil = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    Gerenciar = Err.Description
    
End Function
