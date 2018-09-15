VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConciliacaoLote"
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
Dim objTConciliacaoLote         As A8LQS.clsTConciliacaoLote
On Error GoTo ErrorHandler

    Set objTConciliacaoLote = CreateObject("A8LQS.clsTConciliacaoLote")
    LiberarLoteParaAlcada = objTConciliacaoLote.LiberarLoteParaAlcada
    Set objTConciliacaoLote = Nothing
   
    Exit Function
ErrorHandler:
    Set objTConciliacaoLote = Nothing
    
    fgGravaArquivo App.EXEName & "." & "LiberarLoteParaAlcada", Err.Description
    
End Function
Public Function AlterarStatusLote(ByVal plngIdLote As Long, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv As Date, ByVal pintSituacaoLote As Integer)
    Dim objTConciliacaoLote                     As A8LQS.clsTConciliacaoLote
    On Error GoTo ErrorHandler
    
    Set objTConciliacaoLote = CreateObject("A8LQS.clsTConciliacaoLote")
    AlterarStatusLote = objTConciliacaoLote.AlterarStatusLote(plngIdLote, pstrSiglaSistema, pdtDataOperAtiv, pintSituacaoLote)
    Set objTConciliacaoLote = Nothing
    
    Exit Function
ErrorHandler:
    Set objTConciliacaoLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
Public Function LiquidarLoteCorretoras(ByVal pintAcao As Integer, _
                                   ByVal strXMLOperacoes As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim objTConciliacaoLote                     As A8LQS.clsTConciliacaoLote

On Error GoTo ErrorHandler


    Set objTConciliacaoLote = CreateObject("A8LQS.clsTConciliacaoLote")
    LiquidarLoteCorretoras = objTConciliacaoLote.LiquidarLoteCorretoras(pintAcao, strXMLOperacoes, pvntCodErro, pvntMensagemErro)
    Set objTConciliacaoLote = Nothing

Exit Function
ErrorHandler:
    Set objTConciliacaoLote = Nothing
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
End Function

Public Function VerificarConciliacaoR2PagtoCorretora(ByVal pstrMensagemR2 As String, ByVal plngLocalLiquidacaoR2) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objclsTConciliacaoLote                  As A8LQS.clsTConciliacaoLote
Dim objRegra                                As A8LQS.clsRegra
Dim rsLotesALiberar                         As ADODB.Recordset
Dim rsOperacao                              As ADODB.Recordset
Dim udtSTR0008R2                            As udtSTR0008R2
Dim udtSTR0008R2Aux                         As udtSTR0008R2Aux

Dim strLoteISPBIFDebitada                   As String
Dim strLoteNomeRemetDebitado                As String
Dim strLoteCNPJCPFDebitado                  As String
Dim lngLoteAgenciaDebitada                  As Long
Dim lngLoteContaDebitada                    As Long
Dim lngLoteLocalLiquidacao                  As Long
Dim dblLoteValor                            As Double
Dim blnConciliacaoOK                        As Boolean
Dim dblMsgR2Valor                           As Double
Dim intAcaoProcessamento                    As Integer
Dim blnConfirmacaoAutomatica                As Boolean
Dim blnLiberacaoAutomatica                  As Boolean
Dim blnIntegracaoCCAutomatica               As Boolean
Dim intTipoCorretora                        As Integer

On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objclsTConciliacaoLote = CreateObject("A8LQS.clsTConciliacaoLote")
    Set objRegra = CreateObject("A8LQS.clsRegra")
    
    udtSTR0008R2Aux.String = pstrMensagemR2
    LSet udtSTR0008R2 = udtSTR0008R2Aux
    
    dblMsgR2Valor = Val(Left$(Trim$(udtSTR0008R2.VA_RECB_PGTO), 16)) & "," & Right$(Trim$(udtSTR0008R2.VA_RECB_PGTO), 2)
    
    'obtem lotes de recebimento que estão em aberto
    Set rsLotesALiberar = objclsTConciliacaoLote.SelecionarLotesALiberar()
    
    'verifica lote a lote se concilia com a R2, se SIM faz a Confirmação e Liberação
    If Not rsLotesALiberar.EOF Then
        rsLotesALiberar.MoveFirst
        Do Until rsLotesALiberar.EOF
            
            If rsLotesALiberar!TP_DEBT_CRED_LOTE = enumTipoDebitoCredito.Credito Then
            
                'zera as variáveis
                strLoteISPBIFDebitada = vbNullString
                strLoteCNPJCPFDebitado = vbNullString
                strLoteNomeRemetDebitado = vbNullString
                lngLoteAgenciaDebitada = 0
                lngLoteContaDebitada = 0
                lngLoteLocalLiquidacao = 0
                intTipoCorretora = 0
                dblLoteValor = 0
            
                Set rsOperacao = CreateObject("ADODB.Recordset")
                Set rsOperacao = objclsTConciliacaoLote.ObterOperacoesPorLote(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
                rsOperacao.MoveFirst
                Do Until rsOperacao.EOF
                    
                    strLoteISPBIFDebitada = rsOperacao!CO_ISPB_BANC_LIQU_CNPT
                    strLoteCNPJCPFDebitado = rsOperacao!CO_CNPJ_CNPT
                    strLoteNomeRemetDebitado = rsOperacao!NO_CNPT
                    lngLoteAgenciaDebitada = rsOperacao!CO_AGEN_COTR
                    lngLoteContaDebitada = rsOperacao!NU_CC_COTR
                    lngLoteLocalLiquidacao = rsOperacao!CO_LOCA_LIQU
                    intTipoCorretora = rsOperacao!TP_IF_CRED_DEBT
                    dblLoteValor = rsLotesALiberar!VL_TOTL_LOTE
                    
                    Exit Do
                Loop
                
                If intTipoCorretora = enumTipoInstituicaoCreditadaDebitada.CorretoraExterna Then
                
                    blnConciliacaoOK = False
                
                    Select Case udtSTR0008R2.CO_MESG
                        
                        Case "STR0008R2", "PAG0108R2"
                            If Val(strLoteISPBIFDebitada) = Val(udtSTR0008R2.CO_ISPB_REMT) _
                            And Val(strLoteCNPJCPFDebitado) = Val(udtSTR0008R2.NU_CNPJ_CPF_REMT) _
                            And Trim(UCase(strLoteNomeRemetDebitado)) = Trim(UCase(udtSTR0008R2.NO_REMT)) _
                            And lngLoteAgenciaDebitada = Val(udtSTR0008R2.CO_AGEN_REMT) _
                            And lngLoteContaDebitada = Val(udtSTR0008R2.NU_CNTA_REMT) _
                            And lngLoteLocalLiquidacao = plngLocalLiquidacaoR2 _
                            And dblLoteValor = dblMsgR2Valor Then
                                blnConciliacaoOK = True
                            End If
                        Case "STR0006R2"
                            If Val(strLoteISPBIFDebitada) = Val(udtSTR0008R2.CO_ISPB_REMT) _
                            And Val(strLoteCNPJCPFDebitado) = Val(udtSTR0008R2.NU_CNPJ_CPF_REMT) _
                            And Trim(UCase(strLoteNomeRemetDebitado)) = Trim(UCase(udtSTR0008R2.NO_REMT)) _
                            And lngLoteAgenciaDebitada = Val(udtSTR0008R2.CO_AGEN_REMT) _
                            And lngLoteLocalLiquidacao = plngLocalLiquidacaoR2 _
                            And dblLoteValor = dblMsgR2Valor Then
                                blnConciliacaoOK = True
                            End If
                        Case "STR0007R2"
                            If Val(strLoteISPBIFDebitada) = Val(udtSTR0008R2.CO_ISPB_REMT) _
                            And lngLoteLocalLiquidacao = plngLocalLiquidacaoR2 _
                            And dblLoteValor = dblMsgR2Valor Then
                                blnConciliacaoOK = True
                            End If
                    End Select
                
                    If blnConciliacaoOK = True Then
                    
                        blnConfirmacaoAutomatica = objRegra.VerificarRegraAutomatica2(rsOperacao!NU_SEQU_OPER_ATIV, enumFuncaoSistema.Confirmar)
                        blnLiberacaoAutomatica = objRegra.VerificarRegraAutomatica2(rsOperacao!NU_SEQU_OPER_ATIV, enumFuncaoSistema.Liberar)
                        blnIntegracaoCCAutomatica = objRegra.VerificarRegraAutomatica2(rsOperacao!NU_SEQU_OPER_ATIV, enumFuncaoSistema.IntegracaoCC)
                        
                        'dependendo do status da operação, precisa Confirmar e Liberar, ou então apenas Liberar
                        Select Case rsOperacao!CO_ULTI_SITU_PROC
                            
                            Case enumStatusOperacao.EmSer, enumStatusOperacao.ManualEmSer
                                If blnConfirmacaoAutomatica = True Then
                                    objclsTConciliacaoLote.ProcessarLote rsLotesALiberar, enumAcaoConciliacao.BOConcordar
                                    If blnLiberacaoAutomatica = True Then
                                        objclsTConciliacaoLote.ProcessarLote rsLotesALiberar, enumAcaoConciliacao.AdmAreaLiberar
                                        VerificarConciliacaoR2PagtoCorretora = True
                                        If blnIntegracaoCCAutomatica = True Then
                                            Call objclsTConciliacaoLote.IntegrarContaCorrente(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
                                        End If
                                    End If
                                End If
                            
                            Case enumStatusOperacao.ConcordanciaBackoffice, enumStatusOperacao.ConcordanciaBackofficeAutomatico
                                If blnLiberacaoAutomatica = True Then
                                    objclsTConciliacaoLote.ProcessarLote rsLotesALiberar, enumAcaoConciliacao.AdmAreaLiberar
                                    VerificarConciliacaoR2PagtoCorretora = True
                                    If blnIntegracaoCCAutomatica = True Then
                                        Call objclsTConciliacaoLote.IntegrarContaCorrente(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
                                    End If
                                End If
                        
                        End Select
                        
                    End If
                    
                End If
        
            End If
            rsLotesALiberar.MoveNext
        Loop
    End If
        
    Set objMensagemEnviadaRecebida = Nothing
    Set objclsTConciliacaoLote = Nothing
    Set objRegra = Nothing
    Set rsLotesALiberar = Nothing
    Set rsOperacao = Nothing

Exit Function
ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    Set objclsTConciliacaoLote = Nothing
    Set objRegra = Nothing
    Set rsLotesALiberar = Nothing
    Set rsOperacao = Nothing
    
    Err.Clear
    VerificarConciliacaoR2PagtoCorretora = False
    
End Function
