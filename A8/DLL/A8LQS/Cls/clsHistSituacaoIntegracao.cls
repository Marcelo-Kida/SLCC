VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistSituacaoIntegracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'Objeto responsável pelo controle dos históricos da situação de integração com Conta Corrente e Contabil

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

'Inlcuir um registro na tabela A8.TB_HIST_SITU_LANC_CC_CNTB

Public Function Incluir(ByVal pvntSequenciaOperacao As Variant, _
                        ByVal pintTipoLancamentoIntegracao As Integer, _
                        ByVal pintStatusIntegracao As Integer, _
               Optional ByVal plngCodigoErro As Long = 0, _
               Optional ByVal pstrMensagemErro As String = vbNullString, _
               Optional ByVal pstrJustificativaCancelamento As String = vbNullString, _
               Optional ByVal plngNumSeqLancamento As Long = 0) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Insert Into A8.TB_HIST_SITU_LANC_CC_CNTB (                                           " & vbNewLine & _
             "             DH_SITU_LANC_CC,                                                         " & vbNewLine & _
             "             NU_SEQU_OPER_ATIV,                                                       " & vbNewLine & _
             "             TP_LANC_ITGR,                                                            " & vbNewLine & _
             "             CO_SITU_PROC,                                                            " & vbNewLine & _
             "             CO_ERRO,                                                                 " & vbNewLine & _
             "             TX_MESG_ERRO,                                                            " & vbNewLine & _
             "             TX_JUST_CANC,                                                            " & vbNewLine & _
             "             CO_USUA_ATLZ,                                                            " & vbNewLine & _
             "             CO_ETCA_TRAB_ATLZ,                                                       " & vbNewLine & _
             "             NR_SEQU_LANC     )                                                       " & vbNewLine & _
             " Values     (" & fgDataHoraServidor_To_Oracle & ",                                    " & vbNewLine & _
             "             " & pvntSequenciaOperacao & ",                                           " & vbNewLine & _
             "             " & pintTipoLancamentoIntegracao & ",                                    " & vbNewLine & _
             "             " & pintStatusIntegracao & ",                                            " & vbNewLine & _
             "             " & IIf(plngCodigoErro = 0, " Null ", plngCodigoErro) & ",                 " & vbNewLine & _
             "             '" & pstrMensagemErro & "',  " & vbNewLine & _
             "             '" & pstrJustificativaCancelamento & "', " & vbNewLine & _
             "             '" & fgUsuarioRede & "',                                                 " & vbNewLine & _
             "             '" & fgEstacaoTrabalhoUsuario & "',                                       " & vbNewLine & _
             "             " & plngNumSeqLancamento & " )"

    Call fgExecuteSQL(strSQL)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter os nomes das colunas da tabela A8.TB_HIST_SITU_LANC_CC_CNTB

Public Function ObterPropriedades() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT DH_SITU_LANC_CC,           " & vbNewLine & _
             "        NU_SEQU_OPER_ATIV,         " & vbNewLine & _
             "        TP_LANC_ITGR,              " & vbNewLine & _
             "        CO_SITU_PROC,              " & vbNewLine & _
             "        CO_ERRO,                   " & vbNewLine & _
             "        TX_MESG_ERRO,              " & vbNewLine & _
             "        TX_JUST_CANC,              " & vbNewLine & _
             "        CO_USUA_ATLZ,              " & vbNewLine & _
             "        CO_ETCA_TRAB_ATLZ,         " & vbNewLine & _
             "        NR_SEQU_LANC               " & vbNewLine & _
             " FROM   A8.TB_HIST_SITU_LANC_CC_CNTB"

    ObterPropriedades = fgPropriedades("Grupo_HistSituacaoIntegracao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar registros da tabela A8.TB_HIST_SITU_LANC_CC_CNTB

Public Function LerTodos(Optional ByVal pvntNumeroSequencia As Variant = vbNullString, _
                         Optional ByVal pvntLancamentoITGR As Variant = vbNullString, _
                         Optional ByVal plngStatusIntegracao As Long = 0, _
                         Optional ByVal plngNumSeqLancamento As Long = 0) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT A.DH_SITU_LANC_CC,                    " & vbNewLine & _
             "        A.NU_SEQU_OPER_ATIV,                  " & vbNewLine & _
             "        A.TP_LANC_ITGR,                       " & vbNewLine & _
             "        A.CO_SITU_PROC,                       " & vbNewLine & _
             "        A.CO_ERRO,                            " & vbNewLine & _
             "        A.TX_MESG_ERRO,                       " & vbNewLine & _
             "        A.TX_JUST_CANC,                       " & vbNewLine & _
             "        A.CO_USUA_ATLZ,                       " & vbNewLine & _
             "        B.DE_SITU_PROC,                       " & vbNewLine & _
             "        A.NR_SEQU_LANC                        " & vbNewLine & _
             " FROM   A8.TB_HIST_SITU_LANC_CC_CNTB A,       " & vbNewLine & _
             "        A8.TB_SITU_PROC              B        " & vbNewLine & _
             " WHERE  A.CO_SITU_PROC = B.CO_SITU_PROC       " & vbNewLine
    
    If pvntNumeroSequencia <> vbNullString Then
        strSQL = strSQL & vbNewLine & _
            " And       A.NU_SEQU_OPER_ATIV = " & Abs(pvntNumeroSequencia) & "    "
    End If

    If pvntLancamentoITGR <> 0 Then
        strSQL = strSQL & vbNewLine & _
            " And       A.TP_LANC_ITGR      = " & Val(pvntLancamentoITGR) & "   "
    End If

    If plngStatusIntegracao <> 0 Then
        strSQL = strSQL & vbNewLine & _
            " And       A.CO_SITU_PROC      = " & plngStatusIntegracao & "   "
    End If
    
    strSQL = strSQL & vbNewLine & _
            " And       A.NR_SEQU_LANC      = " & plngNumSeqLancamento & "   "

    strSQL = strSQL & _
            " ORDER BY A.DH_SITU_LANC_CC DESC "
    
    If pvntNumeroSequencia < 0 Then
        'Consulta na base historica
        strSQL = Replace(strSQL, "A8.TB_HIST_SITU_LANC_CC_CNTB", "A8HIST.TB_HIST_SITU_LANC_CC_CNTB")
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

Public Function Ler(Optional ByVal pstrClausulaWhere As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT A.DH_SITU_LANC_CC,                    " & vbNewLine & _
             "        A.NU_SEQU_OPER_ATIV,                  " & vbNewLine & _
             "        A.TP_LANC_ITGR,                       " & vbNewLine & _
             "        A.CO_SITU_PROC,                       " & vbNewLine & _
             "        A.CO_ERRO,                            " & vbNewLine & _
             "        A.TX_MESG_ERRO,                       " & vbNewLine & _
             "        A.TX_JUST_CANC,                       " & vbNewLine & _
             "        A.CO_USUA_ATLZ,                       " & vbNewLine & _
             "        B.DE_SITU_PROC                        " & vbNewLine & _
             " FROM   A8.TB_HIST_SITU_LANC_CC_CNTB A,       " & vbNewLine & _
             "        A8.TB_SITU_PROC              B        " & vbNewLine & _
             " WHERE  A.CO_SITU_PROC = B.CO_SITU_PROC       " & vbNewLine
    
    If pstrClausulaWhere <> vbNullString Then
        strSQL = strSQL & vbNewLine & _
            " AND " & pstrClausulaWhere
    End If

    strSQL = strSQL & _
            " ORDER BY A.DH_SITU_LANC_CC DESC "
    
    Ler = fgQueryXMLLerTodos("HistSituacaoIntegracao", strSQL, App.EXEName & "." & TypeName(Me))

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

'Executar a operação de LErTodos

Public Function Executar(ByVal pxmlHistSituacaoIntegracao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String
Dim vntNumeroSequencia                      As Variant
Dim vntTipoLancamentoITGR                   As Variant
Dim lngSequenciaLancamento                  As Long

On Error GoTo ErrorHandler

    strOperacao = pxmlHistSituacaoIntegracao.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodos"
            vntNumeroSequencia = pxmlHistSituacaoIntegracao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
            vntTipoLancamentoITGR = pxmlHistSituacaoIntegracao.selectSingleNode("//TP_LANC_ITGR").Text
            
            lngSequenciaLancamento = 0
            If Not pxmlHistSituacaoIntegracao.selectSingleNode("//NR_SEQU_LANC") Is Nothing Then
                lngSequenciaLancamento = Val(pxmlHistSituacaoIntegracao.selectSingleNode("//NR_SEQU_LANC").Text)
            End If
            
            Executar = LerTodos(vntNumeroSequencia, vntTipoLancamentoITGR, , lngSequenciaLancamento)
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

