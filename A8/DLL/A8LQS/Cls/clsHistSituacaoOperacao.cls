VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistSituacaoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável por fornecer informações históricas das operações cadastradas

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

'Incluir um registro na tabela TB_HIST_SITU_OPER_ATIV

Public Function Incluir(ByVal pvntSequenciaOperacao As Variant, _
               Optional ByVal pintStatusOperacao As Integer = 0, _
               Optional ByVal plngJustificativa As Long = 0, _
               Optional ByVal pintTipoAcao As Integer = 0, _
               Optional ByVal pstrTextoConteudoAnteriorAcao As String = "") As Boolean

Dim strSQL                                  As String
Dim strDataGravacao                         As String

    On Error GoTo ErrorHandler
    
    strDataGravacao = flObterDataGravacao(pvntSequenciaOperacao, _
                                          pintStatusOperacao, _
                                          fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora)))

    strSQL = "Insert Into A8.TB_HIST_SITU_ACAO_OPER_ATIV (                                          " & vbNewLine & _
             "            NU_SEQU_OPER_ATIV,                                                        " & vbNewLine & _
             "            DH_SITU_ACAO_OPER_ATIV,                                                   " & vbNewLine & _
             "            CO_SITU_PROC,                                                             " & vbNewLine & _
             "            TP_ACAO_OPER_ATIV,                                                        " & vbNewLine & _
             "            TP_JUST_SITU_PROC,                                                        " & vbNewLine & _
             "            TX_CNTD_ANTE_ACAO,                                                        " & vbNewLine & _
             "            CO_USUA_ATLZ,                                                             " & vbNewLine & _
             "            CO_ETCA_USUA_ATLZ)                                                        " & vbNewLine & _
             " Values    (" & pvntSequenciaOperacao & ",                                            " & vbNewLine & _
             "            " & fgDtHrXML_To_Oracle(strDataGravacao) & ",                             " & vbNewLine & _
             "            " & IIf(pintStatusOperacao = 0, "NULL", pintStatusOperacao) & ",          " & vbNewLine & _
             "            " & IIf(pintTipoAcao = 0, "NULL", pintTipoAcao) & ",                      " & vbNewLine & _
             "            " & IIf(plngJustificativa = 0, "Null", plngJustificativa) & ",            " & vbNewLine & _
             "            " & IIf(pstrTextoConteudoAnteriorAcao = vbNullString, "NULL", _
             "'" & Left$(pstrTextoConteudoAnteriorAcao, 50) & "'") & ",                                      " & vbNewLine & _
             "            '" & fgUsuarioRede & "',                                                  " & vbNewLine & _
             "            '" & fgEstacaoTrabalhoUsuario & "')"

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Function Ler(Optional ByVal pstrClausulaWhere As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Select    A.NU_SEQU_OPER_ATIV,                            " & vbNewLine & _
             "          A.DH_SITU_ACAO_OPER_ATIV,                       " & vbNewLine & _
             "          A.CO_SITU_PROC,                                 " & vbNewLine & _
             "          B.DE_SITU_PROC,                                 " & vbNewLine & _
             "          A.TP_ACAO_OPER_ATIV,                            " & vbNewLine & _
             "          A.TP_JUST_SITU_PROC,                            " & vbNewLine & _
             "          C.NO_TIPO_JUST_SITU_PROC,                       " & vbNewLine & _
             "          A.TX_CNTD_ANTE_ACAO,                            " & vbNewLine & _
             "          A.CO_USUA_ATLZ,                                 " & vbNewLine & _
             "          A.CO_ETCA_USUA_ATLZ                             " & vbNewLine & _
             " From     A8.TB_HIST_SITU_ACAO_OPER_ATIV   A,             " & vbNewLine & _
             "          A8.TB_SITU_PROC                  B,             " & vbNewLine & _
             "          A8.TB_TIPO_JUST_SITU_PROC        C              " & vbNewLine & _
             " Where    A.CO_SITU_PROC      = B.CO_SITU_PROC            " & vbNewLine & _
             " And      A.TP_JUST_SITU_PROC = C.TP_JUST_SITU_PROC   (+) "
    
    If pstrClausulaWhere <> vbNullString Then
        strSQL = strSQL & vbNewLine & _
            " AND " & pstrClausulaWhere
    End If

    strSQL = strSQL & _
            " ORDER BY A.DH_SITU_ACAO_OPER_ATIV DESC "
    
    Ler = fgQueryXMLLerTodos("SituacaoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

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

'Recuperar registros da tabela TB_HIST_SITU_OPER_ATIV

Public Function LerTodos(Optional ByVal pvntSequenciaOperacao As Variant, _
                         Optional ByVal plngStatusOperacao As Long) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Select    A.NU_SEQU_OPER_ATIV,                            " & vbNewLine & _
             "          A.DH_SITU_ACAO_OPER_ATIV,                       " & vbNewLine & _
             "          A.CO_SITU_PROC,                                 " & vbNewLine & _
             "          B.DE_SITU_PROC,                                 " & vbNewLine & _
             "          A.TP_ACAO_OPER_ATIV,                            " & vbNewLine & _
             "          A.TP_JUST_SITU_PROC,                            " & vbNewLine & _
             "          C.NO_TIPO_JUST_SITU_PROC,                       " & vbNewLine & _
             "          A.TX_CNTD_ANTE_ACAO,                            " & vbNewLine & _
             "          A.CO_USUA_ATLZ,                                 " & vbNewLine & _
             "          A.CO_ETCA_USUA_ATLZ                             " & vbNewLine & _
             " From     A8.TB_HIST_SITU_ACAO_OPER_ATIV   A,             " & vbNewLine & _
             "          A8.TB_SITU_PROC                  B,             " & vbNewLine & _
             "          A8.TB_TIPO_JUST_SITU_PROC        C              " & vbNewLine & _
             " Where    A.CO_SITU_PROC      = B.CO_SITU_PROC            " & vbNewLine & _
             " And      A.TP_JUST_SITU_PROC = C.TP_JUST_SITU_PROC   (+) "
    
    If pvntSequenciaOperacao <> Empty Then
    
        If pvntSequenciaOperacao < 0 Then
            strSQL = Replace$(strSQL, "A8.TB_HIST_SITU_ACAO_OPER_ATIV", "A8HIST.TB_HIST_SITU_ACAO_OPER_ATIV")
            pvntSequenciaOperacao = pvntSequenciaOperacao * -1
        End If
    
        strSQL = strSQL & vbNewLine & _
            " And       A.NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao
            
    End If
    
    If plngStatusOperacao > 0 Then
        strSQL = strSQL & vbNewLine & _
            " And       A.CO_SITU_PROC      = " & plngStatusOperacao
    End If
    
    strSQL = strSQL & _
            " ORDER BY  A.DH_SITU_ACAO_OPER_ATIV DESC, ROWNUM DESC "
    
    LerTodos = fgQueryXMLLerTodos("SituacaoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

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

'Alterar a justificativa da situacao de uma operação

Public Function AlterarJustificativa(ByRef pvntSequenciaOperacao As Variant, _
                                     ByRef plngStatusSistema As Long, _
                                     ByRef plngJustificativa As Long) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Update    A8.TB_HIST_SITU_ACAO_OPER_ATIV " & vbCrLf & _
             "   Set    TP_JUST_SITU_PROC       =   " & plngJustificativa & vbCrLf & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao & vbCrLf & _
             "   And    CO_SITU_PROC            =   " & plngStatusSistema & vbCrLf & _
             "   And    DH_SITU_ACAO_OPER_ATIV  =   (Select Max(DH_SITU_ACAO_OPER_ATIV) " & vbCrLf & _
             "                                         From A8.TB_HIST_SITU_ACAO_OPER_ATIV " & vbCrLf & _
             "                                        Where NU_SEQU_OPER_ATIV   =   " & pvntSequenciaOperacao & vbCrLf & _
             "                                          And CO_SITU_PROC        =   " & plngStatusSistema & ")"

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter Data e Hora do ultimo registro de situação de uma operação

Private Function flObterDataGravacao(ByRef pvntSequenciaOperacao As Variant, _
                                     ByRef pintStatusOperacao As Integer, _
                                     ByVal pstrDataSituacao As String) As String

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim datDataGravacao                         As Date

    On Error GoTo ErrorHandler

    strSQL = " SELECT   MAX(DH_SITU_ACAO_OPER_ATIV) AS DH_SITU_ACAO_OPER_ATIV " & vbNewLine & _
             " FROM     A8.TB_HIST_SITU_ACAO_OPER_ATIV    " & vbNewLine & _
             " WHERE    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao & vbNewLine & _
             " AND      CO_SITU_PROC            =   " & pintStatusOperacao & vbNewLine & _
             " AND      DH_SITU_ACAO_OPER_ATIV  =   " & fgDtHrXML_To_Oracle(pstrDataSituacao)

    Set objRS = fgQuerySQL(strSQL)

    If IsNull(objRS!DH_SITU_ACAO_OPER_ATIV) Then
        datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
    Else
        If DateDiff("s", fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux), objRS!DH_SITU_ACAO_OPER_ATIV) = 0 Then
            datDataGravacao = DateAdd("s", 1, fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux))
        Else
            datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
        End If
    End If

    flObterDataGravacao = fgDtHr_To_Xml(datDataGravacao)

    objRS.Close
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDataGravacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
