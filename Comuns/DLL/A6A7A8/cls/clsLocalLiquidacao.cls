VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocalLiquidacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1D4E0F0363"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

'Objeto responsável pela manutenção do cadastro de veiculo legal

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Implements COMSVCSLib.ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

'Recuperar os nomes das colunas da tabela A8.TB_LOCAL_LIQUIDACAO

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " Select    CO_LOCA_LIQU,           " & vbNewLine & _
             "           CO_EMPR_FUSI,           " & vbNewLine & _
             "           SG_LOCA_LIQU,           " & vbNewLine & _
             "           DE_LOCA_LIQU,           " & vbNewLine & _
             "           DT_INIC_VIGE,           " & vbNewLine & _
             "           DT_FIM_VIGE,            " & vbNewLine & _
             "           DH_ULTI_ATLZ,           " & vbNewLine & _
             "           'S' AS TP_VIGE,         " & vbNewLine & _
             "           'S' AS TP_SEGR          " & vbNewLine & _
             " From      A8.TB_LOCAL_LIQUIDACAO  "

    ObterPropriedades = fgPropriedades("Grupo_LocalLiquidacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar a operação de LerTodos

Public Function Executar(ByRef pxmlLocalLiquidacao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlLocalLiquidacao.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodos"
            Executar = LerTodos(pxmlLocalLiquidacao.selectSingleNode("//TP_SEGR").Text, pxmlLocalLiquidacao.selectSingleNode("//TP_VIGE").Text)
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

'Recuparar os registros da tabela A8.TB_LOCAL_LIQUIDACAO
'pstrSegregaDados       - Indicador de Segragação de dados (S-Sim;N-Não)
'pstrIndicadorVigencia  - Indicador de Vigencia (S-Sim;N-Não)
'prsLocalLiquidacao     - Retorno ADODB.Recordset


Public Function LerTodos(Optional ByVal pstrSegregaDados As String = "S", _
                         Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByRef prsLocalLiquidacao As ADODB.Recordset) As String

Dim strSQL                                  As String
Dim strCondicao                             As String
    
On Error GoTo ErrorHandler

    strSQL = " Select    CO_LOCA_LIQU,           " & vbNewLine & _
             "           CO_EMPR_FUSI,           " & vbNewLine & _
             "           SG_LOCA_LIQU,           " & vbNewLine & _
             "           DE_LOCA_LIQU,           " & vbNewLine & _
             "           DT_INIC_VIGE,           " & vbNewLine & _
             "           DT_FIM_VIGE,            " & vbNewLine & _
             "           DH_ULTI_ATLZ,           " & vbNewLine & _
             "           'S' AS TP_VIGE,         " & vbNewLine & _
             "           'S' AS TP_SEGR          " & vbNewLine & _
             " From      A8.TB_LOCAL_LIQUIDACAO  " & vbNewLine & _
             " Where     CO_EMPR_FUSI =  1       " & vbNewLine & _
             " And       CO_LOCA_LIQU <> 2       " & vbNewLine
    
    If pstrIndicadorVigencia = "S" Then
        strCondicao = " AND DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                      " AND (DT_FIM_VIGE IS  NULL       " & _
                      " OR  DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                       
        strSQL = strSQL & strCondicao & vbNewLine
    End If
    
    If pstrSegregaDados = "S" Then
        strCondicao = fgSegregaDados("A8.TB_LOCAL_LIQUIDACAO", _
                                            False, _
                                            "", _
                                            "", _
                                            Len(strCondicao) = 0, _
                                            False, _
                                            False, _
                                            True, _
                                            False, _
                                            False)
        
        strSQL = strSQL & strCondicao & vbNewLine
        
    End If
    
    strSQL = strSQL & " Order By  CO_LOCA_LIQU ASC"
    
    If prsLocalLiquidacao Is Nothing Then
       LerTodos = fgQueryXMLLerTodos("LocalLiquidacao", strSQL, App.EXEName & "." & TypeName(Me))
    Else
       Set prsLocalLiquidacao = fgQuerySQL(strSQL)
       LerTodos = ""
    End If

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

'Recuperar um registro da tabela A8.TB_LOCAL_LIQUIDACAO

Public Function Ler(ByVal plngCodigoLocalLiquidacao As Long)

Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler

    strSQL = " Select    CO_EMPR_FUSI,           " & vbNewLine & _
             "           CO_LOCA_LIQU,           " & vbNewLine & _
             "           SQ_ISPB,                " & vbNewLine & _
             "           SG_LOCA_LIQU,           " & vbNewLine & _
             "           DE_LOCA_LIQU,           " & vbNewLine & _
             "           DT_INIC_VIGE,           " & vbNewLine & _
             "           DT_FIM_VIGE,            " & vbNewLine & _
             "           ID_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "           DH_ULTI_ATLZ,           " & vbNewLine & _
             "           'S' AS TP_VIGE,         " & vbNewLine & _
             "           'S' AS TP_SEGR          " & vbNewLine & _
             " From      A8.TB_LOCAL_LIQUIDACAO  " & vbNewLine & _
             " Where     CO_EMPR_FUSI = 1        " & vbNewLine & _
             " And       CO_LOCA_LIQU = " & plngCodigoLocalLiquidacao
    
    Ler = fgQueryXMLLerTodos("LocalLiquidacao", strSQL, App.EXEName & "." & TypeName(Me))
    
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
