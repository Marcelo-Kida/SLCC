VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJustificativaArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável por fornecer informações das justificativas da area da liquidacao multilateral

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

'Alterar um registro da tabela TB_JUST_AREA_LIQU_MULT
Public Function Alterar(ByVal plngTipoBackOffice As Long, _
                        ByVal pintCodigoEmpresa As Integer, _
                        ByVal pstrDataLiquidacao As String, _
                        ByVal pstrJustificativa As String) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update     A8.TB_JUST_AREA_LIQU_MULT                                   " & vbNewLine & _
             " Set        DE_JUST                = '" & pstrJustificativa & "',       " & vbNewLine & _
             "            DH_ULTI_ATLZ           = SYSDATE,                           " & vbNewLine & _
             "            CO_USUA_ULTI_ATLZ      = '" & fgUsuarioRede & "',           " & vbNewLine & _
             "            CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "' " & vbNewLine & _
             " Where      TP_BKOF = " & plngTipoBackOffice & "                        " & vbNewLine & _
             " And        CO_EMPR = " & pintCodigoEmpresa & "                         " & vbNewLine & _
             " And        DT_LIQU = " & fgDtXML_To_Oracle(pstrDataLiquidacao)

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Alterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Método genérico acionado pelo objeto A8MIU.clsConsulta para a execução de métodos desta classe
Public Function Executar(ByRef pxmlExecucao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlExecucao.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodos"
            Executar = LerTodos(pxmlExecucao.selectSingleNode("//TP_BKOF").Text, _
                                pxmlExecucao.selectSingleNode("//CO_EMPR").Text, _
                                pxmlExecucao.selectSingleNode("//DT_LIQU").Text)
            
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
            
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Incluir um registro na tabela TB_JUST_AREA_LIQU_MULT
Public Function Incluir(ByVal plngTipoBackOffice As Long, _
                        ByVal pintCodigoEmpresa As Integer, _
                        ByVal pstrDataLiquidacao As String, _
                        ByVal pstrJustificativa As String) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Insert Into A8.TB_JUST_AREA_LIQU_MULT (" & vbNewLine & _
             "            TP_BKOF,               " & vbNewLine & _
             "            CO_EMPR,               " & vbNewLine & _
             "            DT_LIQU,               " & vbNewLine & _
             "            DE_JUST,               " & vbNewLine & _
             "            DH_ULTI_ATLZ,          " & vbNewLine & _
             "            CO_USUA_ULTI_ATLZ,     " & vbNewLine & _
             "            CO_ETCA_TRAB_ULTI_ATLZ)" & vbNewLine & _
             " Values    (" & plngTipoBackOffice & "," & vbNewLine & _
             "            " & pintCodigoEmpresa & ", " & vbNewLine & _
             "            " & fgDtXML_To_Oracle(pstrDataLiquidacao) & ", " & vbNewLine & _
             "           '" & IIf(pstrJustificativa = vbNullString, "Null", pstrJustificativa) & "', " & vbNewLine & _
             "            SYSDATE, " & vbNewLine & _
             "            '" & fgUsuarioRede & "', " & vbNewLine & _
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

'Recuperar registros da tabela TB_JUST_AREA_LIQU_MULT
Public Function LerTodos(Optional ByVal plngTipoBackOffice As Long = 0, _
                         Optional ByVal pintCodigoEmpresa As Integer = 0, _
                         Optional ByVal pstrDataLiquidacao As String = vbNullString) As String

Dim strSQL                                  As String
Dim blnTemWhere                             As Boolean

    On Error GoTo ErrorHandler

    strSQL = "Select    TP_BKOF,               " & vbNewLine & _
             "          CO_EMPR,               " & vbNewLine & _
             "          DT_LIQU,               " & vbNewLine & _
             "          DE_JUST,               " & vbNewLine & _
             "          DH_ULTI_ATLZ,          " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ,     " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ " & vbNewLine & _
             " From     A8.TB_JUST_AREA_LIQU_MULT " & vbNewLine

    blnTemWhere = False

    If plngTipoBackOffice > 0 Then
        strSQL = strSQL & " WHERE   TP_BKOF     =   " & plngTipoBackOffice
        blnTemWhere = True
    End If
    
    If pintCodigoEmpresa > 0 Then
        If blnTemWhere Then
            strSQL = strSQL & " AND     CO_EMPR     =   " & pintCodigoEmpresa
        Else
            strSQL = strSQL & " WHERE   CO_EMPR     =   " & pintCodigoEmpresa
        End If
        blnTemWhere = True
    End If
    
    If pstrDataLiquidacao <> vbNullString Then
        If blnTemWhere Then
            strSQL = strSQL & " AND     DT_LIQU     =   " & fgDtXML_To_Oracle(pstrDataLiquidacao)
        Else
            strSQL = strSQL & " WHERE   DT_LIQU     =   '" & fgDtXML_To_Oracle(pstrDataLiquidacao)
        End If
        blnTemWhere = True
    End If
    
    LerTodos = fgQueryXMLLerTodos("JustificativaArea", strSQL, App.EXEName & "." & TypeName(Me))

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
