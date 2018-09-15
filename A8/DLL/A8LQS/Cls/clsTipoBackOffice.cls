VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoBackOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 

'' Objeto responsável pelo fornecimento de informações relativas aos tipos de
'' BackOffices cadastrados

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

'Executar a operação de LerTodos

Public Function Executar(ByVal pxmlTipoBackOffice As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlTipoBackOffice.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodos"
            Executar = LerTodos(pxmlTipoBackOffice.documentElement.selectSingleNode("//TP_SEGR").Text, pxmlTipoBackOffice.documentElement.selectSingleNode("//TP_VIGE").Text)
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

'Obter os nomes das colunas da tabela de tipo de backoffice

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " Select   TP_BKOF,                " & vbNewLine & _
            "           DE_BKOF,                " & vbNewLine & _
            "           DT_INIC_VIGE,           " & vbNewLine & _
            "           DT_FIM_VIGE,            " & vbNewLine & _
            "           CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
            "           DH_ULTI_ATLZ,           " & vbNewLine & _
            "           '' AS TP_VIGE,          " & vbNewLine & _
            "           '' AS TP_SEGR           " & vbNewLine & _
            " From      A8.TB_TIPO_BKOF         "

    ObterPropriedades = fgPropriedades("Grupo_TipoBackOffice", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'recuperar registros da tabela de tipo de backoffice

Public Function LerTodos(Optional ByVal pstrSegregaDados As String = "S", _
                         Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByRef prsTipoBackOffice As ADODB.Recordset) As String

Dim strSQL                                  As String
Dim strCondicao                             As String
    
On Error GoTo ErrorHandler

    strSQL = " Select   TP_BKOF,                " & vbNewLine & _
            "           DE_BKOF,                " & vbNewLine & _
            "           DT_INIC_VIGE,           " & vbNewLine & _
            "           DT_FIM_VIGE,            " & vbNewLine & _
            "           CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
            "           DH_ULTI_ATLZ            " & vbNewLine & _
            " From      A8.TB_TIPO_BKOF         "
            
    If pstrIndicadorVigencia = "S" Then
        strCondicao = " WHERE DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
                       " AND  (DT_FIM_VIGE    IS  NULL       " & vbCrLf & _
                       " OR    DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
        strSQL = strSQL & strCondicao & vbNewLine
    End If
    
    If pstrSegregaDados = "S" Then
        strCondicao = fgSegregaDados("A8.TB_TIPO_BKOF", _
                                      False, _
                                      "", _
                                      "", _
                                      Len(strCondicao) = 0, _
                                      False, _
                                      True, _
                                      False, _
                                      False, _
                                      False)
        
        strSQL = strSQL & strCondicao & vbNewLine
    End If
    
    strSQL = strSQL & " Order By  TP_BKOF ASC "
            
    If prsTipoBackOffice Is Nothing Then
        LerTodos = fgQueryXMLLerTodos("TipoBackOffice", strSQL, App.EXEName & "." & TypeName(Me))
    Else
        Set prsTipoBackOffice = fgQuerySQL(strSQL)
        LerTodos = ""
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Obter tipo BackOffice Partipante Camara (Agente Compensação , Agente Custódia) CBLC
Public Function ObterTipoBackOfficeParticipanteCamara(ByVal vntCodigoPartCamara As Variant) As Long

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " Select   TP_BKOF,                    " & _
             "          CO_PARP_CAMR                " & _
             " From     A8.TB_TIPO_BKOF_PARP_CAMR   " & _
             " where    CO_PARP_CAMR = '" & vntCodigoPartCamara & "'"
    
    Set objRS = fgQuerySQL(strSQL)
    
    
    If objRS.EOF Then
        ObterTipoBackOfficeParticipanteCamara = 0
    Else
        ObterTipoBackOfficeParticipanteCamara = objRS!TP_BKOF
    End If
    
    objRS.Close
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOfficeParticipanteCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


