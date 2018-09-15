VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSistemaMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo fornecimento de informações relativas às associações
'' entre sistemas e mensagens SPB
''
''

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

'' Consultar todas as associações de sistemas x mensagem cadastradas
Public Function LerTodos(Optional ByVal pstrSiglaSistema As String = "A8") As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    'RATS 312 - OCOR 0001 - Correção Impacta somente fluxo da STR0008R2
    'RATS 352 - OCOR 001
    
    strSQL = " SELECT DISTINCT CO_MESG,                   " & vbCrLf & _
             "                 NO_MESG                    " & vbCrLf & _
             " FROM   A8.TB_MENSAGEM      A,              " & vbCrLf & _
             "        A8.TB_SIST_MENSAGEM B               " & vbCrLf & _
             " WHERE SG_SIST = '" & pstrSiglaSistema & "' " & vbCrLf & _
             " AND   A.SQ_MESG = B.SQ_MESG                " & vbCrLf & _
             " AND   SUBSTR(A.CO_MESG,8,2) = '  '         " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'STR0008'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'STR0005'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'PAG0107'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'STR0006'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'STR0007'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'PAG0108'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'LDL0005'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'LTR0005'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'LDL0006'  " & vbCrLf & _
             " AND    SUBSTR(A.CO_MESG,1,7) <> 'LTR0006'  "

    strSQL = strSQL & " UNION ALL " & vbCrLf & _
             " SELECT 'STR0008' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'STR0008R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'STR0005' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'STR0005R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'PAG0107' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'PAG0107R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'STR0006' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'STR0006R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'STR0007' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'STR0007R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'PAG0108' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'PAG0108R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'LDL0005' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'LDL0005R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'LTR0005' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'LTR0005R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'LDL0006' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'LDL0006R2'" & vbCrLf & _
             " UNION ALL " & vbCrLf & _
             " SELECT 'LTR0006' AS CO_MESG , NO_MESG From A8.TB_MENSAGEM WHERE  CO_MESG = 'LTR0006R2'" & vbCrLf & _
             " ORDER BY CO_MESG"

    LerTodos = fgQueryXMLLerTodos("SistemaMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

