VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Recuperação dos dados de LOG (Trilha de auditoria dos cadastros)

Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Private ObjectContext                       As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

'------------------------------------------------------------------------------------------------------------
'Todas as tabelas que possuem log de trilha de auditoria estão armazenadas na tabela <OWNER>.TB_CTRL_TABE_LOG
'------------------------------------------------------------------------------------------------------------

Public Function LerTodosControleLog(ByVal pintOwner As enumSLCCOwner) As String
    
Dim strSQL                                  As String
Dim strOwner                                As String
    
On Error GoTo ErrorHandler
    
    Select Case pintOwner
        Case enumSLCCOwner.OwnerA6
            strOwner = "A6"
        Case enumSLCCOwner.OwnerA7
            strOwner = "A7"
        Case enumSLCCOwner.OwnerA8
            strOwner = "A8"
        Case enumSLCCOwner.OwnerA6Coli
            strOwner = "A6"
    
    End Select
    
    If pintOwner <> enumSLCCOwner.OwnerA6Coli Then
        strSQL = " SELECT * FROM " & strOwner & ".TB_CTRL_TABE_LOG " & _
                 " WHERE   NO_TABE_PAI IS NULL " & _
                 " ORDER BY NO_TABE "
    Else
        strSQL = " SELECT * FROM " & strOwner & ".TB_CTRL_TABE_LOG " & _
                 " WHERE   NO_TABE_PAI IS NOT NULL " & _
                 " ORDER BY NO_TABE "
    
    End If
    
    LerTodosControleLog = fgQueryXMLLerTodos("ControleTabelaLog", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosControleLog Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Owner:" & pintOwner & "." & strOwner)

End Function

'------------------------------------------------------------------------------------------------------------
'Obter os nomes das colunas de uma tabela
'------------------------------------------------------------------------------------------------------------

Public Function ObterColunasTabela(ByVal pstrNomeTabela As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    strSQL = "  SELECT COLUMN_NAME,COLUMN_ID " & _
             "  FROM   ALL_TAB_COLUMNS " & _
             "  WHERE TABLE_NAME = '" & pstrNomeTabela & "' " & _
             "  ORDER BY COLUMN_ID"
    
    ObterColunasTabela = fgQueryXMLLerTodos("ColunaTabela", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterColunasTabela", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


'------------------------------------------------------------------------------
'recuperar os dados de uma tabela de Log
'-------------------------------------------------------------------------------
Public Function LerTodosLogTabela(ByVal pstrNomeTabelaLog As String, _
                                  ByVal pintOwner As enumSLCCOwner, _
                                  ByVal pdtmDataInicio As Date, _
                                  ByVal pdtmDataFim As Date) As String

Dim strSQL                                  As String
Dim strOwner                                As String

On Error GoTo ErrorHandler
    
    Select Case pintOwner
        Case enumSLCCOwner.OwnerA6
            strOwner = "A6"
        Case enumSLCCOwner.OwnerA7
            strOwner = "A7"
        Case enumSLCCOwner.OwnerA8
            strOwner = "A8"
        Case enumSLCCOwner.OwnerA6Coli
            strOwner = "A6"
    End Select
        
    
    If pintOwner <> enumSLCCOwner.OwnerA6Coli Then
               
        strSQL = " SELECT * FROM " & strOwner & "." & pstrNomeTabelaLog
        
        If CDate(pdtmDataInicio) = CDate(pdtmDataFim) Then
            strSQL = strSQL & " WHERE TO_CHAR(DH_OPER,'DD/MM/YYYY') = '" & Format(pdtmDataInicio, "DD/MM/YYYY") & "' "
        Else
            strSQL = strSQL & " WHERE DH_OPER >= to_date(' " & Format(pdtmDataInicio, "dd/mm/yyyy") & "','dd/mm/yyyy') " & _
                             "    AND DH_OPER <= to_date(' " & Format(pdtmDataFim, "dd/mm/yyyy") & " 23:59:59','dd/mm/yyyy hh24:mi:ss')"
        End If
        
        strSQL = strSQL & " ORDER BY DH_OPER DESC "
    Else
    
        strSQL = " SELECT * FROM A6.TB_TRIL_AUDI "
        
        If CDate(pdtmDataInicio) = CDate(pdtmDataFim) Then
            strSQL = strSQL & " WHERE TO_CHAR(DH_ALTR,'DD/MM/YYYY') = '" & Format(pdtmDataInicio, "DD/MM/YYYY") & "' "
        Else
            strSQL = strSQL & " WHERE DH_ALTR >= to_date(' " & Format(pdtmDataInicio, "dd/mm/yyyy") & "','dd/mm/yyyy') " & _
                             "    AND DH_ALTR <= to_date(' " & Format(pdtmDataFim, "dd/mm/yyyy") & " 23:59:59','dd/mm/yyyy hh24:mi:ss')"
        End If
        
        strSQL = strSQL & " AND NM_TABE = '" & UCase(pstrNomeTabelaLog) & "' "
        
        strSQL = strSQL & " ORDER BY DH_ALTR DESC "
    
    
    End If
    
    LerTodosLogTabela = fgQueryXMLLerTodos("TabelaLog", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosLogTabela", lngCodigoErroNegocio, intNumeroSequencialErro, "Owner:" & pintOwner)
End Function



