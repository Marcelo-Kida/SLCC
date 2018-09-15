VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNoTransactionOra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente     : CA
'Classe         : clsNoTransactionOra
'Data Criação   : 09-09-2001 18:00
'Objetivo       : Execução de Querys (Update, Delete, Select) através de
'                 sentenças SQL especificas para o Oracle
'Analista       : Adilson Gonçalves Damasceno
'
'Programador    :
'Data           :
'
'Teste          :
'Autor          :

'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

Public Function ExecuteSQL_ORA(ByVal pstrSQL As String) As Long

Dim objADODBConnection                        As ADODB.Connection

On Error GoTo ErrorHandler
    
    fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
            
    objADODBConnection.Open gstrConexaoOra
    
    objADODBConnection.CursorLocation = adUseClient
    
    objADODBConnection.Execute pstrSQL, ExecuteSQL_ORA, adCmdText + adExecuteNoRecords
                   
    If objADODBConnection.State <> 0 Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing

    Exit Function

ErrorHandler:
    
    If objADODBConnection.State <> 0 Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSQL_Ora", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function UpdateCLOB(ByVal pstrNomeTabela As String, _
                           ByRef pvntNomeCampoChave() As String, _
                           ByVal pstrNomeCampoCLOB As String, _
                           ByRef pvntConteudoChave() As Variant, _
                           ByVal pstrConteudoCLOB As String) As Long

Dim objADODBConnectionOra                     As ADODB.Connection
Dim objRS                                     As ADODB.Recordset
Dim strSQL                                    As String
Dim lngCont                                   As Long

On Error GoTo ErrorHandler
    
    fgObterVariaveisAmbiente
    
    Set objADODBConnectionOra = CreateObject("ADODB.Connection")
    Set objRS = CreateObject("ADODB.Recordset")
            
    objADODBConnectionOra.Open gstrConexaoOra

    strSQL = " SELECT " & pstrNomeCampoCLOB & _
             "   FROM " & pstrNomeTabela
    
    For lngCont = 0 To UBound(pvntNomeCampoChave)
        If lngCont = 0 Then
            If IsNumeric(pvntConteudoChave(lngCont)) Then
                strSQL = strSQL & " WHERE  " & pvntNomeCampoChave(lngCont) & " = " & pvntConteudoChave(lngCont)
            Else
                strSQL = strSQL & " WHERE  " & pvntNomeCampoChave(lngCont) & " = '" & pvntConteudoChave(lngCont) & "'"
            End If
        Else
            If IsNumeric(pvntConteudoChave(lngCont)) Then
                strSQL = strSQL & " AND  " & pvntNomeCampoChave(lngCont) & " = " & pvntConteudoChave(lngCont)
            Else
                If IsDate(pvntConteudoChave(lngCont)) Then
                    strSQL = strSQL & " AND  " & pvntNomeCampoChave(lngCont) & " = TO_DATE('" & Format(pvntConteudoChave(lngCont), "yyyymmdd") & "','YYYYMMDD')"
                Else
                    strSQL = strSQL & " AND  " & pvntNomeCampoChave(lngCont) & " = '" & pvntConteudoChave(lngCont) & "'"
                End If
            End If
        End If
    Next
    
    objRS.Open strSQL, objADODBConnectionOra, adOpenStatic, adLockOptimistic
    
    objRS.Fields(0).AppendChunk (pstrConteudoCLOB)
    
    objRS.Update
    objRS.Close

    Set objRS = Nothing
   
    If objADODBConnectionOra.State <> 0 Then
        objADODBConnectionOra.Close
    End If
    
    Set objADODBConnectionOra = Nothing
    
    Exit Function

ErrorHandler:
    
    If objADODBConnectionOra.State <> 0 Then
        objADODBConnectionOra.Close
    End If
    
    Set objRS = Nothing
    Set objADODBConnectionOra = Nothing
     
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




