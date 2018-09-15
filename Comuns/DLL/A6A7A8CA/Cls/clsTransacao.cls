VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Objeto responsável pela execução e controle de processos transacionais do sistema.

Option Explicit

'Variável utilizada para tratamento de erros
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

'Executar as sentenças de Insert e Update na base de dados

Public Function ExecuteSQL(ByVal pstrSQL As String) As Long

Dim objADODBConnection                      As ADODB.Connection
Dim lngRecordsAffected                      As Long

On Error GoTo ErrorHandler
    
    fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    
    objADODBConnection.CursorLocation = adUseClient
    objADODBConnection.CommandTimeout = ADO_CONNECTION_TIMEOUT
    objADODBConnection.Open gstrConexao
    
    objADODBConnection.Execute pstrSQL, lngRecordsAffected, adCmdText + adExecuteNoRecords
                           
    ExecuteSQL = lngRecordsAffected
    
    If lngRecordsAffected > 0 Then
        flAtualizaUsuLog pstrSQL
    End If
                       
    If objADODBConnection.State <> 0 Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
    
    Exit Function

ErrorHandler:

    If fgTratarErroExclusaoFisica Then
        'Informação não pode ser excluída pois já está sendo utilizada
        lngCodigoErroNegocio = 8
        Err.Clear
        
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
        
        Set objADODBConnection = Nothing
        If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
        Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro, pstrSQL)
    
    ElseIf fgTratarInclusaoDuplicadaOperacao Then
        '3037 - Identificador da Operação já existe.
        lngCodigoErroNegocio = 3037
        Err.Clear
        
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
        
        Set objADODBConnection = Nothing
        If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
        Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro, pstrSQL)
    
    ElseIf fgTratarErroInclusaoDuplicada Then
        '36 - Informação já existe, não pode ser incluída.
        lngCodigoErroNegocio = 36
        Err.Clear
        
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
        
        Set objADODBConnection = Nothing
        If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
        Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro, pstrSQL)
    Else

        If objADODBConnection.State <> 0 Then
            objADODBConnection.Close
        End If
        
        Set objADODBConnection = Nothing
        
        If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
        Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro, pstrSQL)
    End If

End Function

'Executar as sentenças de procedures na base de dados

Public Function ExecuteCMD(ByVal pstrNomeProc As String, _
                           ByVal pintPosicaoRetorno As Integer, _
                           ByRef pvntParametros() As Variant) As Variant

Dim objADODBConnection                      As ADODB.Connection
Dim objCommand                              As ADODB.Command

On Error GoTo ErrorHandler
        
    fgObterVariaveisAmbiente
    
    Set objCommand = CreateObject("ADODB.Command")
    Set objADODBConnection = CreateObject("ADODB.Connection")
        
    objADODBConnection.CursorLocation = adUseClient
    objADODBConnection.CommandTimeout = ADO_CONNECTION_TIMEOUT
    objADODBConnection.Open gstrConexao
                
    objCommand.ActiveConnection = objADODBConnection
    objCommand.CommandText = pstrNomeProc
    objCommand.CommandType = adCmdStoredProc
    If pintPosicaoRetorno > 0 Then objCommand.Parameters(pintPosicaoRetorno) = pintPosicaoRetorno
    objCommand.Execute ExecuteCMD, pvntParametros
    
    If pintPosicaoRetorno >= 0 Then ExecuteCMD = objCommand.Parameters(pintPosicaoRetorno)

    Set objCommand = Nothing

    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing

    Exit Function
    
ErrorHandler:
    
    If Not objADODBConnection Is Nothing Then
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
    End If
    
    Set objCommand.ActiveConnection = Nothing
    
    Set objCommand = Nothing
    Set objADODBConnection = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro, pstrNomeProc)
    
End Function

'Obter o valor de uma sequence

Public Function ExecuteSequence(ByVal pstrNomeSequence As String) As Variant
    
Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    fgObterVariaveisAmbiente
    
    strSQL = " SELECT " & pstrNomeSequence & ".NEXTVAL FROM DUAL"
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    
    objADODBConnection.CommandTimeout = ADO_CONNECTION_TIMEOUT
    objADODBConnection.CursorLocation = adUseClient
    Call objADODBConnection.Open(gstrConexao)
    
    
    rsDisconnect.CursorType = adOpenStatic
    rsDisconnect.LockType = adLockBatchOptimistic
    
    Set rsDisconnect = objADODBConnection.Execute(strSQL, , adCmdText)
    
    ExecuteSequence = rsDisconnect.fields(0)
       
    If rsDisconnect.State <> adStateClosed Then
        rsDisconnect.Close
    End If
    Set rsDisconnect = Nothing
    
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
    
    Exit Function
ErrorHandler:
    
    If Not objADODBConnection Is Nothing Then
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
    End If
    
    If Not rsDisconnect Is Nothing Then
        If rsDisconnect.State <> adStateClosed Then
            rsDisconnect.Close
        End If
    End If
    
    
    Set rsDisconnect = Nothing
    Set objADODBConnection = Nothing
    
    ExecuteSequence = 0
    
     
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteSequence", lngCodigoErroNegocio, intNumeroSequencialErro, pstrNomeSequence)
    
End Function


'Decompor e inserir registros na tabela <OWER>.TB_TEXT_XML

Public Function InsertVarchar4000(ByVal pstrNomeTabela As String, _
                                  ByVal pstrNomeCampoSequencia As String, _
                                  ByVal pstrNomeCampoVarchar As String, _
                                  ByVal pstrConteudoCampoVarchar As String, _
                                  ByVal pstrNomeCampoOrdem As String, _
                                  ByVal pstrNomeSequence As String, _
                         Optional ByVal pstrCodigoTextoXML As String = vbNullString, _
                         Optional ByVal pblnConverterBase64 As Boolean = True) As Long

'RETORNA O NUMERO SEQUENCIAL

Dim lngSequencial                                 As Long
Dim strBase64Encode                               As String
Dim lngOrdem                                      As Long
Dim lngCont                                       As Long
Dim strSQL                                        As String

On Error GoTo ErrorHandler
     
    If pstrCodigoTextoXML = vbNullString Then
        lngSequencial = ExecuteSequence(pstrNomeSequence)
    Else
        lngSequencial = Val(pstrCodigoTextoXML)
    End If
    
    If pblnConverterBase64 Then
        strBase64Encode = fgBase64Encode(pstrConteudoCampoVarchar)
    Else
        strBase64Encode = pstrConteudoCampoVarchar
    End If
    
    For lngCont = 1 To Len(strBase64Encode) Step 4000
    
        lngOrdem = lngOrdem + 1
        
        strSQL = " INSERT INTO " & pstrNomeTabela & " (" & _
                 pstrNomeCampoSequencia & ", " & _
                 pstrNomeCampoOrdem & ", " & _
                 pstrNomeCampoVarchar & _
                 " ) VALUES ( " & _
                 lngSequencial & "," & _
                 lngOrdem & "," & _
                 "'" & Mid(strBase64Encode, lngCont, 4000) & "')"
    
        ExecuteSQL strSQL
    
    Next
    
    InsertVarchar4000 = lngSequencial
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "InsertVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAtualizaUsuLog(ByVal pstrSQL As String) As Boolean
    
Dim strSQL                                  As String
Dim strNomeTabela                           As String
Dim strNomeTabelaLog                        As String
Dim lngPosicaoIniSQL                        As Long
Dim vntSQL()                                As String
Dim strWhere                                As String
Dim strUsuario                              As String
Dim strEstacaoTrabalho                      As String
Dim lngCont                                 As Long
Dim strOwner                                As String
Dim objCAControleAcesso                     As A6A7A8CA.clsControleAcesso

    
On Error GoTo ErrorHandler
    
    If InStr(1, UCase(pstrSQL), "TB_LOG", vbBinaryCompare) > 0 Then
        Exit Function
    End If
    
    'Verifica se uma sentença SQL de DELETE or UPDATE
    If InStr(1, pstrSQL, "DELETE", vbBinaryCompare) = 0 Or _
       InStr(1, pstrSQL, "UPDATE", vbBinaryCompare) = 0 Then
       Exit Function
    End If
    
    strUsuario = flObterUsuario
    
    Set objCAControleAcesso = CreateObject("A6A7A8CA.clsControleAcesso")
    strEstacaoTrabalho = objCAControleAcesso.ObterEstacaoTrabalhoUsuario(strUsuario)
    Set objCAControleAcesso = Nothing
    
    pstrSQL = Trim(Replace(pstrSQL, vbCrLf, vbNullString))
    
    If InStr(1, UCase(pstrSQL), "BEGIN") > 0 Then
        vntSQL() = Split(pstrSQL, ";", , vbBinaryCompare)
    Else
        ReDim vntSQL(0)
        vntSQL(0) = pstrSQL
    End If
        
    For lngCont = 0 To UBound(vntSQL())
        
        strSQL = Trim(vntSQL(lngCont))
                
        lngPosicaoIniSQL = InStr(1, UCase(strSQL), "DELETE", vbBinaryCompare)
        
        If lngPosicaoIniSQL > 0 Then
            strOwner = Trim(Mid(UCase(strSQL), InStr(1, UCase(strSQL), "FROM", vbBinaryCompare) + 5))
            strOwner = Trim(Mid(strOwner, 1, InStr(1, strOwner, ".", vbBinaryCompare) - 1))
        End If
        
        If lngPosicaoIniSQL = 0 Then
            lngPosicaoIniSQL = InStr(1, UCase(strSQL), "UPDATE", vbBinaryCompare)
            strOwner = Trim(Mid(UCase(strSQL), lngPosicaoIniSQL + 6, InStr(1, UCase(strSQL), "TB_", vbBinaryCompare) - (lngPosicaoIniSQL + 7)))
        End If
        
        If lngPosicaoIniSQL > 0 Then
            strNomeTabela = Mid(strSQL, InStr(1, UCase(strSQL), "TB_"))
            strNomeTabela = Replace(strNomeTabela, Chr(13), " ")
            strNomeTabela = Mid(strNomeTabela, 1, InStr(1, strNomeTabela, " "))
            
            strWhere = ""
            If InStr(1, UCase(strSQL), "WHERE") > 0 Then
                strWhere = strSQL
                strWhere = " " & Mid(strWhere, InStr(1, UCase(strWhere), "WHERE"))
            End If
            
            'Obtem nome tabela de log : TB_CONTROLE_TABELA_LOG
            strNomeTabelaLog = flObterNomeTabelaLog(strNomeTabela, strOwner)
                        
            If strNomeTabelaLog <> "" Then
                strSQL = " UPDATE " & strOwner & "." & strNomeTabelaLog
                strSQL = strSQL & " SET CO_USUA_OPER       = '" & strUsuario & "',"
                strSQL = strSQL & "      CO_ETCA_TRAB_OPER = '" & strEstacaoTrabalho & "'"
                
                If strWhere <> "" Then
                    strSQL = strSQL & strWhere
                    strSQL = strSQL & " AND CO_USUA_OPER IS NULL "
                Else
                    strSQL = strSQL & " WHERE CO_USUA_OPER IS NULL "
                End If
                
                flAtualizaUsuLog = ExecuteSQL(strSQL)
            End If
            
        End If
    Next
   
    Exit Function
ErrorHandler:
    
    Set objCAControleAcesso = Nothing
    
    If InStr(1, Err.Description, "ORA-00942") <> 0 Then
        'Tabela de Log não existe
        lngCodigoErroNegocio = 9
        Err.Clear
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAtualizaUsuLog", lngCodigoErroNegocio, intNumeroSequencialErro, "Nome Tabela:" & strNomeTabela)

End Function

Private Function flObterNomeTabelaLog(ByVal pstrNomeTabela As String, _
                                      ByVal pstrOwner As String) As String
    
Dim strSQL                                                 As String
Dim objConsulta                                           As A6A7A8CA.clsConsulta
Dim objRs                                                 As ADODB.Recordset
    
On Error GoTo ErrorHandler
    
    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
                
    strSQL = " SELECT NO_TABE_LOG       " & _
             "   FROM " & pstrOwner & ".TB_CTRL_TABE_LOG " & _
             "  WHERE NO_TABE  = '" & Trim(pstrNomeTabela) & "'"
        
    Set objRs = objConsulta.QuerySQL(strSQL)
    
    Set objConsulta = Nothing
    
    If objRs.EOF Then
        flObterNomeTabelaLog = ""
    Else
        flObterNomeTabelaLog = objRs.fields(0)
    End If
   
    Exit Function
ErrorHandler:
    
    If Not objRs Is Nothing Then
        If objRs.State <> adStateClosed Then
            objRs.Close
        End If
    End If
    
    Set objRs = Nothing

    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterNomeTabelaLog", lngCodigoErroNegocio, intNumeroSequencialErro, "Nome Tabela:" & pstrNomeTabela)
End Function

Private Function flObterUsuario() As String

Dim objSecurity                             As COMSVCSLib.SecurityCallContext
Dim objIdentity                             As COMSVCSLib.SecurityIdentity
Dim objActiveDS                             As ActiveDs.WinNTSystemInfo
Dim strAccount                              As String
Dim strUsuario                              As String

On Error Resume Next

    If Not objContext Is Nothing Then
        
                
        Set objSecurity = GetSecurityCallContext()
        Set objIdentity = objSecurity.Item("OriginalCaller")
        
        strAccount = objIdentity.Item("AccountName")
        strUsuario = Right(strAccount, Len(strAccount) - InStr(strAccount, "\"))
       
        Set objIdentity = Nothing
        Set objSecurity = Nothing
    Else
        
        Set objActiveDS = New ActiveDs.WinNTSystemInfo
        strUsuario = objActiveDS.UserName
        Set objActiveDS = Nothing
    
    End If
    
    strUsuario = Trim(strUsuario)
    
    If Len(strUsuario) > 8 Then
        strUsuario = Left(strUsuario, 8)
    End If
    
    flObterUsuario = strUsuario
    
    Exit Function
ErrorHandler:
        
    Set objSecurity = Nothing
    Set objIdentity = Nothing
    Set objActiveDS = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

