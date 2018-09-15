VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Objeto responsável pela execução de comandos na base de dados, e devolver o resultado às classes e métodos
' que a instanciaram.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

Private objContext                           As COMSVCSLib.ObjectContext

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

'Executar um SELECT na base de dados, retornando um objeto ADODB.Recordset

Public Function QuerySQL(ByVal pstrSQL As String, _
                Optional ByVal pintQtdeRegistros As Integer) As ADODB.Recordset

Dim objADODBConnection                          As ADODB.Connection
Dim rsDisconnect                                As ADODB.Recordset
Dim dblLog                                      As Double

On Error GoTo ErrorHandler

    Set QuerySQL = Nothing
    
    fgObterVariaveisAmbiente
    
    fgLogText "QuerySQL start, sql=" & fgSimpleText(pstrSQL), dblLog
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")

    objADODBConnection.Mode = adModeRead
    objADODBConnection.CursorLocation = adUseClient
    objADODBConnection.ConnectionTimeout = ADO_CONNECTION_TIMEOUT
                            
    objADODBConnection.Open gstrConexao
    
    rsDisconnect.CursorType = adOpenStatic
    rsDisconnect.LockType = adLockBatchOptimistic

    If pintQtdeRegistros > 0 Then
        rsDisconnect.MaxRecords = pintQtdeRegistros
    End If
    
    Set rsDisconnect = objADODBConnection.Execute(pstrSQL, , adCmdText)
            
    Set rsDisconnect.ActiveConnection = Nothing
        
    Set QuerySQL = rsDisconnect
    
    If objADODBConnection.State <> 0 Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
   
    fgLogText "QuerySQL executada", dblLog
   
    Exit Function

ErrorHandler:
    
    If Not objADODBConnection Is Nothing Then
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
    End If
    
    If Not rsDisconnect Is Nothing Then
        If rsDisconnect.State <> ADODB.adStateClosed Then
            rsDisconnect.Close
        End If
    End If
    
    Set rsDisconnect = Nothing
    Set objADODBConnection = Nothing

    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QuerySQL", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)
    
End Function

'Executar um SELECT na base de dados, retornando um XML contendo um registro

Public Function QueryXMLLer(ByVal pstrNomeXML As String, _
                            ByVal pstrSQL As String, _
                            ByVal pstrNomeObjeto As String) As String

Dim xmlLer                                  As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlAttrib                               As MSXML2.IXMLDOMAttribute

Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field
Dim dblLog                                  As Double

On Error GoTo ErrorHandler

    fgObterVariaveisAmbiente
    fgLogText "QueryXMLLer start, sql=" & fgSimpleText(pstrSQL), dblLog

    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")

    objADODBConnection.ConnectionTimeout = ADO_CONNECTION_TIMEOUT
    objADODBConnection.CursorLocation = adUseClient
    
    Call objADODBConnection.Open(gstrConexao)

    rsDisconnect.CursorType = adOpenStatic
    rsDisconnect.LockType = adLockBatchOptimistic

    Set rsDisconnect = objADODBConnection.Execute(pstrSQL, , adCmdText)

    fgLogText "QueryXMLLer executada", dblLog
    
    If rsDisconnect.RecordCount > 0 Then
        
        Set xmlNode = xmlLer.createElement(pstrNomeXML)
        Call xmlLer.appendChild(xmlNode)
                
        'Criar Atributo padrão Operação
        Set xmlAttrib = xmlLer.createAttribute("Operacao")
        xmlAttrib.Text = "Consultar"
        Call xmlNode.attributes.setNamedItem(xmlAttrib)
        
        'Criar Atributo padrão Objeto
        Set xmlAttrib = xmlLer.createAttribute("Objeto")
        xmlAttrib.Text = pstrNomeObjeto
        Call xmlNode.attributes.setNamedItem(xmlAttrib)
        
        For Each objFiled In rsDisconnect.fields
            
            Set xmlNode = xmlLer.createElement(objFiled.Name)
            
            Select Case objFiled.Type
                Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                    If Not IsNull(objFiled.Value) Then
                        Select Case Mid$(objFiled.Name, 1, 2)
                            Case "DH", "HO"
                                xmlNode.Text = Format(objFiled.Value, "YYYYMMDDHHmmss")
                            Case "DT"
                                xmlNode.Text = Format(objFiled.Value, "YYYYMMDD")
                            Case Else
                                ' 101 - Tipo do Banco de Dados diferente da classe de dados
                                lngCodigoErroNegocio = 11
                                GoTo ErrorHandler
                        End Select
                    Else
                        xmlNode.Text = "00:00:00"
                    End If
                Case Is = ADODB.adNumeric, ADODB.adDouble, ADODB.adVarNumeric, ADODB.adCurrency, ADODB.adDecimal
                    If Not IsNull(objFiled.Value) Then
                        If UCase(Right$(objFiled.Name, 7)) = "_VLRXML" Then
                            xmlNode.Text = Replace$(objFiled.Value, ",", ".")
                        Else
                            'Coloca no padrão de separação de decimais em XML => vírgula (,)
                            xmlNode.Text = Replace$(objFiled.Value, ".", ",")
                        End If
                    Else
                        xmlNode.Text = 0
                    End If
                Case Else
                    If Not IsNull(objFiled.Value) Then
                        xmlNode.Text = objFiled.Value
                    Else
                        xmlNode.Text = ""
                    End If
            End Select
            
            Call xmlLer.documentElement.appendChild(xmlNode)
            
            Set xmlAttrib = xmlLer.createAttribute("Type")
            xmlAttrib.Text = objFiled.Type
            Call xmlNode.attributes.setNamedItem(xmlAttrib)
            
            Set xmlAttrib = xmlLer.createAttribute("Tamanho")
            
            If objFiled.Type = adChar Or objFiled.Type = adLongVarChar Or objFiled.Type = adVarChar Then
                xmlAttrib.Text = objFiled.DefinedSize
            Else
                xmlAttrib.Text = objFiled.Precision
            End If
            
            Call xmlNode.attributes.setNamedItem(xmlAttrib)
        
        Next
        QueryXMLLer = xmlLer.xml
    Else
        QueryXMLLer = ""
    End If

    If rsDisconnect.State <> ADODB.adStateClosed Then
        rsDisconnect.Close
    End If
    
    Set rsDisconnect = Nothing

    If objADODBConnection.State <> ADODB.adStateClosed Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
    Set xmlLer = Nothing
    
    fgLogText "QueryXMLLer transformada", dblLog
    Exit Function

ErrorHandler:

    If Not objADODBConnection Is Nothing Then
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
    End If
    
    If Not rsDisconnect Is Nothing Then
        If rsDisconnect.State <> ADODB.adStateClosed Then
            rsDisconnect.Close
        End If
    End If
    
    Set rsDisconnect = Nothing
    Set xmlLer = Nothing
    Set objADODBConnection = Nothing

    QueryXMLLer = vbNullString

    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QueryXMLLer", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)
    
End Function

'Obter os nomes das colunas de uma tabela, retornando um xml com as colunas

Public Function Propriedades(ByVal pstrNomeXML As String, _
                             ByVal pstrSQL As String, _
                             ByVal pstrNomeObjeto As String) As String

Dim xmlPropriedades                         As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlAttrib                               As MSXML2.IXMLDOMAttribute

Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field

On Error GoTo ErrorHandler
    
    Call fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")
        
    Call objADODBConnection.Open(gstrConexao)
    
    objADODBConnection.CursorLocation = ADODB.adUseClient
    rsDisconnect.CursorType = ADODB.adOpenStatic
    rsDisconnect.LockType = ADODB.adLockBatchOptimistic
    
    Set rsDisconnect = objADODBConnection.Execute(pstrSQL & " WHERE 1 = 2 ", , ADODB.adCmdText)
    
    Set xmlNode = xmlPropriedades.createElement(pstrNomeXML)
    Call xmlPropriedades.appendChild(xmlNode)
    Set xmlAttrib = xmlPropriedades.createAttribute("Operacao")
    xmlAttrib.Text = "Consultar"
    
    Call xmlNode.attributes.setNamedItem(xmlAttrib)
    Set xmlAttrib = xmlPropriedades.createAttribute("Objeto")
    xmlAttrib.Text = pstrNomeObjeto
    Call xmlNode.attributes.setNamedItem(xmlAttrib)
    
    For Each objFiled In rsDisconnect.fields
        Set xmlNode = xmlPropriedades.createElement(objFiled.Name)
        Call xmlPropriedades.documentElement.appendChild(xmlNode)
        Set xmlAttrib = xmlPropriedades.createAttribute("Type")
        xmlAttrib.Text = objFiled.Type
        Call xmlNode.attributes.setNamedItem(xmlAttrib)
        
        Set xmlAttrib = xmlPropriedades.createAttribute("Tamanho")
        
        If objFiled.Type = adChar Or objFiled.Type = adLongVarChar Or objFiled.Type = adVarChar Then
            xmlAttrib.Text = objFiled.DefinedSize
        Else
            xmlAttrib.Text = objFiled.Precision
        End If
        
        Call xmlNode.attributes.setNamedItem(xmlAttrib)
        
    Next
        
    Propriedades = xmlPropriedades.xml
    
    Set xmlPropriedades = Nothing
    
    If rsDisconnect.State <> ADODB.adStateClosed Then
        rsDisconnect.Close
    End If
    
    Set rsDisconnect = Nothing
    
    If objADODBConnection.State <> ADODB.adStateClosed Then
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
        If rsDisconnect.State <> ADODB.adStateClosed Then
            rsDisconnect.Close
        End If
    End If
    
    Set rsDisconnect = Nothing
    Set objADODBConnection = Nothing
    Set xmlPropriedades = Nothing
    
    Propriedades = vbNullString
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Propriedades", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)
    
End Function

'Executar um SELECT na base de dados, retornando um XML contendo grupos de resitros


Public Function QueryXMLLerTodos(ByVal pstrNomeXML As String, _
                                 ByVal pstrSQL As String, _
                                 ByVal pstrNomeObjeto As String, _
                        Optional ByVal pblnReturnDataType As Boolean = True)

Dim xmlLer                                  As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeAux                              As MSXML2.IXMLDOMNode
Dim xmlAttrib                               As MSXML2.IXMLDOMAttribute
    
Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field
Dim dblLog                                  As Double
Dim strValor                                As String


On Error GoTo ErrorHandler
    
    fgObterVariaveisAmbiente
    fgLogText "QueryXMLLerTodos start, sql=" & fgSimpleText(pstrSQL), dblLog
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    
    objADODBConnection.ConnectionTimeout = ADO_CONNECTION_TIMEOUT
    objADODBConnection.CursorLocation = ADODB.adUseClient
    
    Call objADODBConnection.Open(gstrConexao)
    
    rsDisconnect.CursorType = ADODB.adOpenStatic
    rsDisconnect.LockType = ADODB.adLockBatchOptimistic
    
    Set rsDisconnect = objADODBConnection.Execute(pstrSQL, , ADODB.adCmdText)
    fgLogText "QueryXMLLerTodos executada", dblLog
    
    If rsDisconnect.RecordCount > 0 Then
        
        Set xmlNode = xmlLer.createElement("Repeat_" & pstrNomeXML)
        Call xmlLer.appendChild(xmlNode)
        
        Set xmlNodeAux = xmlLer.createElement("Grupo_" & pstrNomeXML)
        Call xmlNode.appendChild(xmlNodeAux)
        
        Set xmlNode = xmlNodeAux
        Set xmlAttrib = xmlLer.createAttribute("Operacao")
        xmlAttrib.Text = "Consultar"
        
        Call xmlNodeAux.attributes.setNamedItem(xmlAttrib)
        
        Set xmlNode = xmlNodeAux
        Set xmlAttrib = xmlLer.createAttribute("Objeto")
        xmlAttrib.Text = pstrNomeObjeto
        
        Call xmlNodeAux.attributes.setNamedItem(xmlAttrib)
        
        Do While Not rsDisconnect.EOF
                   
            For Each objFiled In rsDisconnect.fields
            
                Set xmlNodeAux = xmlLer.createElement(objFiled.Name)
                
                Select Case objFiled.Type
                    
                    Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                        
                        If Not IsNull(objFiled.Value) Then
                            Select Case Mid(objFiled.Name, 1, 2)
                                Case "DH", "HO"
                                    xmlNodeAux.Text = Format(objFiled.Value, "YYYYMMDDHHmmss")
                                Case "DT"
                                    xmlNodeAux.Text = Format(objFiled.Value, "YYYYMMDD")
                                Case Else
                                    ' 101 - Tipo do Banco de Dados diferente da classe de dados
                                    lngCodigoErroNegocio = 11
                                    GoTo ErrorHandler
                            End Select
                        Else
                            xmlNodeAux.Text = "00:00:00"
                        End If
                    Case Is = ADODB.adNumeric, ADODB.adDouble, ADODB.adVarNumeric, ADODB.adCurrency, ADODB.adDecimal
                        
                        If Not IsNull(objFiled.Value) Then
                            
                            If UCase(Right$(objFiled.Name, 7)) = "_VLRXML" Then
                                xmlNodeAux.Text = Replace$(objFiled.Value, ",", ".")
                            Else
                                'Coloca no padrão de separação de decimais em XML => vírgula (,)
                                xmlNodeAux.Text = Replace$(objFiled.Value, ".", ",")
                            End If
                            
                        Else
                            xmlNodeAux.Text = 0
                        End If
                    Case Else
                        
                        If Not IsNull(objFiled.Value) Then
                            strValor = Replace(objFiled.Value, vbNullChar, vbNullString)
                            strValor = Replace(strValor, Chr(2), vbNullString)
                            strValor = Replace(strValor, Chr(26), vbNullString)
                            
                            xmlNodeAux.Text = strValor
                        Else
                            xmlNodeAux.Text = ""
                        End If
                
                End Select
                
                If pblnReturnDataType Then
                    Set xmlAttrib = xmlLer.createAttribute("Type")
                    xmlAttrib.Text = objFiled.Type
                    Call xmlNodeAux.attributes.setNamedItem(xmlAttrib)
                End If
                
                Call xmlNode.appendChild(xmlNodeAux)
            Next
            
            rsDisconnect.MoveNext
            
            If Not rsDisconnect.EOF Then
                Set xmlNode = xmlLer.createElement("Grupo_" & pstrNomeXML)
                Call xmlLer.childNodes.Item(0).appendChild(xmlNode)
                Set xmlAttrib = xmlLer.createAttribute("Operacao")
                xmlAttrib.Text = "Consultar"
                Call xmlNode.attributes.setNamedItem(xmlAttrib)
                
                Set xmlAttrib = xmlLer.createAttribute("Objeto")
                xmlAttrib.Text = pstrNomeObjeto
                Call xmlNode.attributes.setNamedItem(xmlAttrib)
            End If
        Loop
        QueryXMLLerTodos = xmlLer.xml
    Else
        QueryXMLLerTodos = ""
    End If
       
    If rsDisconnect.State <> adStateClosed Then
        rsDisconnect.Close
    End If
    Set rsDisconnect = Nothing
    
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
    Set xmlLer = Nothing
    
    fgLogText "QueryXMLLerTodos transformada", dblLog
    Exit Function

ErrorHandler:
    

    If Not objADODBConnection Is Nothing Then
        If objADODBConnection.State <> adStateClosed Then
            objADODBConnection.Close
        End If
    End If
    
    If Not rsDisconnect Is Nothing Then
        If rsDisconnect.State <> ADODB.adStateClosed Then
            rsDisconnect.Close
        End If
    End If
    
    Set rsDisconnect = Nothing
    Set xmlLer = Nothing
    Set objADODBConnection = Nothing
    
    QueryXMLLerTodos = vbNullString
    
    Set rsDisconnect = Nothing
    Set xmlLer = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QueryXMLLerTodos", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)

End Function

'Recuperar e compor os registros da tabela <OWBER>.TB_TEXT_XML

Public Function SelectVarchar4000(ByVal pstrNomeTabela As String, _
                                  ByVal pstrNomeCampoSequencia As String, _
                                  ByVal plngSequencial As Long, _
                                  ByVal pstrNomeCampoVarchar As String, _
                                  ByVal pstrNomeCampoOrdem As String, _
                         Optional ByVal pblnConverterBase64 As Boolean = True) As String

'Retorna decode base64

Dim strSQL                                        As String
Dim objRs                                         As ADODB.Recordset
Dim strConteudoVarchar                            As String

On Error GoTo ErrorHandler
         
    strSQL = " SELECT " & pstrNomeCampoVarchar & _
             "   FROM " & pstrNomeTabela & _
             "  WHERE " & pstrNomeCampoSequencia & " =   " & plngSequencial & _
             "  ORDER BY " & pstrNomeCampoOrdem
    
    Set objRs = QuerySQL(strSQL)
    
    If Not objRs.EOF Then
        While Not objRs.EOF
            strConteudoVarchar = strConteudoVarchar & objRs.fields(0)
            objRs.MoveNext
        Wend
    End If
    
    objRs.Close
        
    If Trim(strConteudoVarchar) <> "" Then
        If pblnConverterBase64 Then
            SelectVarchar4000 = fgBase64Decode(strConteudoVarchar)
        Else
            SelectVarchar4000 = strConteudoVarchar
        End If
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "SelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Public Function QueryXMLLerTodosCache(ByVal pstrNomeXML As String, _
                                      ByVal pstrSQL As String, _
                                      ByVal pstrNomeObjeto As String, _
                                      ByVal pstrNomeCache As String)

Dim xmlLer                                  As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeAux                              As MSXML2.IXMLDOMNode
Dim xmlAttrib                               As MSXML2.IXMLDOMAttribute
    
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field
Dim objValidaRemessa                        As Object

On Error GoTo ErrorHandler
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
    
    Select Case pstrNomeCache
        Case "CONTROLE_PROCESSAMENTO"
            Set rsDisconnect = objValidaRemessa.CarregarControleProcessamento(pstrSQL)
            If Not rsDisconnect Is Nothing Then
                rsDisconnect.Filter = pstrSQL
            Else
                Set objValidaRemessa = Nothing
                QueryXMLLerTodosCache = vbNullString
                Exit Function
            End If
    End Select
    Set objValidaRemessa = Nothing
    
    If rsDisconnect.RecordCount > 0 Then
        Set xmlNode = xmlLer.createElement("Repeat_" & pstrNomeXML)
        Call xmlLer.appendChild(xmlNode)
        
        Set xmlNodeAux = xmlLer.createElement("Grupo_" & pstrNomeXML)
        Call xmlNode.appendChild(xmlNodeAux)
        
        Set xmlNode = xmlNodeAux
        Set xmlAttrib = xmlLer.createAttribute("Operacao")
        xmlAttrib.Text = "Consultar"
        
        Call xmlNodeAux.attributes.setNamedItem(xmlAttrib)
        
        Set xmlNode = xmlNodeAux
        Set xmlAttrib = xmlLer.createAttribute("Objeto")
        xmlAttrib.Text = pstrNomeObjeto
        
        Call xmlNodeAux.attributes.setNamedItem(xmlAttrib)
        
        Do While Not rsDisconnect.EOF
            For Each objFiled In rsDisconnect.fields
                Set xmlNodeAux = xmlLer.createElement(objFiled.Name)
                Select Case objFiled.Type
                    Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                        If Not IsNull(objFiled.Value) Then
                            Select Case Mid(objFiled.Name, 1, 2)
                                Case "DH", "HO"
                                    xmlNodeAux.Text = Format(objFiled.Value, "YYYYMMDDHHmmss")
                                Case "DT"
                                    xmlNodeAux.Text = Format(objFiled.Value, "YYYYMMDD")
                                Case Else
                                    ' 101 - Tipo do Banco de Dados diferente da classe de dados
                                    lngCodigoErroNegocio = 11
                                    GoTo ErrorHandler
                            End Select
                        Else
                            xmlNodeAux.Text = "00:00:00"
                        End If
                    Case Is = ADODB.adNumeric, ADODB.adDouble, ADODB.adVarNumeric, ADODB.adCurrency, ADODB.adDecimal
                        
                        If Not IsNull(objFiled.Value) Then
                            
                            If UCase(Right$(objFiled.Name, 7)) = "_VLRXML" Then
                                xmlNodeAux.Text = Replace$(objFiled.Value, ",", ".")
                            Else
                                'Coloca no padrão de separação de decimais em XML => vírgula (,)
                                xmlNodeAux.Text = Replace$(objFiled.Value, ".", ",")
                            End If
                            
                        Else
                            xmlNodeAux.Text = 0
                        End If
                    Case Else
                        
                        If Not IsNull(objFiled.Value) Then
                            xmlNodeAux.Text = Trim(Replace(objFiled.Value, vbNullChar, vbNullString))
                            xmlNodeAux.Text = Replace(objFiled.Value, Chr(2), vbNullString)
                        Else
                            xmlNodeAux.Text = ""
                        End If
                
                End Select
                
                Call xmlNode.appendChild(xmlNodeAux)
            Next
            
            rsDisconnect.MoveNext
            
            If Not rsDisconnect.EOF Then
                Set xmlNode = xmlLer.createElement("Grupo_" & pstrNomeXML)
                Call xmlLer.childNodes.Item(0).appendChild(xmlNode)
                Set xmlAttrib = xmlLer.createAttribute("Operacao")
                xmlAttrib.Text = "Consultar"
                Call xmlNode.attributes.setNamedItem(xmlAttrib)
                
                Set xmlAttrib = xmlLer.createAttribute("Objeto")
                xmlAttrib.Text = pstrNomeObjeto
                Call xmlNode.attributes.setNamedItem(xmlAttrib)
            End If
        Loop
        QueryXMLLerTodosCache = xmlLer.xml
    Else
        QueryXMLLerTodosCache = vbNullString
    End If
    
    Set rsDisconnect = Nothing
    Set xmlLer = Nothing
    
    Exit Function

ErrorHandler:
    
    If Not rsDisconnect Is Nothing Then
        If rsDisconnect.State <> ADODB.adStateClosed Then
            rsDisconnect.Close
        End If
    End If
    
    Set rsDisconnect = Nothing
    Set xmlLer = Nothing
    Set objValidaRemessa = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QueryXMLLerTodosCache", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)

End Function




