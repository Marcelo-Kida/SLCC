VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBLConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente     : CA
'Classe         : clsDBLConsulta
'Data Criação   : 17-10-2003 12:14
'Objetivo       : Execução de Querys de Consultas via DBLink - Esta classe dever ser configurada no COM+ como Not Suported
'Analista       : Adilson Gonçalves Damasceno / Carlos Fortes
'
'Programador    : Carlos Fortes
'Data           : 17-10-2003
'
'Teste          :
'Autor          :
'

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

Public Function QuerySQL(ByVal pstrSQL As String, _
                Optional ByVal pintQtdeRegistros As Integer) As ADODB.Recordset

Dim objADODBConnection                        As ADODB.Connection
Dim rsDisconnect                              As ADODB.Recordset

On Error GoTo ErrorHandler

    Set QuerySQL = Nothing
    
    fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")

    objADODBConnection.Mode = adModeRead
                    
    objADODBConnection.Open gstrConexao
    
    objADODBConnection.CursorLocation = adUseClient
    
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

    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QuerySQL", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)
    
End Function

Public Function QueryXMLLer(ByVal pstrNomeXML As String, _
                            ByVal pstrSQL As String, _
                            ByVal pstrNomeObjeto As String) As String

Dim objDOMDocument                          As MSXML2.DOMDocument40
Dim objDOMNode                              As MSXML2.IXMLDOMNode
Dim objDomAttrib                            As MSXML2.IXMLDOMAttribute
Dim objProcessingInstruction                As IXMLDOMProcessingInstruction

Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field
Dim strProcessingInstruction                As String

On Error GoTo ErrorHandler

    fgObterVariaveisAmbiente

    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set objDOMDocument = CreateObject("MSXML2.DOMDocument.4.0")

    Call objADODBConnection.Open(gstrConexao)

    objADODBConnection.CursorLocation = adUseClient
    
    rsDisconnect.CursorType = adOpenStatic
    rsDisconnect.LockType = adLockBatchOptimistic

    Set rsDisconnect = objADODBConnection.Execute(pstrSQL, , adCmdText)

    If rsDisconnect.RecordCount > 0 Then
        
        Set objDOMNode = objDOMDocument.createElement(pstrNomeXML)
        Call objDOMDocument.appendChild(objDOMNode)
        
        'Documentar XML
        'Pikachu
        'Comentado 24/09/2003
        
        'strProcessingInstruction = "version='1.0'"
        'Set objProcessingInstruction = objDOMDocument.createProcessingInstruction("xml", strProcessingInstruction)
        'objDOMDocument.insertBefore objProcessingInstruction, objDOMDocument.childNodes.Item(0)
        
        'strProcessingInstruction = "='ISO-8859-1'"
        'Set objProcessingInstruction = objDOMDocument.createProcessingInstruction("encoding", strProcessingInstruction)
        'objDOMDocument.insertBefore objProcessingInstruction, objDOMDocument.childNodes.Item(1)
        
        'Criar Atributo padrão Operação
        Set objDomAttrib = objDOMDocument.createAttribute("Operacao")
        objDomAttrib.Text = "Consultar"
        Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
        
        'Criar Atributo padrão Objeto
        Set objDomAttrib = objDOMDocument.createAttribute("Objeto")
        objDomAttrib.Text = pstrNomeObjeto
        Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
        
        For Each objFiled In rsDisconnect.Fields
            
            Set objDOMNode = objDOMDocument.createElement(objFiled.Name)
            
            Select Case objFiled.Type
                Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                    If Not IsNull(objFiled.Value) Then
                        Select Case Mid(objFiled.Name, 1, 2)
                            Case "DH", "HO"
                                objDOMNode.Text = Format(objFiled.Value, "YYYYMMDDHHmmss")
                            Case "DT"
                                objDOMNode.Text = Format(objFiled.Value, "YYYYMMDD")
                            Case Else
                                ' 101 - Tipo do Banco de Dados diferente da classe de dados
                                lngCodigoErroNegocio = 11
                                GoTo ErrorHandler
                        End Select
                    Else
                        objDOMNode.Text = "00:00:00"
                    End If
                Case Else
                    If Not IsNull(objFiled.Value) Then
                        objDOMNode.Text = objFiled.Value
                    Else
                        objDOMNode.Text = ""
                    End If
            End Select
            
            Call objDOMDocument.documentElement.appendChild(objDOMNode)
            
            Set objDomAttrib = objDOMDocument.createAttribute("Type")
            objDomAttrib.Text = objFiled.Type
            Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
            
            Set objDomAttrib = objDOMDocument.createAttribute("Tamanho")
            
            If objFiled.Type = adChar Or objFiled.Type = adLongVarChar Or objFiled.Type = adVarChar Then
                objDomAttrib.Text = objFiled.DefinedSize
            Else
                objDomAttrib.Text = objFiled.Precision
            End If
            
            Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
        
        Next
        QueryXMLLer = objDOMDocument.xml
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
    Set objDOMDocument = Nothing
    
    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    QueryXMLLer = vbNullString

    Set rsDisconnect = Nothing
    Set objDOMDocument = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QueryXMLLer", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)
    
End Function

Public Function Propriedades(ByVal pstrNomeXML As String, _
                             ByVal pstrSQL As String, _
                             ByVal pstrNomeObjeto As String) As String

Dim objDOMDocument                          As MSXML2.DOMDocument40
Dim objDOMNode                              As MSXML2.IXMLDOMNode
Dim objDomAttrib                            As MSXML2.IXMLDOMAttribute

Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field

On Error GoTo ErrorHandler
    
    Call fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set objDOMDocument = CreateObject("MSXML2.DOMDocument.4.0")
        
    Call objADODBConnection.Open(gstrConexao)
    
    objADODBConnection.CursorLocation = ADODB.adUseClient
    rsDisconnect.CursorType = ADODB.adOpenStatic
    rsDisconnect.LockType = ADODB.adLockBatchOptimistic
    
    Set rsDisconnect = objADODBConnection.Execute(pstrSQL & " WHERE 1 = 2 ", , ADODB.adCmdText)
    
    Set objDOMNode = objDOMDocument.createElement(pstrNomeXML)
    Call objDOMDocument.appendChild(objDOMNode)
    Set objDomAttrib = objDOMDocument.createAttribute("Operacao")
    objDomAttrib.Text = "Consultar"
    
    Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
    Set objDomAttrib = objDOMDocument.createAttribute("Objeto")
    objDomAttrib.Text = pstrNomeObjeto
    Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
    
    For Each objFiled In rsDisconnect.Fields
        Set objDOMNode = objDOMDocument.createElement(objFiled.Name)
        Call objDOMDocument.documentElement.appendChild(objDOMNode)
        Set objDomAttrib = objDOMDocument.createAttribute("Type")
        objDomAttrib.Text = objFiled.Type
        Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
        
        Set objDomAttrib = objDOMDocument.createAttribute("Tamanho")
        
        If objFiled.Type = adChar Or objFiled.Type = adLongVarChar Or objFiled.Type = adVarChar Then
            objDomAttrib.Text = objFiled.DefinedSize
        Else
            objDomAttrib.Text = objFiled.Precision
        End If
        
        Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
        
    Next
        
    Propriedades = objDOMDocument.xml
    
    Set objDOMDocument = Nothing
    
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
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Propriedades = vbNullString
    
    Set objDOMDocument = Nothing
    Set rsDisconnect = Nothing
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Propriedades", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)
    
End Function

Public Function QueryXMLLerTodos(ByVal pstrNomeXML As String, _
                                 ByVal pstrSQL As String, _
                                 ByVal pstrNomeObjeto As String, _
                        Optional ByVal pblnReturnDataType As Boolean = True)

Dim objDOMDocument                          As MSXML2.DOMDocument40
Dim objDOMNode                              As MSXML2.IXMLDOMNode
Dim objDOMNodeAux                           As MSXML2.IXMLDOMNode
Dim objDomAttrib                            As MSXML2.IXMLDOMAttribute
    
Dim objADODBConnection                      As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objFiled                                As ADODB.Field

On Error GoTo ErrorHandler
    
    fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set objDOMDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call objADODBConnection.Open(gstrConexao)
    
    objADODBConnection.CursorLocation = ADODB.adUseClient
    rsDisconnect.CursorType = ADODB.adOpenStatic
    rsDisconnect.LockType = ADODB.adLockBatchOptimistic
    
    Set rsDisconnect = objADODBConnection.Execute(pstrSQL, , ADODB.adCmdText)
    
    If rsDisconnect.RecordCount > 0 Then
        
        Set objDOMNode = objDOMDocument.createElement("Repeat_" & pstrNomeXML)
        Call objDOMDocument.appendChild(objDOMNode)
        
        Set objDOMNodeAux = objDOMDocument.createElement("Grupo_" & pstrNomeXML)
        Call objDOMNode.appendChild(objDOMNodeAux)
        
        Set objDOMNode = objDOMNodeAux
        Set objDomAttrib = objDOMDocument.createAttribute("Operacao")
        objDomAttrib.Text = "Consultar"
        
        Call objDOMNodeAux.Attributes.setNamedItem(objDomAttrib)
        
        Set objDOMNode = objDOMNodeAux
        Set objDomAttrib = objDOMDocument.createAttribute("Objeto")
        objDomAttrib.Text = pstrNomeObjeto
        
        Call objDOMNodeAux.Attributes.setNamedItem(objDomAttrib)
        
        Do While Not rsDisconnect.EOF
                   
            For Each objFiled In rsDisconnect.Fields
                
                Set objDOMNodeAux = objDOMDocument.createElement(objFiled.Name)
                
                Select Case objFiled.Type
                    
                    Case Is = ADODB.adDate, ADODB.adDBDate, ADODB.adDBTimeStamp
                        
                        If Not IsNull(objFiled.Value) Then
                            Select Case Mid(objFiled.Name, 1, 2)
                                Case "DH", "HO"
                                    objDOMNodeAux.Text = Format(objFiled.Value, "YYYYMMDDHHmmss")
                                Case "DT"
                                    objDOMNodeAux.Text = Format(objFiled.Value, "YYYYMMDD")
                                Case Else
                                    ' 101 - Tipo do Banco de Dados diferente da classe de dados
                                    lngCodigoErroNegocio = 11
                                    GoTo ErrorHandler
                            End Select
                        Else
                            objDOMNodeAux.Text = "00:00:00"
                        End If
                    Case Is = ADODB.adNumeric, ADODB.adDouble, ADODB.adVarNumeric, ADODB.adCurrency, ADODB.adDecimal
                        If Not IsNull(objFiled.Value) Then
                            'Coloca no padrão de separação de decimais em XML => vírgula (,)
                            objDOMNodeAux.Text = Replace$(objFiled.Value, ".", ",")
                        Else
                            objDOMNodeAux.Text = 0
                        End If
                    Case Else
                        
                        If Not IsNull(objFiled.Value) Then
                            objDOMNodeAux.Text = Trim(Replace(objFiled.Value, vbNullChar, vbNullString))
                            objDOMNodeAux.Text = Replace(objFiled.Value, Chr(2), vbNullString)
                        Else
                            objDOMNodeAux.Text = ""
                        End If
                
                End Select
                
                If pblnReturnDataType Then
                    Set objDomAttrib = objDOMDocument.createAttribute("Type")
                    objDomAttrib.Text = objFiled.Type
                    Call objDOMNodeAux.Attributes.setNamedItem(objDomAttrib)
                End If
                
                Call objDOMNode.appendChild(objDOMNodeAux)
            Next
            
            rsDisconnect.MoveNext
            
            If Not rsDisconnect.EOF Then
                Set objDOMNode = objDOMDocument.createElement("Grupo_" & pstrNomeXML)
                Call objDOMDocument.childNodes.Item(0).appendChild(objDOMNode)
                Set objDomAttrib = objDOMDocument.createAttribute("Operacao")
                objDomAttrib.Text = "Consultar"
                Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
                
                Set objDomAttrib = objDOMDocument.createAttribute("Objeto")
                objDomAttrib.Text = pstrNomeObjeto
                Call objDOMNode.Attributes.setNamedItem(objDomAttrib)
            End If
        Loop
        QueryXMLLerTodos = objDOMDocument.xml
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
    Set objDOMDocument = Nothing
    
    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    QueryXMLLerTodos = vbNullString
    
    Set rsDisconnect = Nothing
    Set objDOMDocument = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QueryXMLLerTodos", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & pstrSQL)

End Function

Public Function SelectVarchar4000(ByVal pstrNomeTabela As String, _
                                  ByVal pstrNomeCampoSequencia As String, _
                                  ByVal plngSequencial As Long, _
                                  ByVal pstrNomeCampoVarchar As String, _
                                  ByVal pstrNomeCampoOrdem As String) As String
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
            strConteudoVarchar = strConteudoVarchar & objRs.Fields(0)
            objRs.MoveNext
        Wend
    End If
    
    objRs.Close
        
    If Trim(strConteudoVarchar) <> "" Then
        SelectVarchar4000 = fgBase64Decode(strConteudoVarchar)
    End If
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterIdentificadorRemessaPJ(ByVal pintOpcaoPJ_NZ As Integer, _
                                            ByVal plngCodigoEmpresa As Long, _
                                            ByVal plngQuantMensagem As Long, _
                                            ByRef pstrControleInicio As String, _
                                   Optional ByRef pstrControleFim As String) As Boolean

Dim objADODBCommand                         As ADODB.Command
Dim objADODBParameter                       As ADODB.Parameter
Dim objADODBConnection                      As ADODB.Connection

Dim strComplemento                          As String

On Error GoTo ErrorHandler

    fgObterVariaveisAmbiente

    Set objADODBConnection = CreateObject("ADODB.Connection")
    
    Call objADODBConnection.Open(gstrConexao)

    Set objADODBCommand = New ADODB.Command
    Set objADODBCommand.ActiveConnection = objADODBConnection
    Set objADODBParameter = New ADODB.Parameter

    objADODBParameter.Name = "CO_OPCA"
    objADODBParameter.Direction = adParamInput
    objADODBParameter.Type = adNumeric
    objADODBParameter.Value = pintOpcaoPJ_NZ

    objADODBCommand.Parameters.Append objADODBParameter
    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "CO_EMPRE"
    objADODBParameter.Direction = adParamInput
    objADODBParameter.Type = adNumeric
    objADODBParameter.Value = plngCodigoEmpresa

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "SG_SISTE"
    objADODBParameter.Direction = adParamInput
    objADODBParameter.Type = adVariant
    objADODBParameter.Value = "A8"

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "DT_MOVT"
    objADODBParameter.Direction = adParamInput
    objADODBParameter.Type = adDate
    objADODBParameter.Value = Date

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "QT_NU_SEQU"
    objADODBParameter.Direction = adParamInput
    objADODBParameter.Type = adNumeric
    objADODBParameter.Value = plngQuantMensagem

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "NU_SEQU_INIC"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adVariant

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "NU_SEQU_FINA"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adVariant

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "TABELA_ORACLE"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adVariant

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "FUNCAO_ORACLE"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adVariant
 
    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "SQL_CODE"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adNumeric

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "RC_ROTINA"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adNumeric

    objADODBCommand.Parameters.Append objADODBParameter

    Set objADODBParameter = Nothing
    Set objADODBParameter = New ADODB.Parameter
    objADODBParameter.Name = "MSG_RC_ROTINA"
    objADODBParameter.Direction = adParamOutput
    objADODBParameter.Type = adVariant
 
    objADODBCommand.Parameters.Append objADODBParameter

    objADODBCommand.CommandText = "A8PROC.A8P_NZO0040"
    objADODBCommand.CommandType = adCmdStoredProc
    objADODBCommand.Execute

    If objADODBCommand.Parameters("RC_ROTINA").Value = 0 Then
        pstrControleInicio = objADODBCommand.Parameters("NU_SEQU_INIC").Value
        pstrControleFim = objADODBCommand.Parameters("NU_SEQU_FINA").Value
    Else
        ' 39 - Erro na execução da Procedure A8P_NZO0040.
        lngCodigoErroNegocio = 39
        strComplemento = objADODBCommand.Parameters("RC_ROTINA").Value & " - " & objADODBCommand.Parameters("MSG_RC_ROTINA").Value
        GoTo ErrorHandler
    End If

    Set objADODBCommand = Nothing
    Set objADODBParameter = Nothing

    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If

    Set objADODBConnection = Nothing

Exit Function

ErrorHandler:

    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If

    Set objADODBCommand = Nothing
    Set objADODBParameter = Nothing
    Set objADODBConnection = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterIdentificadorRemessaPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro, strComplemento)
End Function

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
    
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    Set objCommand = Nothing
    Set objADODBConnection = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro, pstrNomeProc)
    
End Function
