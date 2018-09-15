Attribute VB_Name = "basA6A7A8VerificaServer"

'Funções genéricas e Atalhos para utilização de outros objetos

Option Explicit

Private Enum enumSourceAlerta
    Oracle = 1
    MQSeries = 2
    MSMQ = 3
    Todos = 4
End Enum


Public Enum enumMQOO_Open
    MQOO_BROWSE = 8
    MQOO_INPUT_SHARED = 2
    MQOO_INPUT_EXCLUSIVE = 4
    MQOO_OUTPUT = 16
End Enum

Public Enum enumPutOptions
    Binario = 1
    MainFrame = 2
End Enum

Public Const glngMaxProcCount                     As Long = 50


'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer


'Gerar informações de alerta de erro do A7

Public Function fgGeraInformacaoAlerta(ByVal pstrDescricaoOcorrencia As String, _
                                       ByVal pstrDataHoraOcorrencia As String, _
                                       ByVal pstrSorce As String, _
                                       ByVal pstrErro As String) As Boolean

Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlAlertaAux                            As MSXML2.DOMDocument40
Dim strxmlAlerta                            As String

On Error GoTo ErrorHandler
    
    strxmlAlerta = fgObterInformacaoAlerta()
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAlertaAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlAlertaAux, "", "Grupo_Alerta", "")
    Call fgAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_OCOR", pstrDescricaoOcorrencia)
    Call fgAppendNode(xmlAlertaAux, "Grupo_Alerta", "DH_OCOR", pstrDataHoraOcorrencia)
    Call fgAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_FONT_ERRO", pstrSorce)
    Call fgAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_ERRO", pstrErro)
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then
        Call fgAppendNode(xmlAlerta, "", "Alerta", "")
        Call fgAppendNode(xmlAlerta, "Alerta", "Repet_Alerta", "")
    Else
        xmlAlerta.loadXML strxmlAlerta
    End If
    
    Call fgAppendXML(xmlAlerta, "Repet_Alerta", xmlAlertaAux.xml)
    
    Call flGerarPropriedadeAlerta(xmlAlerta.xml)
    
    Set xmlAlerta = Nothing
    Set xmlAlertaAux = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlAlerta = Nothing
    Set xmlAlertaAux = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Ler as informações de alerta de erro do A7 no Shared Property Memory (COM+)

Public Function fgObterInformacaoAlerta() As String

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    fgObterInformacaoAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Gravar as informações de alerta de erro do A7 no Shared Property Memory (COM+)

Private Function flGerarPropriedadeAlerta(ByVal pstrxmlAlerta As String)

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    Call objSharedPropMem.SetSPMProperty("A7BUSALERTA", "ALERTA", pstrxmlAlerta)
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Sub flGravaArquivoOra(ByVal psErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer

On Error GoTo ErrorHandler

    strNomeArquivoLogErro = App.Path & "/log/" & "Oracle_" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, psErro
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'---------------------------------------- COPY EXE ----------------


Public Function fgVerificarBancoDados() As Boolean

Dim objA6A7A8CA                             As Object
Dim strTextoIncialErroDB                    As String
Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strErro                                 As String
Dim strErroAux                              As String
Dim strErroMQ                               As String

On Error GoTo ErrorHandler
    
    Set objA6A7A8CA = CreateObject("A6A7A8CA.clsConsulta")
    strTextoIncialErroDB = "Verificar Banco de Dados do SLCC"
    Call objA6A7A8CA.QuerySQL("SELECT * FROM DUAL")
        
    Set objA6A7A8CA = Nothing
    
    'Se a conexão com Oracle estiver OK e existir alerta de Oracle, exlcuir Alerta
    If flExisteInformacaoAlerta(enumSourceAlerta.Oracle) Then
        'Excluir alerta de Oracle
        Call flExcluirInformacaoAlerta(enumSourceAlerta.Oracle)
    End If
    
    fgVerificarBancoDados = True
    
    Exit Function
    
ErrorHandler:
    
    fgVerificarBancoDados = False
    
    strErroAux = Err.Description

    Set objA6A7A8CA = Nothing
    
    If flExisteInformacaoAlerta(enumSourceAlerta.Oracle) Then Exit Function
    
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlErro.loadXML(strErroAux) Then
        
        Call flGeraInformacaoAlerta(strTextoIncialErroDB, _
                                    Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                    "Oracle", _
                                    strErro)
        
        fgGravaArquivo "ORACLE", strTextoIncialErroDB & vbCrLf & strErroAux
        
    Else
        
        strErro = xmlErro.selectSingleNode("//Description").Text
        
        strErroMQ = fgGetError(strErroAux, "QM.SLCC.01", "")
            
        If Trim(strErroMQ) <> "" Then
            strErro = strErroMQ
        End If
                   
        fgGravaArquivo "SLCC", strTextoIncialErroDB & vbCrLf & _
                       strErroMQ & vbCrLf & _
                       String(50, "*") & vbCrLf & _
                       strErroAux
    
    End If
   
    Set xmlErro = Nothing

End Function

'Verifica se existe informação de alerta de errro do A7

Private Function flExisteInformacaoAlerta(ByVal plngSouce As enumSourceAlerta) As Boolean

Dim objSharedPropMem                        As Object
Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String
Dim blnExisteAlerta                         As Boolean
Dim strAux                                  As String

On Error GoTo ErrorHandler
        
    blnExisteAlerta = False
    
'    Exit Function
        
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    strxmlAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then Exit Function
    
    xmlAlerta.loadXML strxmlAlerta
        
    For Each xmlNode In xmlAlerta.selectNodes("Alerta/Repet_Alerta/*")
             
         Select Case plngSouce
            Case enumSourceAlerta.MQSeries
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MQSERIES" Then
                    blnExisteAlerta = True
                    Exit For
                End If
            Case enumSourceAlerta.MSMQ
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MSMQ" Then
                    blnExisteAlerta = True
                    Exit For
                End If
            Case enumSourceAlerta.Oracle
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "ORACLE" Then
                    blnExisteAlerta = True
                    Exit For
                End If
            Case enumSourceAlerta.Todos
            
                strAux = strAux & xmlNode.selectSingleNode("DE_FONT_ERRO").Text
         End Select
    Next
    
    If plngSouce = enumSourceAlerta.Todos Then
        If strAux <> vbNullString Then
            blnExisteAlerta = True
        End If
    End If
    
    flExisteInformacaoAlerta = blnExisteAlerta
    
    Set xmlAlerta = Nothing
    Set xmlNode = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlNode = Nothing
    Set xmlAlerta = Nothing
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Exclusão de infrmações de alerta de erro do A7

Private Function flExcluirInformacaoAlerta(ByVal plngSouce As enumSourceAlerta) As Boolean

Dim objSharedPropMem                        As Object
Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String
Dim blnExisteAlerta                         As Boolean

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    strxmlAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then Exit Function
    
    xmlAlerta.loadXML strxmlAlerta
        
    For Each xmlNode In xmlAlerta.selectNodes("Alerta/Repet_Alerta/*")
         Select Case plngSouce
            Case enumSourceAlerta.MQSeries
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MQSERIES" Then
                    xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").removeChild xmlNode
                End If
            Case enumSourceAlerta.MSMQ
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MSMQ" Then
                    xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").removeChild xmlNode
                End If
            Case enumSourceAlerta.Oracle
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "ORACLE" Then
                    xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").removeChild xmlNode
                End If
         End Select
    Next
    
    Call flGerarPropriedadeAlerta(xmlAlerta.xml)
    
    flExcluirInformacaoAlerta = True
    
    Set xmlAlerta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlAlerta = Nothing
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Gerar informações de informações de alerta de erro no A7

Private Function flGeraInformacaoAlerta(ByVal pstrDescricaoOcorrencia As String, _
                                        ByVal pstrDataHoraOcorrencia As String, _
                                        ByVal pstrSorce As String, _
                                        ByVal pstrErro As String) As Boolean

Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlAlertaAux                            As MSXML2.DOMDocument40
Dim strxmlAlerta                            As String

On Error GoTo ErrorHandler
    
    strxmlAlerta = flObterInformacaoAlerta()
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAlertaAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call flAppendNode(xmlAlertaAux, "", "Grupo_Alerta", "")
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_OCOR", pstrDescricaoOcorrencia)
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DH_OCOR", pstrDataHoraOcorrencia)
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_FONT_ERRO", pstrSorce)
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_ERRO", pstrErro)
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then
        Call flAppendNode(xmlAlerta, "", "Alerta", "")
        Call flAppendNode(xmlAlerta, "Alerta", "Repet_Alerta", "")
    Else
        xmlAlerta.loadXML strxmlAlerta
    End If
    
    Call flAppendXML(xmlAlerta, "Repet_Alerta", xmlAlertaAux.xml)
    
    Call flGerarPropriedadeAlerta(xmlAlerta.xml)
    
    Set xmlAlerta = Nothing
    Set xmlAlertaAux = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlAlerta = Nothing
    Set xmlAlertaAux = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


'Gravar aqruivos de log

Public Sub fgGravaArquivo(ByVal nomeArq As String, ByVal pstrErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strMensagem                              As String

On Error GoTo ErrorHandler

    strMensagem = String(50, "*") & vbCrLf
    strMensagem = strMensagem & pstrErro & vbCrLf
    strMensagem = strMensagem & String(50, "*")

    strNomeArquivoLogErro = App.Path & "\log\" & nomeArq & "_" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, strMensagem
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'Verifica se o erro é de MQSeries e Oracle

Public Function fgGetError(ByVal pstrErro As String, _
                           ByVal pstrQMgrName As String, _
                           ByVal pstrQueueName As String) As String

Dim objDomError                             As MSXML2.DOMDocument40
Dim intErrPosicaoInicio                     As Integer
Dim intErrPosicaoFim                        As Integer
Dim strErrSource                            As String
Dim strErrDesciption                        As String
Dim strTxtPesquisa                          As String
Dim lngErrNumber                            As Long
Dim strCodOcorrencia                        As String
Dim strOcorrencia                           As String
Dim strComplemento                          As String
    
On Error GoTo ErrorHandler
    
    strErrDesciption = pstrErro
    
    If InStr(1, UCase(strErrDesciption), UCase("mqax200")) > 0 Then
        
        strTxtPesquisa = "ReasonCode = "
        
        intErrPosicaoInicio = InStr(1, pstrErro, strTxtPesquisa)
        strErrSource = "MQSeries"
        
        If intErrPosicaoInicio > 0 Then
            intErrPosicaoInicio = intErrPosicaoInicio + Len(strTxtPesquisa)
            intErrPosicaoFim = InStr(intErrPosicaoInicio, pstrErro, ",")
            lngErrNumber = Mid(pstrErro, intErrPosicaoInicio, intErrPosicaoFim - intErrPosicaoInicio)
            strCodOcorrencia = "MQ-" & Format(lngErrNumber, "00000")
        End If
       
        Select Case lngErrNumber
            Case 2009
                strComplemento = "Queue Manager name: " & pstrQMgrName
                strOcorrencia = "A conexão com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2035
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
                strOcorrencia = "O usuário não está autorizado a executar a operação(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2058, 2059
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = "O gerenciador de fila não está disponível para conexão(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2085, 2086
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = "Fila não Criada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2162
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = "o gerenciador de fila está sendo encerrando(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case Else
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = strErrDesciption & " - (" & strComplemento & ")"
                fgGetError = strOcorrencia
        End Select
        
    ElseIf InStr(1, UCase(strErrDesciption), UCase("Microsoft OLE DB Provider for Oracle")) > 0 Then
        
        strTxtPesquisa = "ORA"
        
        intErrPosicaoInicio = InStr(1, pstrErro, strTxtPesquisa)
        
        strErrSource = "Oracle"
        
        If intErrPosicaoInicio > 0 Then
            intErrPosicaoInicio = intErrPosicaoInicio
            intErrPosicaoFim = InStr(intErrPosicaoInicio, pstrErro, ":", vbBinaryCompare)
            strCodOcorrencia = Mid(pstrErro, intErrPosicaoInicio, intErrPosicaoFim - intErrPosicaoInicio)
        End If
        
        Select Case strCodOcorrencia
            Case "ORA-12545"
                strOcorrencia = "ORA-12545: A conexão falhou porque o objeto ou host de destino não existe."
                fgGetError = strOcorrencia
            
            Case "ORA-12154"
                strOcorrencia = "ORA-12154: TNS - Não foi possível resorver nome de serviço."
                fgGetError = strOcorrencia
            
            Case "ORA-12541"
                strOcorrencia = "ORA-12541: TNS - não há listener."
                fgGetError = strOcorrencia
            
            Case "ORA-12500"
                strOcorrencia = "ORA-12500: TNS - Listener falhou ao iniciar um processo de servidor dedicado."
                fgGetError = strOcorrencia
            
            Case "ORA-00018"
                strOcorrencia = "ORA-00020: Maximum number of sessions exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-00019"
                strOcorrencia = "ORA-00019: Maximum number of session licenses exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-00022"
                strOcorrencia = "ORA-00022: Invalid session ID; access denied."
                fgGetError = strOcorrencia
            
            Case "ORA-00028"
                strOcorrencia = "ORA-00028 your session has been killed."
                fgGetError = strOcorrencia
            
            Case "ORA-00029"
                strOcorrencia = "ORA-00029 session is not a user session."
                fgGetError = strOcorrencia

            Case "ORA-00035"
                strOcorrencia = "ORA-00035 LICENSE_MAX_USERS cannot be less than current number of users."
                fgGetError = strOcorrencia
            
            Case "ORA-00036"
                strOcorrencia = "ORA-00036 maximum number of recursive SQL levels (string) exceeded."
                fgGetError = strOcorrencia

            Case "ORA-00052"
                strOcorrencia = "ORA-00052 maximum number of enqueue resources (string) exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-00052"
                strOcorrencia = "ORA-00052 maximum number of enqueue resources (string) exceeded"
                fgGetError = strOcorrencia
            
            Case "ORA-00053"
                strOcorrencia = "ORA-00053 maximum number of enqueues exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00055"
                strOcorrencia = "ORA-00055 maximum number of DML locks exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00057"
                strOcorrencia = "ORA-00057 maximum number of temporary table locks exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00063"
                strOcorrencia = "ORA-00063 maximum number of LOG_FILES exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00107"
                strOcorrencia = "ORA-00107 failed to connect to ORACLE listener process"
                fgGetError = strOcorrencia
                
            Case "ORA-00115"
                strOcorrencia = "ORA-00115 connection refused; dispatcher connection table is full"
                fgGetError = strOcorrencia
                
            Case "ORA-00152"
                strOcorrencia = "ORA-00152 current session does not match requested session"
                fgGetError = strOcorrencia
                
            Case "ORA-00603"
                strOcorrencia = "ORA-00603 ORACLE server session terminated by fatal error"
                fgGetError = strOcorrencia
                
            Case "ORA-00606"
                strOcorrencia = "ORA-00606 Internal error code"
                fgGetError = strOcorrencia
                
            Case "ORA-00816"
                strOcorrencia = "ORA-00816 error message translation failed"
                fgGetError = strOcorrencia

            Case "ORA-00020"
                strOcorrencia = "ORA-00020: maximum number of processes exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-01034", "ORA-12535", "ORA-12560", "ORA-12541"
                strOcorrencia = "ORA-01034: ORACLE não disponível."
                fgGetError = strOcorrencia
        
            Case Else
                strOcorrencia = "Verificar Banco de Dados Oracle - " & strErrDesciption
                fgGetError = strOcorrencia
        End Select
    
    End If
        
    If strOcorrencia <> "" Then
        Call flGeraInformacaoAlerta(strOcorrencia, _
                                    Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                    strErrSource, _
                                    strErrDesciption)
    End If
        
    Set objDomError = Nothing
    
    Exit Function

ErrorHandler:
       
    Set objDomError = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



'Obter informações de alerta de errro do A7 no Shared Property memory (COM+)

Private Function flObterInformacaoAlerta() As String

Dim objSharedPropMem                        As Object
Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    flObterInformacaoAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
        
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Adicionar um Node a um xml

Private Function flAppendNode(ByRef objDomDocument As MSXML2.DOMDocument40, _
                              ByVal psNodeContext As String, _
                              ByVal psNodeNome As String, _
                              ByVal psNodeValor As String, _
                     Optional ByVal psNodeRepet As String = "") As Boolean

Dim objDOMNodeAux                           As MSXML2.IXMLDOMNode
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    If psNodeContext = "" Then
        'Se a Tag for o root passar psNodeContext = Nome Tag Principal
        Set objDomNodeContext = objDomDocument
    Else
        'Para criar um Grupo (Ex.:Grupo_Usuario) dentro de Uma Repeticao (Ex. Repeat_Usuario)
        'Passar o argumento pNomeRepet = Repet_Usuario
        If psNodeRepet <> "" Then
            Set objDomNodeContext = objDomDocument.selectSingleNode("//Repet_Origem").childNodes.Item(objDomDocument.selectSingleNode("//Repet_Origem").childNodes.length - 1)
        Else
            Set objDomNodeContext = objDomDocument.documentElement.selectSingleNode("//" & psNodeContext)
        End If
    End If
    
    Set objDOMNodeAux = objDomDocument.createElement(psNodeNome)
    objDOMNodeAux.Text = psNodeValor
    objDomNodeContext.appendChild objDOMNodeAux

    Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.Description

End Function


'Adicional um estrutura de xml em um node de um xml

Private Sub flAppendXML(ByRef objMapaNavegacao As MSXML2.DOMDocument30, _
                        ByVal pstrNodeContext As String, _
                        ByVal pstrXMLFilho As String)

Dim objDomFilho                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    'Leitura do XML Filho
    Set objDomFilho = CreateObject("MSXML2.DOMDocument.4.0")
    If Not objDomFilho.loadXML(pstrXMLFilho) Then
        '100 - Documento XML Inválido.
        'llCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
       
    'Setar o nivel que deverá entrar o XML Filho
     Set objDomNode = objMapaNavegacao.documentElement.selectSingleNode("//" & pstrNodeContext)

    'Adicionar XML Filho na Saida
    objDomNode.appendChild objDomFilho.childNodes.Item(0)
    
    Set objDomFilho = Nothing
    Set objDomNode = Nothing
    Exit Sub

ErrorHandler:
    
    Set objDomFilho = Nothing
    Set objDomNode = Nothing
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd

End Sub


Public Function fgVerificaMQSeries(strQMName As String) As Boolean

Dim strErro                                 As String
Dim strErroMQ                               As String
Dim objMQSession                            As Object 'MQAX200.MQSession
Dim objMQQueueManager                       As Object 'MQAX200.MQQueueManager

On Error GoTo ErrorHandler
        
    'Pikachu - 29/07/2005
    'Conectar QueueManager somente uma vez , na inicialização
    
    Set objMQSession = CreateObject("MQAX200.MQSession")
    Set objMQQueueManager = objMQSession.AccessQueueManager(strQMName)
    
    fgVerificaMQSeries = True
            
    If flExisteInformacaoAlerta(enumSourceAlerta.MQSeries) Then
        'Excluir alerta de MSMQ
        Call flExcluirInformacaoAlerta(enumSourceAlerta.MQSeries)
    End If
            
    Set objMQSession = Nothing
    Set objMQQueueManager = Nothing
            
    Exit Function
ErrorHandler:
    
    strErro = Err.Description
    
    Set objMQSession = Nothing
    Set objMQQueueManager = Nothing
    
    fgVerificaMQSeries = False
        
    If flExisteInformacaoAlerta(enumSourceAlerta.MQSeries) Then Exit Function
    
    Call flGeraInformacaoAlerta("Verificar MQSeries (Queue Manager: " & strQMName & ")", _
                                Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                "MQSeries", _
                                strErro)
        
    fgGravaArquivo "MQSERIES", "Verificar Gerenciador de Filas - MQSeries" & vbCrLf & _
                   "Queue Manager: " & strQMName & vbCrLf & _
                   strErro
        
       
End Function

'-----------------------------------------------------------------------
'Put de mensagens de erro nas filas de erro
'-----------------------------------------------------------------------

Public Function PutFilaErro(ByVal pstrNomeFilaGet As String, _
                            ByVal pstrNomeFilaErro As String, _
                            ByVal pstrMessageIdHex As String, _
                            ByRef pstrLogErro As String, _
                            ByRef pstrCorrelationID As String) As Boolean
                         
Dim objMQAX200                              As Object  'A7Server.clsFilaErro
                         
On Error GoTo ErrorHandler
        
    Set objMQAX200 = CreateObject("A7Server.clsFilaErro")
    
    Call objMQAX200.PutFilaErro(pstrNomeFilaGet, _
                                pstrNomeFilaErro, _
                                pstrMessageIdHex, _
                                pstrLogErro, _
                                pstrCorrelationID)
    Set objMQAX200 = Nothing
    
    Exit Function

ErrorHandler:
   
    Set objMQAX200 = Nothing
    
    Err.Clear
    
End Function

