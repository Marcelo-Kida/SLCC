Attribute VB_Name = "basA6A7A8VerificaServer"

'Funções genéricas e Atalhos para utilização de outros objetos

Option Explicit

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


'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

'Verificaçã se o erro é proveniente de erro de MQSeries ou Oracle

Public Function fgGetError(ByVal pstrErro As String) As String

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

    Set objDomError = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDomError.loadXML(pstrErro) Then
       pstrErro = ""
       Exit Function
    End If

    strErrSource = objDomError.selectSingleNode("Erro/Grupo_ErrorInfo/Source").Text
    strErrDesciption = objDomError.selectSingleNode("Erro/Grupo_ErrorInfo/Description").Text
    strComplemento = objDomError.selectSingleNode("//Complemento").Text
    
    If UCase(strErrSource) = UCase("mqax200") Then
        
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
                strOcorrencia = "A conexão com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2035
                strOcorrencia = "O usuário não está autorizado a executar a operação tentada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2058
                strOcorrencia = "O gerenciador de fila não está disponível para conexão(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2059
                strOcorrencia = "O gerenciador de fila não está disponível para conexão(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2085
                strOcorrencia = "Fila não Criada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2086
                strOcorrencia = "Fila não Criada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2009
                strOcorrencia = "A conexão com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2162
                strOcorrencia = "O gerenciador de fila está sendo encerrando(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2072
                strOcorrencia = ""
                fgGetError = ""
            Case Else
                strOcorrencia = strErrDesciption & "(" & strComplemento & ")."
                fgGetError = strOcorrencia
        End Select
                
                        
    ElseIf UCase(Trim(strErrSource)) = UCase("Microsoft OLE DB Provider for Oracle") Then
        
        strTxtPesquisa = "ORA"
        
        intErrPosicaoInicio = InStr(1, pstrErro, strTxtPesquisa)
        
        strErrSource = "Oracle"
        
        If intErrPosicaoInicio > 0 Then
            intErrPosicaoInicio = intErrPosicaoInicio
            intErrPosicaoFim = InStr(intErrPosicaoInicio, pstrErro, ":", vbBinaryCompare)
            strCodOcorrencia = Mid(pstrErro, intErrPosicaoInicio, intErrPosicaoFim - intErrPosicaoInicio)
        End If
        
'"ORA-02291- integrity constraint (palavra chave) violated - parent key not found"
'"ORA-00001- unique constraint (palavra-chave) violated"
'"ORA-02292- integrity constraint (palavra chave) violated - child record found"
        
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
            Case "ORA-03114"
                strOcorrencia = "ORA-03114 not connected to ORACLE."
                fgGetError = strOcorrencia
            Case "ORA-12547"
                strOcorrencia = "ORA-12547 TNS:lost contact."
                fgGetError = strOcorrencia
            Case "ORA-00029"
                strOcorrencia = "ORA-00029 session is not a user session."
                fgGetError = strOcorrencia
            Case "ORA-03113"
                strOcorrencia = "ORA-03113: end-of-file on communication channel."
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
        
        End Select
        
    End If
        
    If strOcorrencia <> "" Then
        Call fgGeraInformacaoAlerta(strOcorrencia, _
                                    Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                    strErrSource, _
                                    strErrDesciption)
    End If
        
    Set objDomError = Nothing
    
    Exit Function

ErrorHandler:
       
    Set objDomError = Nothing
    
    Err.Clear

End Function

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


