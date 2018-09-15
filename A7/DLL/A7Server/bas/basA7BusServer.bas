Attribute VB_Name = "basA7BusServer"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF9D2EF0366"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"
'Empresa        : Regerbanc
'Componente     : CA
'Classe         : basCA
'Data Criação   : 01-05-2001 17:13
'Objetivo       : Funções genéricas e Atalhos para utilização de outros objetos
'                 dentro do mesmo Componente
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 06/07/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

Public Type udtProtocolo
    TipoMensagem                            As String * 4
    SiglaSistemaOrigem                      As String * 3
    SiglaSistemaDestino                     As String * 3
    CodigoEmpresa                           As String * 5
End Type

Public Type udtProtocoloAux
    String                                  As String * 15
End Type

'Protocolo de comunicação com Sistema NZ
Public Type udtProtocoloNZ
    SiglaSistemaEnviouNZ          As String * 3
    CodigoMensagem                As String * 9
    ControleRemessaNZ             As String * 20
    DataRemessa                   As String * 8
    CodigoEmpresa                 As String * 5
    CodigoMoeda                   As String * 5
    FormatoMensagem               As String * 1
    AssinaturaInterna             As String * 50
    SiglaSistemaLegadoOrigem      As String * 3
    ReferenciaContabil            As String * 8
    BancoAgencia                  As String * 15
    QuantidadeMensagem            As String * 6
    NuOP                          As String * 23
    '***Este filler foi incluido pois o NuOP estava com o tamanho errado (24)
    Filler1                       As String * 1
    '****
    Filler2                       As String * 43
End Type

Public Type udtProtocoloNZAux
    String                        As String * 200
End Type


'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

Public Sub fgRaiseError(ByVal pstrComponente As String, _
                        ByVal pstrClasse As String, _
                        ByVal pstrMetodo As String, _
                        ByRef plngCodigoErroNegocio As Long, _
               Optional ByRef pintNumeroSequencialErro As Integer = 0, _
               Optional ByVal pstrComplemento As String = "")

Dim objLogErro                              As CA.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String
    
    Set objLogErro = CreateObject("CA.clsLogErro")

    objLogErro.RaiseError pstrComponente, _
                          pstrClasse, _
                          pstrMetodo, _
                          plngCodigoErroNegocio, _
                          ErrNumber, _
                          ErrSource, _
                          ErrDescription, _
                          pintNumeroSequencialErro, _
                          pstrComplemento, _
                          Err
    
    Set objLogErro = Nothing
    
    Err.Raise ErrNumber, ErrSource, ErrDescription
'**************************************************************************
End Sub

Public Function fgExecuteSQL(ByVal pstrSQL As String) As Long

Dim objTransacao                            As CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("CA.clsTransacao")
    fgExecuteSQL = objTransacao.ExecuteSQL(pstrSQL)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgQuerySQL(ByVal pstrSQL As String) As Object

Dim objConsulta                             As CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("CA.clsConsulta")
    Set fgQuerySQL = objConsulta.QuerySQL(pstrSQL)
    Set objConsulta = Nothing
    
    Exit Function

ErrHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgQuerySQL", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function


Public Function fgExecuteCMD(ByVal pstrNomeProc As String, _
                             ByVal pintPosicaoRetorno As Integer, _
                             ByRef pvntParametros() As Variant) As Variant

Dim objTransacao                            As CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("CA.clsTransacao")
    fgExecuteCMD = objTransacao.ExecuteCMD(pstrNomeProc, pintPosicaoRetorno, pvntParametros())
    Set objTransacao = Nothing

    Exit Function

ErrHandler:
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgPropriedades(ByVal pstrNomeXML As String, _
                               ByVal pstrSQL As String, _
                      Optional ByVal pstrNomeObjeto As String) As String

Dim objConsulta                            As CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("CA.clsConsulta")
    fgPropriedades = objConsulta.Propriedades(pstrNomeXML, pstrSQL, pstrNomeObjeto)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgPropriedades", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgQueryXMLLerTodos(ByVal pstrNomeXML As String, _
                                   ByVal pstrSQL As String, _
                                   ByVal pstrNomeObjeto As String, _
                          Optional ByVal pblnType As Boolean = True) As String

Dim objConsulta                             As CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("CA.clsConsulta")
    fgQueryXMLLerTodos = objConsulta.QueryXMLLerTodos(pstrNomeXML, pstrSQL, pstrNomeObjeto, pblnType)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgQueryXMLLerTodos", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgQueryXMLLer(ByVal pstrNomeXML As String, _
                              ByVal pstrSQL As String, _
                              ByVal pstrNomeObjeto As String) As String

Dim objConsulta                             As CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("CA.clsConsulta")
    fgQueryXMLLer = objConsulta.QueryXMLLer(pstrNomeXML, pstrSQL, pstrNomeObjeto)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BUSServer", "fgQueryXMLLer", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgErroLoadXML(ByRef objDOMDocument As MSXML2.DOMDocument40, _
                              ByVal pstrComponente As String, _
                              ByVal pstrClasse As String, _
                              ByVal pstrMetodo As String)
    

    Err.Raise objDOMDocument.parseError.errorCode, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, objDOMDocument.parseError.reason
    
End Function

Public Function fgDataHoraServidor(ByVal pintFormato As enumFormatoDataHora) As Date

On Error GoTo ErrorHandler

    Select Case pintFormato
        Case enumFormatoDataHora.Data
            fgDataHoraServidor = Date
        Case enumFormatoDataHora.Hora
            fgDataHoraServidor = Time
        Case enumFormatoDataHora.DataHora
            fgDataHoraServidor = Now
    End Select

Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgDataHoraServidor", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function


Public Function fgExecuteSequence(ByVal pstrNomeSequence As String) As Long

Dim objTransacao                            As CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("CA.clsTransacao")
    fgExecuteSequence = objTransacao.ExecuteSequence(pstrNomeSequence)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgExecuteSequence", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
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
        fgErroLoadXML objDomError, App.EXEName, "basA7BusServer", "fgGetError"
    End If

    strErrSource = objDomError.selectSingleNode("Erro/Grupo_ErrorInfo/Souce").Text
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
                fgGetError = ""
            Case 2086
                strOcorrencia = "Fila não Criada(" & strComplemento & ")."
                fgGetError = ""
            Case 2009
                strOcorrencia = "A conexão com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2162
                strOcorrencia = "o gerenciador de fila está sendo encerrando(" & strComplemento & ")."
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
            Case "ORA-01034"
                strOcorrencia = "ORA-01034: ORACLE não disponível."
                fgGetError = strOcorrencia
        End Select
        
    End If
        
    If strOcorrencia <> "" Then
        Call fgGeraInformacaoAlerta(strOcorrencia, Format(Now, "dd/mm/yyy HH:mm:ss"), strErrSource, strErrDesciption)
    End If
        
    Set objDomError = Nothing
    
    Exit Function

ErrorHandler:
       
    Set objDomError = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function fgInsertVarchar4000(ByVal pstrConteudoCampoVarchar As String) As Long

'RETORNA O NUMERO SEQUENCIAL

Dim objTransacao                            As CA.clsTransacao

On Error GoTo ErrorHandler
         
    Set objTransacao = CreateObject("CA.clsTransacao")
    
    fgInsertVarchar4000 = objTransacao.InsertVarchar4000("A7.TB_TEXT_XML", _
                                                         "CO_TEXT_XML", _
                                                         "TX_XML", _
                                                         pstrConteudoCampoVarchar, _
                                                         "NU_SEQU_TEXT_XML", _
                                                         "A7.CO_TEXT_XML")

    
    Set objTransacao = Nothing
    
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "InsertVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgSelectVarchar4000(ByVal plngSequencial As Long) As String

'Retorna decode base64

Dim objConsulta                             As CA.clsConsulta

On Error GoTo ErrorHandler
             
    Set objConsulta = CreateObject("CA.clsConsulta")
    
    fgSelectVarchar4000 = objConsulta.SelectVarchar4000("A7.TB_TEXT_XML", _
                                                        "CO_TEXT_XML", _
                                                        plngSequencial, _
                                                        "TX_XML", _
                                                        "NU_SEQU_TEXT_XML")
    
    Set objConsulta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgSelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Public Function fgGeraInformacaoAlerta(ByVal pstrDescricaoOcorrencia As String, _
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
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "flGeraInformacaoAlerta", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flObterInformacaoAlerta() As String

Dim objSharedPropMem                        As A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
        
    flObterInformacaoAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")

    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "flObterInformacaoAlerta", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flGerarPropriedadeAlerta(ByVal pstrxmlAlerta As String)

Dim objSharedPropMem                        As A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    
    Call objSharedPropMem.SetSPMProperty("A7BUSALERTA", "ALERTA", pstrxmlAlerta)
    
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "flGerarPropriedadeAlerta", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgObterUsuarioRede()

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    fgObterUsuarioRede = objUsuario.ObterUsuarioRede
    Set objUsuario = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgObterUsuarioRede", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgObterEstacaoTrabalhoUsuario() As String

Dim objUsuario                              As A6A7A8.clsUsuario
Dim strUsuarioRede                          As String
Dim strEstacaoTrabalho                      As String

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    
    strUsuarioRede = fgObterUsuarioRede
    
    Call objUsuario.ObterEstacaoTrabalhoUsuario(strUsuarioRede, _
                                                strEstacaoTrabalho, _
                                                lngCodigoErroNegocio)
    
    If lngCodigoErroNegocio <> 0 Then
        GoTo ErrorHandler
    End If
    
    fgObterEstacaoTrabalhoUsuario = strEstacaoTrabalho
    
    Set objUsuario = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgObterEstacaoTrabalhoUsuario", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

