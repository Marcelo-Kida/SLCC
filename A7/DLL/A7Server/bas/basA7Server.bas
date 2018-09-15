Attribute VB_Name = "basA7Server"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF9D2EF0366"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"

'Fun��es gen�ricas e Atalhos para utiliza��o de outros objetos

Option Explicit

Public Type udtProtocolo
    TipoMensagem                            As String * 9
    SiglaSistemaOrigem                      As String * 3
    SiglaSistemaDestino                     As String * 3
    CodigoEmpresa                           As String * 5
End Type

Public Type udtProtocoloAux
    String                                  As String * 20
End Type

Public Enum enumTipoMensagemEntrada
    MensagemNZA8 = 1000
    MensagemA8NZ = 1002
    MensagemErroNZA8 = 1003
    MensagemRetornoDV = 1006
    MensagemRetornoBG = 1008
    MensagemA8PJPrevisto = 1004
    MensagemA8PJRealizado = 1001
End Enum

'Header Mensagem Detalhe das mensagens com Tipo Fluxo 4 que chagem nas filas - A8Q.E.CONS_HD e A8Q.E.CONS_DT
Public Type udtHeaderDetalheMensagem
    NuOP                                    As String * 24
    CodigoEmpresa                           As String * 5
    NumeroSequenciaMensagem                 As String * 6
    TipoRegistro                            As String * 2
End Type

Public Type udtHeaderDetalheMensagemAux
    String                                  As String * 37
End Type

'Header DV
Public Type udtHeaderDV
    SiglaSistemaOrigem                      As String * 3
    SiglaSistemaDestino                     As String * 3
    CodigoEmpresa                           As String * 5
    NomePrograma                            As String * 8
    FILLER                                  As String * 30
    CodigoErroInterno                       As String * 4
    NomeErroInterno                         As String * 80
    MessageID                               As String * 48
End Type

Public Type udtHeaderDVAux
    String                                  As String * 181
End Type

'----Integra��o com o Sistema PJ
Public Type udtRemessaMovimento
    TipoRemessa                             As String * 3
    CodigoRemessa                           As String * 23
    DataRemessa                             As String * 8
    HoraRemessa                             As String * 4
    CodigoEmpresa                           As String * 5
    SiglaSistema                            As String * 3
    CodigoMoeda                             As String * 4
    CodigoBanqueiro                         As String * 12
    TipoCaixa                               As String * 3
    CodigoItemCaixa                         As String * 9
    TipoAtivoPassivo                        As String * 1
    CodigoProduto                           As String * 4
    TipoConta                               As String * 3
    CodigoSegmento                          As String * 3
    EventoFinanceiro                        As String * 3
    CodigoIndexador                         As String * 3
    CodigoLocalLiquidacao                   As String * 4
    CodigoFaixaValor                        As String * 3
    TipoMovimento                           As String * 3
    DataMovimento                           As String * 8
    HoraMovimento                           As String * 4
    TipoEntradaSaida                        As String * 1
    ValorMovimento                          As String * 19
    ValorContabil                           As String * 19
    TipoProcessamento                       As String * 1
    TipoEnvio                               As String * 1
    FILLER                                  As String * 46
End Type

Public Type udtRemessaMovimentoAux
    String                                  As String * 201
End Type

Public Type udtMaioresValores
    TipoRemessa                             As String * 3
    CodigoRemessa                           As String * 23
    DataRemessa                             As String * 8
    HoraRemessa                             As String * 4
    CodigoEmpresa                           As String * 5
    SiglaSistema                            As String * 3
    CodigoMoeda                             As String * 4
    CodigoBanqueiro                         As String * 12
    TipoCaixa                               As String * 3
    CodigoItemCaixa                         As String * 9
    CodigoProduto                           As String * 4
    TipoConta                               As String * 3
    CodigoSegmento                          As String * 3
    CodigoEventoFinanceiro                  As String * 3
    CodigoIndexador                         As String * 3
    CodigoLocalLiquidacao                   As String * 4
    TipoMovimento                           As String * 3
    DataMovimento                           As String * 8
    HoraMovimento                           As String * 4
    TipoEntradaSaida                        As String * 1
    ValorMovimento                          As String * 17
    CodigoBanco                             As String * 3
    CodigoAgencia                           As String * 5
    NumeroContaCorrente                     As String * 13
    TipoPessoa                              As String * 1
    CodigoCNPJ_CPF                          As String * 15
    NomeCliente                             As String * 64
    TipoProcessamento                       As String * 1
    TipoEnvio                               As String * 1
    FILLER                                  As String * 20
End Type

Public Type udtMaioresValoresAux
    String                                  As String * 250
End Type

Public Type udtHeaderMensagem
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
    FILLER                        As String * 44
End Type

Public Type udtHeaderMensagemAux
    String                        As String * 200
End Type


'----

'Vari�vel utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

'Retornar erro ao m�todo chamador com informa��es complementares para rastreamento do erro.

Public Sub fgRaiseError(ByVal pstrComponente As String, _
                        ByVal pstrClasse As String, _
                        ByVal pstrMetodo As String, _
                        ByRef plngCodigoErroNegocio As Long, _
               Optional ByRef pintNumeroSequencialErro As Integer = 0, _
               Optional ByVal pstrComplemento As String = "")

Dim objLogErro                              As A6A7A8CA.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String
    
    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")

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

'Executar comandos SQL como update, insert e delete.

Public Function fgExecuteSQL(ByVal pstrSQL As String) As Long

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteSQL = objTransacao.ExecuteSQL(pstrSQL)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar as senten�as de procedures na base de dados

Public Function fgExecuteCMD(ByVal pstrNomeProc As String, _
                             ByVal pintPosicaoRetorno As Integer, _
                             ByRef pvntParametros() As Variant) As Variant

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteCMD = objTransacao.ExecuteCMD(pstrNomeProc, pintPosicaoRetorno, pvntParametros())
    Set objTransacao = Nothing

    Exit Function

ErrHandler:
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter os nomes das colunas de uma tabela, retornando um xml com as colunas

Public Function fgPropriedades(ByVal pstrNomeXML As String, _
                               ByVal pstrSQL As String, _
                      Optional ByVal pstrNomeObjeto As String) As String

Dim objConsulta                            As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgPropriedades = objConsulta.Propriedades(pstrNomeXML, pstrSQL, pstrNomeObjeto)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgPropriedades", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Executar um SELECT na base de dados, retornando um objeto ADODB.Recordset

Public Function fgQuerySQL(ByVal pstrSQL As String) As Object

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    Set fgQuerySQL = objConsulta.QuerySQL(pstrSQL)
    Set objConsulta = Nothing
    
    Exit Function

ErrHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQuerySQL", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function
'Executar um SELECT na base de dados, retornando um XML contendo um registro
Public Function fgQueryXMLLer(ByVal pstrNomeXML As String, _
                              ByVal pstrSQL As String, _
                              ByVal pstrNomeObjeto As String) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgQueryXMLLer = objConsulta.QueryXMLLer(pstrNomeXML, pstrSQL, pstrNomeObjeto)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQueryXMLLer", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar um SELECT na base de dados, retornando um XML contendo grupos de resitros

Public Function fgQueryXMLLerTodos(ByVal pstrNomeXML As String, _
                                   ByVal pstrSQL As String, _
                                   ByVal pstrNomeObjeto As String, _
                          Optional ByVal pblnType As Boolean = True) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgQueryXMLLerTodos = objConsulta.QueryXMLLerTodos(pstrNomeXML, pstrSQL, pstrNomeObjeto, pblnType)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQueryXMLLerTodos", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o valor de uma sequence

Public Function fgExecuteSequence(ByVal pstrNomeSequence As String) As Long

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteSequence = objTransacao.ExecuteSequence(pstrNomeSequence)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgExecuteSequence", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Retornar a descri��o do erro atrav�s do c�digo informado.

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
                strOcorrencia = "A conex�o com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2035
                strOcorrencia = "O usu�rio n�o est� autorizado a executar a opera��o tentada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2058
                strOcorrencia = "O gerenciador de fila n�o est� dispon�vel para conex�o(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2059
                strOcorrencia = "O gerenciador de fila n�o est� dispon�vel para conex�o(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2085
                strOcorrencia = "Fila n�o Criada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2086
                strOcorrencia = "Fila n�o Criada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2009
                strOcorrencia = "A conex�o com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2162
                strOcorrencia = "o gerenciador de fila est� sendo encerrando(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2072
                strOcorrencia = ""
                fgGetError = ""
            Case Else
                strOcorrencia = strErrDesciption
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
                strOcorrencia = "ORA-12545: A conex�o falhou porque o objeto ou host de destino n�o existe."
                fgGetError = strOcorrencia
            Case "ORA-12154"
                strOcorrencia = "ORA-12154: TNS - N�o foi poss�vel resorver nome de servi�o."
                fgGetError = strOcorrencia
            Case "ORA-12541"
                strOcorrencia = "ORA-12541: TNS - n�o h� listener."
                fgGetError = strOcorrencia
            Case "ORA-12500"
                strOcorrencia = "ORA-12500: TNS - Listener falhou ao iniciar um processo de servidor dedicado."
                fgGetError = strOcorrencia
            Case "ORA-01034", "ORA-12535", "ORA-12560", "ORA-12541"
                strOcorrencia = "ORA-01034: ORACLE n�o dispon�vel."
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
    
    Err.Clear

End Function

'Decompor e inserir registros na tabela TB_TEXT_XML.

Public Function fgInsertVarchar4000(ByVal pstrConteudoCampoVarchar As String) As Long

'RETORNA O NUMERO SEQUENCIAL

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrorHandler
         
    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    
    fgInsertVarchar4000 = objTransacao.InsertVarchar4000("A7.TB_TEXT_XML", _
                                                         "CO_TEXT_XML", _
                                                         "TX_XML", _
                                                         pstrConteudoCampoVarchar, _
                                                         "NU_SEQU_TEXT_XML", _
                                                         "A7.SQ_A7_CO_TEXT_XML")

    
    Set objTransacao = Nothing
    
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "InsertVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Inserir o registro de par�metros gerais na tabela TB_TEXT_XML.

Public Function fgInsertVarchar4000ParamGerais(ByVal pstrConteudoCampoVarchar As String, _
                                      Optional ByVal pstrCodigoTextoXML As String = vbNullString, _
                                      Optional ByVal pblnConverterBase64 As Boolean = True) As Long

'RETORNA O NUMERO SEQUENCIAL

Dim objTransacao                            As A6A7A8CA.clsTransacao

    On Error GoTo ErrorHandler
         
    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    
    fgInsertVarchar4000ParamGerais = objTransacao.InsertVarchar4000("A8.TB_TEXT_XML", _
                                                                    "CO_TEXT_XML", _
                                                                    "TX_XML", _
                                                                    pstrConteudoCampoVarchar, _
                                                                    "NU_SEQU_TEXT_XML", _
                                                                    "A8.SQ_A8_CO_TEXT_XML", _
                                                                    pstrCodigoTextoXML, _
                                                                    pblnConverterBase64)
    
    Set objTransacao = Nothing
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7Server", "fgInsertVarchar4000ParamGerais", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar e compor os registros da tabela TB_TEXT_XML.

Public Function fgSelectVarchar4000(ByVal plngSequencial As Long, _
                           Optional ByVal pblnConverterBase64 As Boolean = True) As String

'Retorna decode base64

Dim objConsulta                             As A6A7A8CA.clsConsulta
Dim strTabela                               As String

On Error GoTo ErrorHandler
             
    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    
    If plngSequencial < 0 Then
        strTabela = "A7HIST.TB_TEXT_XML"
    ElseIf plngSequencial = 0 Then
        strTabela = "A8.TB_TEXT_XML"
    ElseIf plngSequencial > 0 Then
        strTabela = "A7.TB_TEXT_XML"
    End If

    fgSelectVarchar4000 = objConsulta.SelectVarchar4000(strTabela, _
                                                        "CO_TEXT_XML", _
                                                        Abs(plngSequencial), _
                                                        "TX_XML", _
                                                        "NU_SEQU_TEXT_XML", _
                                                        pblnConverterBase64)
    
    Set objConsulta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgSelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Criar texto pre-formatado com informa��es de alerta.

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

'Obt�m as informa��es de alerta nas propriedades compartilhadas do COM+.

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

'Gravar informa��es de alerta nas propriedades compartilhadas do COM+.

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

'Obter o login do usu�rio

Public Function fgObterUsuarioRede() As String

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

'Obter o nome da esta��o de trabalho de um usu�rio logado (Controle de Acesso)

Public Function fgObterEstacaoTrabalhoUsuario() As String

Dim objUsuario                              As A6A7A8.clsUsuario
Dim strUsuarioRede                          As String
Dim strEstacaoTrabalho                      As String

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    
    strUsuarioRede = fgObterUsuarioRede
    
    Call objUsuario.ObterEstacaoTrabalhoUsuario(strUsuarioRede, _
                                                strEstacaoTrabalho, _
                                                lngCodigoErroNegocio, _
                                                intNumeroSequencialErro)
    
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

'Remover caracteres n�o v�lidos da string informada.

Public Function fgLimpaCaracterInvalido(ByVal pstrConteudo As String) As String

On Error GoTo ErrorHandler
    
    pstrConteudo = Replace(pstrConteudo, vbFormFeed, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbNullChar, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbCr, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbCrLf, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbLf, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbNewLine, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbTab, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbVerticalTab, vbNullString)
    pstrConteudo = Replace(pstrConteudo, vbBack, vbNullString)
    
    
    fgLimpaCaracterInvalido = pstrConteudo
    
     Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgLimpaCaracterInvalido", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gravar arquivo de log.

Public Sub fgGravarArquivo(ByVal pstrNomeArquivo As String, _
                           ByVal pstrMensagem As String)

Dim intFile                                  As Integer

On Error GoTo ErrorHandler
    
   pstrNomeArquivo = App.Path & "\LOG\" & pstrNomeArquivo
        
    intFile = FreeFile
    Open pstrNomeArquivo For Output As intFile
    Print #intFile, pstrMensagem
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
    
End Sub

'Obter a quantidade de dias configuradas para o expurgo das tabelas.

Public Function fgObterQtdDiasExpurgo() As Long

On Error GoTo ErrorHandler

Dim strXMLParmGeral                         As String
Dim xmlParmGeral                            As DOMDocument40

    strXMLParmGeral = fgSelectVarchar4000(0, False)
    Set xmlParmGeral = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlParmGeral.loadXML(strXMLParmGeral) Then
        GoTo ErrorHandler
    End If
    
    If xmlParmGeral.selectSingleNode("//BASE_HISTORICA") Is Nothing Then
        'Valor default
        fgObterQtdDiasExpurgo = 40
    Else
        fgObterQtdDiasExpurgo = CLng(xmlParmGeral.selectSingleNode("//BASE_HISTORICA").Text)
    End If
    
    Set xmlParmGeral = Nothing
    
    Exit Function
ErrorHandler:
    Set xmlParmGeral = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgObterQtdDiasExpurgo", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Montar header da mensagem a ser enviada para o sistema NZ e retornar o n�mero de controle IF gerado para ela.

Public Function fgMontaHeaderMensageNZ(ByVal pstrCodigoMensagem As String, _
                                       ByVal plngCodigoEmpresa As Long, _
                                       ByRef pstrNumerCrtlIF As String) As String

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

On Error GoTo ErrHandler

    If pstrNumerCrtlIF = vbNullString Then
        pstrNumerCrtlIF = fgObterNumeroControleIF()
    End If
    
    udtHeaderMensagem.ControleRemessaNZ = pstrNumerCrtlIF
    udtHeaderMensagem.CodigoEmpresa = Format$(plngCodigoEmpresa, String(5, "0"))
    udtHeaderMensagem.CodigoMensagem = Trim$(pstrCodigoMensagem)
    udtHeaderMensagem.FormatoMensagem = 2 'enumFormatoMensagem.FormatoMsgXML
    udtHeaderMensagem.DataRemessa = Format$(fgDataHoraServidor(enumFormatoDataHora.Data), "YYYYMMDD")
    udtHeaderMensagem.QuantidadeMensagem = "000001"
    udtHeaderMensagem.SiglaSistemaLegadoOrigem = "A8 "
    udtHeaderMensagem.NuOP = String$(Len(udtHeaderMensagem.NuOP), "0")
    udtHeaderMensagem.CodigoMoeda = "00790"
    udtHeaderMensagem.SiglaSistemaEnviouNZ = "A8 "
    udtHeaderMensagem.AssinaturaInterna = String$(Len(udtHeaderMensagem.AssinaturaInterna), "0")
    udtHeaderMensagem.ReferenciaContabil = Left$("A8" & Right(udtHeaderMensagem.ControleRemessaNZ, 6) & String(8, "0"), 8)
    udtHeaderMensagem.BancoAgencia = String$(Len(udtHeaderMensagem.BancoAgencia), "0")
    udtHeaderMensagem.FILLER = String$(Len(udtHeaderMensagem.FILLER), "0")

    LSet udtHeaderMensagemAux = udtHeaderMensagem
    udtHeaderMensagemAux.String = Replace(udtHeaderMensagemAux.String, vbNullChar, " ")
    
    fgMontaHeaderMensageNZ = udtHeaderMensagemAux.String

    Exit Function
ErrHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7Server", "fgMontaHeaderMensageNZ", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o n�mero de controle IF

Public Function fgObterNumeroControleIF() As String

Dim vntArray()                              As Variant
Dim strNuCtrlIF                             As String

On Error GoTo ErrHandler

    vntArray = Array("A8", strNuCtrlIF)

    fgObterNumeroControleIF = fgExecuteCMD("A8PROC.A8P_SEQUENCIA_NZ", 1, vntArray)

Exit Function

ErrHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7Server", "fgObterNumeroControleIF", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executar comandos de procedures atrav�s de dbLink
'Public Function fgExecuteCMDDBL(ByVal pstrNomeProc As String, _
'                                ByVal pintPosicaoRetorno As Integer, _
'                                ByRef pvntParametros() As Variant) As Variant
'
'Dim objDBLConsulta                          As A6A7A8CA.clsDBLConsulta
'
'On Error GoTo ErrHandler
'
'    Set objDBLConsulta = CreateObject("A6A7A8CA.clsDBLConsulta")
'    fgExecuteCMDDBL = objDBLConsulta.ExecuteCMD(pstrNomeProc, pintPosicaoRetorno, pvntParametros())
'    Set objDBLConsulta = Nothing
'
'    Exit Function
'
'ErrHandler:
'    Set objDBLConsulta = Nothing
'
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, "basA7Server", "fgExecuteCMDDBL", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Function

