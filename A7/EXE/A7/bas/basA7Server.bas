Attribute VB_Name = "basA7Server"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF9D2EF0366"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"
 'Empresa        : Regerbanc
'Componente     : A6A7A8CA
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
    MensagemA8PJRealizado = 1001
    MensagemA8NZ = 1002
    MensagemErroNZA8 = 1003
    MensagemA8PJPrevisto = 1004
    MensagemIDADV = 1005
    MensagemRetornoDV = 1006
    MensagemIDABG = 1007
    MensagemRetornoBG = 1008
    MensagemPZErro = 1009
    MensagemPZR1 = 1010
    MensagemPZR2 = 1011
    MensagemSTR0010R2PZA8 = 1012
    MensagemRetornoPZOk = 1013
    MensagemA8PJMoedaEstrangeira = 1014
    MensagemA8NZR1_HD = 1999
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
    Filler                                  As String * 30
    CodigoErroInterno                       As String * 4
    NomeErroInterno                         As String * 80
    MessageID                               As String * 48
End Type

Public Type udtHeaderDVAux
    String                                  As String * 181
End Type

'----Integração com o Sistema PJ
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
    Filler                                  As String * 46
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
    Filler                                  As String * 20
End Type

Public Type udtMaioresValoresAux
    String                                  As String * 250
End Type

'----

'KIDA
'PJ MOEDA ESTRANGEIRA
Public Type udtMovi_PJ_MoedaEstrangeira
    TipoRemessa                             As String * 3 'A
    CodigoEmpresa                           As String * 5 'N
    SiglaSistema                            As String * 3 'A
    IdentificadorMovimento                  As String * 25 'A
    CodigoMoeda                             As String * 4  'N
    CodigoBanqueiroSwift                    As String * 30 'A Código do Banqueiro no Swift
    CodigoProduto                           As String * 4 'N
    DataMovimento                           As String * 8 'D
    CodigoReferenciaSwift                   As String * 16 'A
    TipoEntradaSaida                        As String * 1 'N - 1 C ; 2 D
    ValorMovimento                          As String * 19 '17,2
    NomeCliente                             As String * 50
    TipoMovimento                           As String * 3 'N - 100 - Previsto;200 - Realizado;300 - Estorno Previsto;400 - Estorno Realizado
    TipoProcessamento                       As String * 1 'N - 1 - On-line;2 - Batch
    ContaBanqueiro                          As String * 35 'A
    Filler                                  As String * 93 'A
End Type

Public Type udtMovi_PJ_MoedaEstrangeiraAux
    String                                  As String * 300
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
                strOcorrencia = "o gerenciador de fila está sendo encerrando(" & strComplemento & ")."
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
            Case "ORA-01034", "ORA-12535", "ORA-12560", "ORA-12541"
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
    
    Err.Clear

End Function

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

'Montar header da mensagem a ser enviada para o sistema NZ e retornar o número de controle IF gerado para ela.

Public Function fgMontaHeaderMensageNZ(ByVal pstrCodigoMensagem As String, _
                                       ByVal plngCodigoEmpresa As Long, _
                                       ByRef pstrNumerCrtlIF As String) As String

Dim udtHeaderMensagem                       As udtProtocoloNZ
Dim udtHeaderMensagemAux                    As udtProtocoloNZAux

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
    udtHeaderMensagem.Filler1 = String$(Len(udtHeaderMensagem.Filler1), "0")

    LSet udtHeaderMensagemAux = udtHeaderMensagem
    udtHeaderMensagemAux.String = Replace(udtHeaderMensagemAux.String, vbNullChar, " ")
    
    fgMontaHeaderMensageNZ = udtHeaderMensagemAux.String

    Exit Function
ErrHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7Server", "fgMontaHeaderMensageNZ", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

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


'Obter o codigo do banco - compensação
Public Function fgObterBanco(ByVal plngCodigoEmpresa As Long) As Long

Dim xmlEmpresa                                As MSXML2.DOMDocument40
Dim xmlInstituicaoISPB                        As MSXML2.DOMDocument40
Dim objEmpresa                                As A6A7A8.clsEmpresa
Dim objInstituicaoISPB                        As Object 'A8LQS.clsInstituicaoISPB

On Error GoTo ErrorHandler

    'Obter o sequencial ISPB
    Set xmlEmpresa = CreateObject("MSXML2.DOMDocument.4.0")
    Set objEmpresa = CreateObject("A6A7A8.clsEmpresa")
    
    If Not xmlEmpresa.loadXML(objEmpresa.Ler(plngCodigoEmpresa)) Then
        'Empresa não cadastrada
        lngCodigoErroNegocio = 3127
        GoTo ErrorHandler
    End If
    
    Set objEmpresa = Nothing

    'Obter o codigo SPB
    Set xmlInstituicaoISPB = CreateObject("MSXML2.DOMDocument.4.0")
    Set objInstituicaoISPB = CreateObject("A8LQS.clsInstituicaoISPB")
    
    If Not xmlInstituicaoISPB.loadXML(objInstituicaoISPB.LerTodos(CLng(xmlEmpresa.documentElement.selectSingleNode("SQ_ISPB").Text))) Then
        'Instituição SPB não cadastrada
        lngCodigoErroNegocio = 3128
        GoTo ErrorHandler
    End If
    
    fgObterBanco = xmlInstituicaoISPB.documentElement.selectSingleNode("//CO_CPEN").Text
    
    Set xmlInstituicaoISPB = Nothing
    Set objInstituicaoISPB = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlEmpresa = Nothing
    Set objEmpresa = Nothing
    Set xmlInstituicaoISPB = Nothing
    Set objInstituicaoISPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7Server", "fgObterBanco Sub", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

