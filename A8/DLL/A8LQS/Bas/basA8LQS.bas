Attribute VB_Name = "basA8LQS"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F4621F400F8"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"

' Este componente tem como objetivo agrupar métodos utilizados na camada de negócios do sistema A8.

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
    FILLER                        As String * 42
    Dominio                       As String * 1
    FILLER_1                      As String * 1
End Type

Public Type udtHeaderMensagemAux
    String                        As String * 200
End Type

'Integração com o Sistema PJ

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
    NumeroOperacaoSelic                     As String * 6
    ContaCedenteCessionaria                 As String * 9
    FILLER                                  As String * 5
End Type

Public Type udtMaioresValoresAux
    String                                  As String * 250
End Type

Public Type udtContabilidade
    EMPRESA                                 As String * 4
    CLAVE_DE_INTERFASE                      As String * 3
    FECHA_CONTABLE                          As String * 8
    FECHA_DE_OPERACION                      As String * 8
    PRODUCTO                                As String * 2
    SUBPRODUCTO                             As String * 4
    GARANTIA                                As String * 3
    TIPO_DE_PLAZO                           As String * 1
    PLAZO                                   As String * 3
    SUBSECTOR                               As String * 1
    SECTOR_B_E                              As String * 2
    CNAE                                    As String * 5
    EMPRESA_TUTELADA                        As String * 4
    AMBITO                                  As String * 2
    MOROSIDAD                               As String * 1
    INVERSION                               As String * 1
    OPERACION                               As String * 3
    CODIGO_CONTABLE                         As String * 5
    DIVISA                                  As String * 3
    TIPO_DE_DIVISA                          As String * 1
    TIPO_NOMINAL                            As String * 5
    Filler1                                 As String * 5
    VARIOS                                  As String * 30
    CLAVE_DE_AUTORIZACION                   As String * 6
    CENTRO_OPERANTE                         As String * 4
    CENTRO_ORIGEN                           As String * 4
    CENTRO_DESTINO                          As String * 4
    NUM_MOVTOS_AL_DEBE                      As String * 7
    NUM_MOVTOS_AL_HABER                     As String * 7
    IMPORTE_DEBE_EN_PESETAS                 As String * 15
    IMPORTE_HABER_EN_PESETAS                As String * 15
    IMPORTE_DEBE_EN_DIVISA                  As String * 15
    IMPORTE_HABER_EN_DIVISA                 As String * 15
    INDICADOR_DE_CORRECCION                 As String * 1
    NUMERO_DE_CONTROL                       As String * 12
    CLAVE_DE_CONCEPTO                       As String * 3
    DESCRIPCION_DE_CONCEPTO                 As String * 14
    TIPODE_CONCEPTO                         As String * 1
    OBSERVACIONES                           As String * 30
    SANCTCCC                                As String * 18
    APLICACION_ORIGEN                       As String * 3
    APLICACION_DESTINO                      As String * 3
    OBSERVACIONES3                          As String * 6
    RESERVAT                                As String * 4
    HACTRGEN                                As String * 4
    HAYCOCAI                                As String * 1
    HAYCTORD                                As String * 1
    SATINTER                                As String * 5
    SACCLVOP                                As String * 3
    SACCEGES                                As String * 4
    SACAPLCP                                As String * 2
    SACCDTGT                                As String * 2
    SAYUTILI                                As String * 1
    SAYROTAC                                As String * 2
    FALTPART                                As String * 8
    OBSERV4                                 As String * 30
    NIO                                     As String * 24
    Filler2                                 As String * 2
End Type

Public Type udtContabilidadeAux
    String                                  As String * 380
End Type

Public Const gstrCodigoMoeda = "0790"

'KIDA
'PJ MOEDA ESTRANGEIRA
Public Type udtMovi_PJ_MoedaEstrangeira
    TipoRemessa                             As String * 3  'A
    CodigoEmpresa                           As String * 5  'N
    SiglaSistema                            As String * 3  'A
    IdentificadorMovimento                  As String * 25 'A
    CodigoMoeda                             As String * 4  'N
    CodigoBanqueiroSwift                    As String * 30 'A Código do Banqueiro no Swift
    CodigoProduto                           As String * 4  'N
    DataMovimento                           As String * 8  'D
    CodigoReferenciaSwift                   As String * 16 'A
    TipoEntradaSaida                        As String * 1  'N - 1 C ; 2 D
    ValorMovimento                          As String * 19 'N - 17,2
    NomeCliente                             As String * 50 'A
    TipoMovimento                           As String * 3  'N - 100 - Previsto;200 - Realizado;300 - Estorno Previsto;400 - Estorno Realizado
    TipoProcessamento                       As String * 1  'N - 1 - On-line;2 - Batch
    ContaBanqueiro                          As String * 35 'A
    FILLER                                  As String * 93 'A
End Type

Public Type udtMovi_PJ_MoedaEstrangeiraAux
    String                                  As String * 300
End Type

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

' Obtém tipo de backoffice do usuário.
Public Function fgObterTipoBackOfficeUsuario() As Integer

Dim objUsuario                           As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsControleAcesso")
    
    fgObterTipoBackOfficeUsuario = objUsuario.ObterTipoBackOfficeUsuario(fgUsuarioRede)
    
    Set objUsuario = Nothing
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgObterTipoBackOfficeUsuario Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Obtém tipo de backoffice do veículo legal
Public Function fgObterTipoBackOffice(ByVal pstrCodigoVeiculoLegal As String, _
                             Optional ByVal pstrSiglaSistema As String = vbNullString) As Long

Dim objVeiculoLegal                           As A6A7A8.clsVeiculoLegal

On Error GoTo ErrorHandler

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")

    fgObterTipoBackOffice = objVeiculoLegal.ObterTipoBackOffice(pstrCodigoVeiculoLegal, pstrSiglaSistema)

    Set objVeiculoLegal = Nothing
    Exit Function
ErrorHandler:

    Set objVeiculoLegal = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgObterTipoBackOffice Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o código do veículo legal
Public Function fgObterCodigoVeiculoLegal(ByVal pstrCodigoMensagem As String, _
                                          ByVal pvntCNPJ_IDENT_PART As Variant, _
                                          ByVal plngCodigoEmpresa As Long, _
                                          ByRef pstrCodigoVeiculoLegal As String, _
                                          ByRef pstrSiglaSistema As String, _
                                          ByRef plngTipoBackOffice As Long, _
                                 Optional ByRef pstrNomeVeicLega As String, _
                                 Optional ByRef pvntIDENT_PART_CAMR As Variant = 0, _
                                 Optional ByVal pblnCliente1 As Boolean) As Boolean

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal

On Error GoTo ErrorHandler

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")

    fgObterCodigoVeiculoLegal = objVeiculoLegal.ObterCodigoVeiculoLegal(pstrCodigoMensagem, _
                                                                        pvntCNPJ_IDENT_PART, _
                                                                        plngCodigoEmpresa, _
                                                                        pstrCodigoVeiculoLegal, _
                                                                        pstrSiglaSistema, _
                                                                        plngTipoBackOffice, _
                                                                        pstrNomeVeicLega, _
                                                                        pvntIDENT_PART_CAMR, _
                                                                        pblnCliente1)

    Set objVeiculoLegal = Nothing
    Exit Function
ErrorHandler:

    Set objVeiculoLegal = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgObterCodigoVeiculoLegal Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Executa uma instrução SQL na base de dados Oracle.
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
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Esta função deve ser utilizada para toda Procedure
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
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Retorna propriedades de uma tabela.
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
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgPropriedades", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Executa um select na base de dados Oracle.
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

' Executa um select na base de dados Oracle e retorna uma string em formato XML.
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

' Executa um select na base de dados Oracle e retorna uma string em formato XML.
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

' Executa um select na base de dados Oracle e retorna uma string em formato XML.
Public Function fgQueryXMLLerTodosCACHE(ByVal pstrNomeXML As String, _
                                        ByVal pstrSQL As String, _
                                        ByVal pstrNomeObjeto As String, _
                                        ByVal pstrNomeCache As String) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgQueryXMLLerTodosCACHE = objConsulta.QueryXMLLerTodosCache(pstrNomeXML, pstrSQL, pstrNomeObjeto, pstrNomeCache)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQueryXMLLerTodosCACHE", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Formata a data/hora para o banco de dados oracle
Public Function fgFormataDataHoraOracle(ByVal pdtmData As Date) As String
    
    fgFormataDataHoraOracle = "TO_DATE('" & Format(pdtmData, "DD/MM/YYYY HH:mm:ss") & "','dd/mm/yyyy HH24:mi:ss')"

End Function

' Executa uma sequence na base de dados Oracle.
Public Function fgExecuteSequence(ByVal pstrNomeSequence As String) As Variant

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteSequence = objTransacao.ExecuteSequence(pstrNomeSequence)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgExecuteSequence", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o erro
Public Function fgGetError(ByVal psErro As String) As String

Dim objDomError                             As MSXML2.DOMDocument40
Dim intErrPosicaoInicio                     As Integer
Dim intErrPosicaoFim                        As Integer
Dim strErrSource                            As String
Dim strTxtPesquisa                          As String
Dim lngErrNumber                            As Long
    
On Error GoTo ErrorHandler

    Set objDomError = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDomError.loadXML(psErro) Then
        fgErroLoadXML objDomError, App.EXEName, "basA8LQS", "fgGetError"
    End If

    strErrSource = objDomError.selectSingleNode("Erro/Grupo_ErrorInfo/Souce").Text
    
    
    If strErrSource = "mqax200" Then
        strTxtPesquisa = "ReasonCode = "
        intErrPosicaoInicio = InStr(1, psErro, strTxtPesquisa)
        
        If intErrPosicaoInicio > 0 Then
            intErrPosicaoInicio = intErrPosicaoInicio + Len(strTxtPesquisa)
            intErrPosicaoFim = InStr(intErrPosicaoInicio, psErro, ",")
            lngErrNumber = Mid(psErro, intErrPosicaoInicio, intErrPosicaoFim - intErrPosicaoInicio)
            fgGetError = "MQ-" & Format(lngErrNumber, "00000")
        End If
        
        
    ElseIf strErrSource = "Microsoft OLE DB Provider for Oracle" Then
        
        
    Else
    
    End If
        
    Set objDomError = Nothing
    
    Exit Function

ErrorHandler:
       
    Set objDomError = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Função genérica para tratamento de erros.
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
'Converter o valor para a formatação do oracle
Public Function fgVlrToDBServer(ByVal pvntNumero As Variant, _
                       Optional ByVal pbDecimalZonado As Boolean = False, _
                       Optional ByVal piQtdeDecimais As Integer = 2) As Variant

'Recebe como parâmetro um valor com o separador decimal em qualquer formato.
'Retorna um valor com o separador decimal no formato do Banco de Dados (Ponto)

Dim strSeparador                As String
Dim strSeparadorTroca           As String
Dim intPosSeparador             As Integer
Dim intPosSeparadorTroca        As Integer

    On Error GoTo ErrorHandler
    
    If Trim(pvntNumero) = vbNullString Then
        fgVlrToDBServer = 0
        Exit Function
    End If
    
    If Int(CDbl("1.1")) = 1 Then
        strSeparador = "."
        strSeparadorTroca = ","
    Else
        strSeparador = ","
        strSeparadorTroca = "."
    End If
    
    intPosSeparador = InStr(pvntNumero, strSeparador)
    intPosSeparadorTroca = InStr(pvntNumero, strSeparadorTroca)
    
    If intPosSeparador > 0 And intPosSeparadorTroca > 0 Then
        pvntNumero = Replace(pvntNumero, _
                IIf(intPosSeparador > intPosSeparadorTroca, strSeparadorTroca, strSeparador), vbNullString)
    End If
    
    fgVlrToDBServer = Replace(pvntNumero, strSeparadorTroca, strSeparador)
    fgVlrToDBServer = Replace(fgVlrToDBServer, ",", ".")
    
    Exit Function

ErrorHandler:
    lngCodigoErroNegocio = 43
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgVlrToDBServer|" & pvntNumero, lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Converter o valor para a formatação do xml
Public Function fgVlrToXML(ByVal pvntNumero As Variant) As Variant

'Recebe como parâmetro um valor com o separador decimal em qualquer formato.
'Retorna um valor com o separador decimal no formato dos XML´s convencionados no LQS (Virgula)

Dim strSeparador                As String
Dim strSeparadorTroca           As String
Dim intPosSeparador             As Integer
Dim intPosSeparadorTroca        As Integer

    On Error GoTo ErrorHandler
    
    If Trim(pvntNumero) = vbNullString Then
        fgVlrToXML = 0
        Exit Function
    End If
    
    If Int(CDbl("1.1")) = 1 Then
        strSeparador = "."
        strSeparadorTroca = ","
    Else
        strSeparador = ","
        strSeparadorTroca = "."
    End If
    
    intPosSeparador = InStr(pvntNumero, strSeparador)
    intPosSeparadorTroca = InStr(pvntNumero, strSeparadorTroca)
    
    If intPosSeparador > 0 And intPosSeparadorTroca > 0 Then
        pvntNumero = Replace(pvntNumero, _
                IIf(intPosSeparador > intPosSeparadorTroca, strSeparadorTroca, strSeparador), vbNullString)
    End If
    
    fgVlrToXML = Replace(pvntNumero, strSeparadorTroca, strSeparador)
    fgVlrToXML = Replace(fgVlrToXML, ".", ",")
    
    Exit Function

ErrorHandler:
    lngCodigoErroNegocio = 43
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgVlrToXML|" & pvntNumero, lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o Usuário da Rede que está acessando a rotina
Public Function fgUsuarioRede() As String

Dim objUsuario                              As A6A7A8.clsUsuario

    On Error GoTo ErrorHandler

    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    fgUsuarioRede = objUsuario.ObterUsuarioRede
    Set objUsuario = Nothing
    
    Exit Function

ErrorHandler:
    Set objUsuario = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgUsuarioRede", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter a estação de trabalho do Usuário da Rede que está acessando a rotina
Public Function fgEstacaoTrabalhoUsuario() As String

Dim objUsuario                              As A6A7A8.clsUsuario
Dim strUsuarioRede                          As String
Dim strEstacaoTrabalho                      As String

    On Error GoTo ErrorHandler
    
    strUsuarioRede = fgUsuarioRede
    lngCodigoErroNegocio = 0

    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    Call objUsuario.ObterEstacaoTrabalhoUsuario(strUsuarioRede, _
                                                strEstacaoTrabalho, _
                                                lngCodigoErroNegocio, _
                                                intNumeroSequencialErro)
    Set objUsuario = Nothing

    If lngCodigoErroNegocio <> 0 Then
        strEstacaoTrabalho = "SERVIDOR"
    End If

    fgEstacaoTrabalhoUsuario = strEstacaoTrabalho
    
    Exit Function

ErrorHandler:
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgEstacaoTrabalhoUsuario", lngCodigoErroNegocio, intNumeroSequencialErro, "Usuário: " & strUsuarioRede)
    
End Function

' Trata da segregação de dados a usuários.
Public Function fgSegregaDados(ByVal strOwnerTabela As String, _
                      Optional ByVal blnRetornaSubSelect As Boolean = True, _
                      Optional ByVal strAliasTabelaFato As String = "TabFato_Sub", _
                      Optional ByVal strAliasTabelaDominio As String = "TabDominio_Sub", _
                      Optional ByVal blnPrefixoWhere As Boolean = True, _
                      Optional ByVal blnEstabeleceJoinComTabelaDominio As Boolean = True, _
                      Optional ByVal blnFiltraTipoBackOffice As Boolean = True, _
                      Optional ByVal blnFiltraLocalLiquidacao As Boolean = True, _
                      Optional ByVal blnFiltraGrupoVeiculoLegal As Boolean = True, _
                      Optional ByVal blnFiltraGrupoUsuario As Boolean = True, _
                      Optional ByVal xmlDocComplementoWhere As MSXML2.DOMDocument40, _
                      Optional ByVal blnConsideraNulos As Boolean = False) As String

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")

    fgSegregaDados = objControleAcesso.SegregaDados(strOwnerTabela, _
                                                    blnRetornaSubSelect, _
                                                    strAliasTabelaFato, _
                                                    strAliasTabelaDominio, _
                                                    blnPrefixoWhere, _
                                                    blnEstabeleceJoinComTabelaDominio, _
                                                    blnFiltraTipoBackOffice, _
                                                    blnFiltraLocalLiquidacao, _
                                                    blnFiltraGrupoVeiculoLegal, _
                                                    blnFiltraGrupoUsuario, _
                                                    xmlDocComplementoWhere, _
                                                    blnConsideraNulos)

    Set objControleAcesso = Nothing

    Exit Function

ErrorHandler:
    Set objControleAcesso = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgSegregaDados Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar a remessa rejeitada para o sistema legado
Public Function fgRemessaRejeitadaLegado(ByRef objDomRemessa As MSXML2.DOMDocument40, _
                                         ByVal objDomErro As MSXML2.DOMDocument40, _
                                Optional ByVal pblnFinaliza As Boolean = False) As Boolean

Dim objLegado                               As A8LQS.clsLegado
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.MontarMensagemRejeicao(objDomRemessa, objDomErro)
    Set objLegado = Nothing

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, pblnFinaliza
    Set objProcessoOperacao = Nothing

    fgRemessaRejeitadaLegado = True

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgRemessaRejeitadaLegado", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Trata inclusão de campos declarados como varchar de 4000 na base de dados Oracle.
Public Function fgInsertVarchar4000(ByVal pstrConteudoCampoVarchar As String, _
                           Optional ByVal pstrCodigoTextoXML As String = vbNullString, _
                           Optional ByVal pblnConverterBase64 As Boolean = True) As Long

'RETORNA O NUMERO SEQUENCIAL

Dim objTransacao                            As A6A7A8CA.clsTransacao

    On Error GoTo ErrorHandler
         
    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    
    fgInsertVarchar4000 = objTransacao.InsertVarchar4000("A8.TB_TEXT_XML", _
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
    Call fgRaiseError(App.EXEName, "basA8LQS", "InsertVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Trata consulta de campos declarados como varchar de 4000 na base de dados Oracle.
Public Function fgSelectVarchar4000(ByVal pvntSequencial As Variant, _
                           Optional ByVal pblnConverterBase64 As Boolean = True) As String

'Retorna decode base64

Dim objConsulta                             As A6A7A8CA.clsConsulta
Dim strTabelaTexto                          As String

On Error GoTo ErrorHandler
    
    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    
    '-----------------------------------------------------------------------------------
    'Pikachu - 22/06/2004'
    'OBS : O registro com CO_TEXT_XML = 0 (zero) é destinado para os parametros gerais
    '-----------------------------------------------------------------------------------
    
    If pvntSequencial >= 0 Then
        strTabelaTexto = "A8.TB_TEXT_XML"
    Else
        strTabelaTexto = "A8HIST.TB_TEXT_XML"
    End If
    
    fgSelectVarchar4000 = objConsulta.SelectVarchar4000(strTabelaTexto, _
                                                        "CO_TEXT_XML", _
                                                        Abs(pvntSequencial), _
                                                        "TX_XML", _
                                                        "NU_SEQU_TEXT_XML", _
                                                        pblnConverterBase64)
    
    Set objConsulta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgSelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar o Header para o sistema NZ
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
    udtHeaderMensagem.FILLER_1 = String$(Len(udtHeaderMensagem.FILLER_1), "0")
    
    If Mid$(Trim$(pstrCodigoMensagem), 1, 3) = "CCR" Or (Mid$(Trim$(pstrCodigoMensagem), 1, 3) = "CAM" _
                                                         And Trim$(pstrCodigoMensagem) <> "CAM0001" _
                                                         And Trim$(pstrCodigoMensagem) <> "CAM0002" _
                                                         And Trim$(pstrCodigoMensagem) <> "CAM0003" _
                                                         And Trim$(pstrCodigoMensagem) <> "CAM0004") Then
        udtHeaderMensagem.Dominio = 2
    Else
        udtHeaderMensagem.Dominio = 0
    End If

    LSet udtHeaderMensagemAux = udtHeaderMensagem
    udtHeaderMensagemAux.String = Replace(udtHeaderMensagemAux.String, vbNullChar, " ")
    
    fgMontaHeaderMensageNZ = udtHeaderMensagemAux.String

    Exit Function
ErrHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgMontaHeaderMensageNZ", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


'Montar o Header para o sistema PZ - TED
Public Function fgMontarHeaderPZTED(ByVal xmlOperacao As MSXML2.DOMDocument40) As String

Dim udtPZW0001                              As udtPZW0001
Dim udtPZW0001Aux                           As udtPZW0001Aux
Dim vntValor                                As Variant
Dim lngBanco                                As Long

Dim objProcessoMesgSTR                      As A8LQS.clsProcessoMensagemSTR
Dim vntPontoVendaUniorg                     As Variant
Dim vntConta                                As Variant
Dim lngAgencia                              As Long

On Error GoTo ErrHandler
            
    If xmlOperacao.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
        Call fgAppendNode(xmlOperacao, xmlOperacao.documentElement.nodeName, "NU_CTRL_IF", fgObterNumeroControleIF())
    End If
    
    If xmlOperacao.selectSingleNode("//NU_DOCT") Is Nothing Then
        Call fgAppendNode(xmlOperacao, xmlOperacao.documentElement.nodeName, "NU_DOCT", Format$(Now, "HHMMSS"))
    End If
    
    If xmlOperacao.selectSingleNode("//CO_ISPB_IF_DEBT") Is Nothing Then
        Call fgAppendNode(xmlOperacao, xmlOperacao.documentElement.nodeName, "CO_ISPB_IF_DEBT", fgObterISPBIF(xmlOperacao.selectSingleNode("//CO_EMPR").Text))
    End If
    
    With udtPZW0001
        .TipoRegistro = "2"
        .BancoOrigem = "033"
        .UnidadeOrigem = fgCompletaString(xmlOperacao.selectSingleNode("//CO_UNI_ORG").Text, "0", Len(.UnidadeOrigem), True)
        .SiglaSistemaOrigem = fgCompletaString$("A8", " ", 3, False)
        .CodigoSistemaOrigem = "0001"
        .ControleLegado = fgCompletaString$(xmlOperacao.selectSingleNode("//NU_CTRL_IF").Text, " ", 23, False)
        .NumeroDocumento = fgCompletaString$(xmlOperacao.selectSingleNode("//NU_DOCT").Text, "0", 6, True)
        .BancoDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_BANC_CRED").Text, "0", 3, True)
        .ISPBDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_ISPB_IF_CRED").Text, " ", 8, False)
        .DataContabil = xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text
        .HoraAgendamento = String$(6, "0")
        .CodigoAcao = "E"               ' Cod. Acatamento (E)fetivar (C)onsulta
        .MeioTransferencia = IIf(Left$(xmlOperacao.selectSingleNode("//CO_MESG").Text, 3) = "STR", "1", "2")
        .IdentificadorAlteracao = "0"   ' (0) A8 controla CC - (1) PZ controla CC
        .NumeroVersao = "01"
        .LancamentoCC = "0"
        .HistoricoCC = String$(5, "0")
        .ContaDebitada = fgCompletaString$(vntConta, "0", 13, True)
        .AgeciaDebitada = fgCompletaString$(lngAgencia, "0", 5, True)
        .OrigemContabilidade = "87"
        .Filler1 = String$(39, " ")
        .Erro1 = String$(5, "0")
        .Erro2 = String$(5, "0")
        .Erro3 = String$(5, "0")
        .CodigoPZ = String$(50, " ")
        .CodigoMensagem = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_MESG").Text, " ", 9, False)
        
        vntValor = Split(xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
        
        If UBound(vntValor) = 0 Then
            .ValorLancamento = fgCompletaString$(vntValor(0), "0", 16, True) & "00"
        Else
            .ValorLancamento = fgCompletaString$(vntValor(0), "0", 16, True) & fgCompletaString$(vntValor(1), "0", 2, False)
        End If
        
        .DataMovimento = xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text
        
        If Not xmlOperacao.selectSingleNode("//CO_AGEN_CRED") Is Nothing Then
            .AgenciaCreditada = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_AGEN_CRED").Text, "0", 5, True)
        Else
            .AgenciaCreditada = String$(5, "0")
        End If

        If Not xmlOperacao.selectSingleNode("//CO_AGEN_DEBT") Is Nothing Then
            .AgenciaRemetente = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_AGEN_DEBT").Text, "0", 5, True)
        Else
            .AgenciaRemetente = String$(5, "0")
        End If
        
        .Filler2 = Space(40)
    End With

    LSet udtPZW0001Aux = udtPZW0001
    udtPZW0001Aux.String = Replace(udtPZW0001Aux.String, vbNullChar, " ")

    fgMontarHeaderPZTED = udtPZW0001Aux.String

    Exit Function
ErrHandler:
    
    Set objProcessoMesgSTR = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgMontarHeaderPZTED", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Montar o Header para o sistema PZ
Public Function fgMontarHeaderPZ(ByVal xmlOperacao As MSXML2.DOMDocument40) As String

Dim udtPZW0001                              As udtPZW0001
Dim udtPZW0001Aux                           As udtPZW0001Aux
Dim vntValor                                As Variant
Dim lngBanco                                As Long

Dim objProcessoMesgSTR                      As A8LQS.clsProcessoMensagemSTR
Dim vntPontoVendaUniorg                     As Variant
Dim vntConta                                As Variant
Dim lngAgencia                              As Long

On Error GoTo ErrHandler
            
    'Obter ponto de venda / uniorg
    Set objProcessoMesgSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
    
    Call objProcessoMesgSTR.ObterAgenciaContaDebitada(CLng(xmlOperacao.selectSingleNode("//CO_EMPR").Text), _
                                                      CLng(xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text), _
                                                      lngAgencia, _
                                                      vntConta, _
                                                      vntPontoVendaUniorg)
    Set objProcessoMesgSTR = Nothing
    
'    lngBanco = fgObterBanco(xmlOperacao.selectSingleNode("//CO_EMPR").Text)
    lngBanco = 33
    
    With udtPZW0001
        .TipoRegistro = "2"
        .BancoOrigem = fgCompletaString$(lngBanco, "0", 3, True)
        .UnidadeOrigem = fgCompletaString(vntPontoVendaUniorg, "0", Len(.UnidadeOrigem), True)
        .SiglaSistemaOrigem = fgCompletaString$("A8", " ", 3, False)
        .CodigoSistemaOrigem = "0001"
        .ControleLegado = fgCompletaString$(xmlOperacao.selectSingleNode("//NU_CTRL_IF").Text, " ", 23, False)
        .NumeroDocumento = fgCompletaString$(xmlOperacao.selectSingleNode("//NU_DOCT").Text, "0", 6, True)
        .BancoDestino = String$(3, "0")
        
        .ISPBDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_ISPB_IF_DEBT").Text, "0", 8, True)
        
        If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
            If Val(xmlOperacao.selectSingleNode("//TP_MESG").Text) <> enumTipoMensagemLQS.DespesasBMC Then
                If Not xmlOperacao.selectSingleNode("//CO_MESG") Is Nothing Then
                    If xmlOperacao.selectSingleNode("//CO_MESG").Text = "STR0007" Then
                        .ISPBDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_ISPB_IF_CRED").Text, "0", 8, True)
                    End If
                End If
            Else
                If Not xmlOperacao.selectSingleNode("//CO_MESG") Is Nothing Then
                    If xmlOperacao.selectSingleNode("//CO_MESG").Text = "STR0007" Then
                        If Not xmlOperacao.selectSingleNode("//CO_BANC") Is Nothing Then
                            .BancoDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_BANC").Text, "0", 3, True)
                        End If
                    End If
                End If
            End If
        ElseIf Not xmlOperacao.selectSingleNode("//CORRETORAS") Is Nothing Then
            If Not xmlOperacao.selectSingleNode("//CO_MESG") Is Nothing Then
                If xmlOperacao.selectSingleNode("//CO_MESG").Text = "STR0007" Then
                    .ISPBDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_ISPB_IF_CRED").Text, "0", 8, True)
                End If
            End If
        ElseIf Not xmlOperacao.selectSingleNode("//CO_MESG") Is Nothing Then
            If xmlOperacao.selectSingleNode("//CO_MESG").Text = "STR0004" Then
                .ISPBDestino = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_ISPB_IF_CRED").Text, "0", 8, True)
            End If
        End If
        
        .DataContabil = xmlOperacao.selectSingleNode("//DT_SIST").Text
        .HoraAgendamento = String$(6, "0")
        .CodigoAcao = "E"               ' Cod. Acatamento (E)fetivar (C)onsulta
        .MeioTransferencia = "1"
        .IdentificadorAlteracao = "0"   ' (0) A8 controla CC - (1) PZ controla CC
        .NumeroVersao = "01"
        .LancamentoCC = "0"
        .HistoricoCC = String$(5, "0")
        .ContaDebitada = fgCompletaString$(vntConta, "0", 13, True)
        .AgeciaDebitada = fgCompletaString$(lngAgencia, "0", 5, True)
        .OrigemContabilidade = "87"
        .Filler1 = String$(39, " ")
        .Erro1 = String$(5, "0")
        .Erro2 = String$(5, "0")
        .Erro3 = String$(5, "0")
        .CodigoPZ = String$(50, " ")
        .CodigoMensagem = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_MESG").Text, " ", 9, False)
        
        vntValor = Split(xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
        
        If UBound(vntValor) = 0 Then
            .ValorLancamento = fgCompletaString$(vntValor(0), "0", 16, True) & "00"
        Else
            .ValorLancamento = fgCompletaString$(vntValor(0), "0", 16, True) & fgCompletaString$(vntValor(1), "0", 2, False)
        End If
        
        .DataMovimento = xmlOperacao.selectSingleNode("//DT_SIST").Text
        
        If Not xmlOperacao.selectSingleNode("//CO_AGEN_CRED") Is Nothing Then
            .AgenciaCreditada = fgCompletaString$(xmlOperacao.selectSingleNode("//CO_AGEN_CRED").Text, "0", 5, True)
        Else
            .AgenciaCreditada = String$(5, "0")
        End If

        .AgenciaRemetente = String$(5, "0")
        
        .Filler2 = Space(40)
    End With

    LSet udtPZW0001Aux = udtPZW0001
    udtPZW0001Aux.String = Replace(udtPZW0001Aux.String, vbNullChar, " ")

    fgMontarHeaderPZ = udtPZW0001Aux.String

    Exit Function
ErrHandler:
    
    Set objProcessoMesgSTR = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgMontarHeaderPZ", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function




'Obter o próximo múmero de controle IF
Public Function fgObterNumeroControleIF() As String

Dim vntArray()                              As Variant
Dim strNuCtrlIF                             As String

On Error GoTo ErrHandler

    vntArray = Array("A8", strNuCtrlIF)

    fgObterNumeroControleIF = fgExecuteCMD("A8PROC.A8P_SEQUENCIA_NZ", 1, vntArray)

Exit Function

ErrHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterNumeroControleIF", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o próximo identificador de remessa para o sistema PJ
Public Function fgObterIdentificadorRemessaPJ(ByVal plngCodigoEmpresa As Long, _
                                              ByVal pdatDataMovimento As Date, _
                                              ByVal plngQuantMensagem As Long, _
                                              ByRef pstrCodigoInicial As String, _
                                              ByRef pstrCodigoFinal As String) As Long

Dim vntRetorno                              As Variant

On Error GoTo ErrHandler

    vntRetorno = fgExecuteSequence("A8.SQ_A8_NU_SEQU_REME_PJ")
    pstrCodigoInicial = fgDt_To_Xml(pdatDataMovimento) & _
                        "A8 1" & _
                        fgCompletaString(vntRetorno, "0", 8, True)

    pstrCodigoFinal = pstrCodigoInicial

Exit Function

ErrHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterIdentificadorRemessaPJ", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Função do ValidarRemessa
Public Function fgValidarRemessa(ByRef pxmlremessa As MSXML2.DOMDocument40, _
                                 ByRef pstrErros As String) As Boolean

Dim objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa

On Error GoTo ErrorHandler
    
    If Not pxmlremessa.selectSingleNode("//IN_VALIDA_REME") Is Nothing Then
        pstrErros = vbNullString
        fgValidarRemessa = True
        Exit Function
    End If
    
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
    fgValidarRemessa = False
    pstrErros = objValidaRemessa.ValidarMensagemA8(pxmlremessa.xml)
    If pstrErros = vbNullString Then
        fgValidarRemessa = True
    End If

    Set objValidaRemessa = Nothing

Exit Function
ErrorHandler:

    fgValidarRemessa = False
    
    pstrErros = ""
    
    Set objValidaRemessa = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgValidarRemessa Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Verificar se o regsitro é de Cliente 1
Public Function fgVerificarCliente1(ByRef pxmlOperacao As MSXML2.DOMDocument40, _
                           Optional ByVal pblnLiberacao As Boolean = False, _
                           Optional ByVal pblnVerificarProduto As Boolean = True) As Boolean

Dim objTipoConta                            As A8LQS.clsTipoConta
Dim strContaVeiculoLegal                    As String
Dim strContaContraparte                     As String
Dim blnCliente1                             As Boolean

On Error GoTo ErrorHandler

    If pxmlOperacao.documentElement.selectSingleNode("CLIENTE1") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "CLIENTE1", enumIndicadorSimNao.Nao
    Else
        pxmlOperacao.documentElement.selectSingleNode("CLIENTE1").Text = enumIndicadorSimNao.Nao
    End If

    If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
        fgVerificarCliente1 = False
        Exit Function
    End If

    If CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.Definitiva And _
        CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaIda And _
        CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaVolta And _
        CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaVoltaConciliacao And _
        CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.Termo And _
        CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.TermoDataLiquidacaoCerta And _
        CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta Then
        fgVerificarCliente1 = False
        Exit Function
    End If

    If pblnLiberacao = False Then
        If pblnVerificarProduto Then
            If pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> "510" And _
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> "511" And _
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> "512" And _
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> "513" And _
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> "514" And _
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> "515" Then
                fgVerificarCliente1 = False
                Exit Function
            End If
        End If
    End If

    Set objTipoConta = CreateObject("A8LQS.clsTipoConta")

    If pblnLiberacao Then
        blnCliente1 = objTipoConta.VerificarCliente1(pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text)
    Else
        blnCliente1 = objTipoConta.VerificarCliente1(pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text)
    End If

    If pblnLiberacao = False Then
        strContaVeiculoLegal = fgCompletaString(pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text, "0", 9, True)
        strContaContraparte = fgCompletaString(pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text, "0", 9, True)
 
        If Mid(strContaVeiculoLegal, 1, 4) <> Mid(strContaContraparte, 1, 4) Then
            fgVerificarCliente1 = False
            Exit Function
        End If
    End If

    If blnCliente1 Then
        pxmlOperacao.documentElement.selectSingleNode("CLIENTE1").Text = enumIndicadorSimNao.Sim
        fgVerificarCliente1 = True
    Else
        fgVerificarCliente1 = False
    End If

    'If fgVerificarCliente1 Then
    '    fgDebitoCreditoCliente1 pxmlOperacao
    'End If

    Set objTipoConta = Nothing

Exit Function
ErrorHandler:

    fgVerificarCliente1 = False
    Set objTipoConta = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgVerificarCliente1 Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Para uso no empilhamento de erros para Remessa Rejeitada
Public Sub fgAdicionaErro(ByRef xmlDOMErro As MSXML2.DOMDocument40, _
                          ByVal lngCodigoErroNegocio As Long, _
                          Optional ByVal lngCodigoJustificativa As Long = 0)


Dim objLogErros                             As A6A7A8CA.clsLogErro

On Error GoTo ErrorHandler

    Set objLogErros = CreateObject("A6A7A8CA.clsLogErro")
    objLogErros.AdicionaErroNegocio xmlDOMErro, lngCodigoErroNegocio, lngCodigoJustificativa
    
    Set objLogErros = Nothing

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgAdicionaErro", 0
End Sub

'Converter o Débito e Crédito para as operações de Cliente 1
Public Function fgDebitoCreditoCliente1(ByRef pxmlOperacao As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler

    Select Case pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text
        Case "499"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            End If
        Case "500"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            End If
        Case "501"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            End If
        Case "502"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            End If
        Case "503"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            End If
        Case "514"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            End If
        Case "510"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            End If
        Case "511"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            End If
        Case "512"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            End If
        Case "513"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            End If
        Case "504"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            End If
        Case "515"
            If CLng(pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.ENTRADA
            Else
                pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            End If
    End Select

Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgDebitoCreditoCliente1 Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o ISPB da Instituição Financeira
Public Function fgObterISPBIF(ByVal plngCodigoEmpresa As Long) As String


On Error GoTo ErrorHandler
    
    Select Case plngCodigoEmpresa
        Case enumCodigoEmpresa.Banespa
            fgObterISPBIF = fgCompletaString(enumISPB.IspbBANESPA, "0", 8, True)
        Case enumCodigoEmpresa.Bozano
            fgObterISPBIF = fgCompletaString(enumISPB.IspbBOZZANO, "0", 8, True)
        Case enumCodigoEmpresa.Meridional
            fgObterISPBIF = fgCompletaString(enumISPB.IspbMERIDIONAL, "0", 8, True)
        Case enumCodigoEmpresa.Santander
            fgObterISPBIF = fgCompletaString(enumISPB.IspbSANTANDER, "0", 8, True)
        Case enumCodigoEmpresa.REAL_ABN
            fgObterISPBIF = fgCompletaString(enumISPB.IspbREAL_ABN, "0", 8, True)
        Case enumCodigoEmpresa.SUDAMERIS
            fgObterISPBIF = fgCompletaString(enumISPB.IspbSUDAMERIS, "0", 8, True)
        Case enumCodigoEmpresa.BANDEPE
            fgObterISPBIF = fgCompletaString(enumISPB.IspbBANDEPE, "0", 8, True)
    End Select
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterISPBIF Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Obter a descrição do Erro de negócio
Public Function fgObterDescricaoErro(ByVal pstrCodigoErro As String) As String

Dim objLogErro                           As A6A7A8CA.clsLogErro

On Error GoTo ErrorHandler
    
    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")
    fgObterDescricaoErro = objLogErro.ObterDescErroNegocio(pstrCodigoErro)
    Set objLogErro = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objLogErro = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgObterDescricaoErro Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verificar se a da de operação e a data de vencimento da operação compromissada são iguais
Public Function fgDatasIguaisCompromissada(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler

    If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = _
        xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text Then
        fgDatasIguaisCompromissada = True
    Else
        fgDatasIguaisCompromissada = False
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgObterDescricaoErro Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verifica se é um micro de desenvolvimento (pelo padrao de nomes adotado na Regerbanc)
 Public Function fgEstacaoDesenvolvimento() As Boolean

Dim strAux                                  As String

    strAux = LCase(fgEstacaoTrabalhoUsuario)
    fgEstacaoDesenvolvimento = strAux = "d1272131" Or strAux = "note-ivan" Or strAux = "note7-bruno" Or strAux = "note7-mfreitas" Or strAux = "note-jusiel" Or strAux = "note-betovega" Or strAux = "note-buzato"
'    fgEstacaoDesenvolvimento = False

End Function

'Obter o CNPJ da empresa
Public Function fgObterCNPJEmpresa(ByVal plngCodigoEmpresa As Long) As String

Dim xmlEmpresa                                As MSXML2.DOMDocument40
Dim objEmpresa                                As A6A7A8.clsEmpresa

On Error GoTo ErrorHandler

    'Obter o sequencial ISPB
    Set xmlEmpresa = CreateObject("MSXML2.DOMDocument.4.0")
    Set objEmpresa = CreateObject("A6A7A8.clsEmpresa")
    
    If xmlEmpresa.loadXML(objEmpresa.Ler(plngCodigoEmpresa)) Then
        fgObterCNPJEmpresa = xmlEmpresa.selectSingleNode("//NU_CNPJ").Text
    Else
        'Instituição SPB não cadastrada
        lngCodigoErroNegocio = 3127
        GoTo ErrorHandler
    End If
        
    Set xmlEmpresa = Nothing
    Set objEmpresa = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlEmpresa = Nothing
    Set objEmpresa = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterCNPJEmpresa Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Converter o produto para quando a operação for de Cliente 1
Public Function fgConverterProduto1(ByRef pxmlOperacao As MSXML2.DOMDocument40, _
                           Optional ByVal pblnCompromissadaIda As Boolean = False, _
                           Optional ByVal pblnSegundaCompromissadaIda As Boolean = False, _
                           Optional ByVal pblnAlterarEntradaSaida As Boolean = True) As Boolean

On Error GoTo ErrorHandler

    Select Case pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text
        Case "513"
            If pblnCompromissadaIda Then
                If pblnSegundaCompromissadaIda Then
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "500"
                Else
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "510"
                End If
            Else
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "501"
            End If
        Case "500", "510", "501"
            If pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.CompromissadaVolta And _
                pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "513"
                Exit Function
            Else
                If pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "510" Then
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "500"
                Else
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "513"
                    Exit Function
                End If
            End If
        Case "512"
            If pblnCompromissadaIda Then
                If pblnSegundaCompromissadaIda Then
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "499"
                Else
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "511"
                End If
            Else
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "502"
            End If
        Case "499", "511", "502"
            If pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.CompromissadaVolta And _
                pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
                pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "512"
                Exit Function
            Else
                If pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "511" Then
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "499"
                Else
                    pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "512"
                    Exit Function
                End If
            End If
        Case "514"
            pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "503"
        Case "503"
            pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "514"
        Case "515"
            pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "504"
        Case "504"
            pxmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "515"
    End Select

    If pblnAlterarEntradaSaida Then fgDebitoCreditoCliente1 pxmlOperacao

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgConverterProduto1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Criar a tag para operações de transferência
Public Function fgCriarTAGTransferencia(ByRef pxmlOperacao As MSXML2.DOMDocument40, _
                                        ByVal penumProdutoCamara As enumIndicadorSimNao) As Boolean

On Error GoTo ErrorHandler

    If pxmlOperacao.documentElement.selectSingleNode("PROD_CAMR") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "PROD_CAMR", penumProdutoCamara
    Else
        pxmlOperacao.documentElement.selectSingleNode("PROD_CAMR").Text = penumProdutoCamara
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgCriarTAGTransferencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Gera o registro de alerta
Public Function fgGerarAlerta(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                     Optional ByRef xmlMensagem As MSXML2.DOMDocument40, _
                     Optional ByVal pstrSituacaoMensagem As String, _
                     Optional ByVal penumFatorGerador As enumFatorGeradorAlerta, _
                     Optional ByVal plngStatusOperacao As Long = 0) As Boolean

Dim objAlerta                               As A8LQS.clsAlerta

On Error GoTo ErrorHandler
    
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    fgGerarAlerta = objAlerta.GerarAlerta(xmlOperacao, xmlMensagem, pstrSituacaoMensagem, penumFatorGerador, plngStatusOperacao)
    Set objAlerta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objAlerta = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgGerarAlerta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gera o registro de alerta na conciliação
Private Sub flGeraAlertaConciliacao(ByVal pstrCodigoMensagem As String, _
                                    ByVal plngTipoBackOffice As Long, _
                                    ByVal pstrNomeVeiculoLegal As String, _
                                    ByVal pstrValor As String, _
                           Optional ByVal plngStatusMensagem As Long = 0, _
                           Optional ByVal vntNumeroSeqOperacao As Variant = "")
                               
Dim objAlerta                               As A8LQS.clsAlerta
Dim xmlPropriedadesAlerta                   As MSXML2.DOMDocument40
Dim lngCodFatorGeraAlerta                   As Long
                              
                               
On Error GoTo ErrorHandler
                               
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    Set xmlPropriedadesAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlPropriedadesAlerta.loadXML objAlerta.ObterPropriedades
    
    xmlPropriedadesAlerta.selectSingleNode("//NU_SEQU_OPER_ATIV").Text = vntNumeroSeqOperacao
    xmlPropriedadesAlerta.selectSingleNode("//CO_FATO_GERA_ALER").Text = lngCodFatorGeraAlerta
    xmlPropriedadesAlerta.selectSingleNode("//TP_BKOF").Text = plngTipoBackOffice
    xmlPropriedadesAlerta.selectSingleNode("//NO_VEIC_LEGA").Text = pstrNomeVeiculoLegal
    xmlPropriedadesAlerta.selectSingleNode("//VA_OPER_ATIV").Text = pstrValor
    xmlPropriedadesAlerta.selectSingleNode("//TX_ANEX").Text = ""

    Call objAlerta.GerarAlertaCamara(xmlPropriedadesAlerta)
                               
    Set objAlerta = Nothing
    Set xmlPropriedadesAlerta = Nothing

    Exit Sub
ErrorHandler:
    
    Set objAlerta = Nothing
    Set xmlPropriedadesAlerta = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "flGeraAlertaCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Verifica se a data é dia util
Public Function fgValidaDataUtil(ByVal pintCodigoBanco As Integer, _
                                 ByVal plngAgencia As Long, _
                                 ByVal pdatDataBase As Date, _
                                 ByRef plngCodErro As Long, _
                                 ByRef pstrMensagemErro As String) As Date
On Error GoTo ErrorHandler

Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
    
    Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    
    fgValidaDataUtil = objA6A7A8Funcoes.ValidaDataUtil(pintCodigoBanco, _
                                                       plngAgencia, _
                                                       pdatDataBase, _
                                                       plngCodErro, _
                                                       pstrMensagemErro)
    
    Set objA6A7A8Funcoes = Nothing
    
    Exit Function
ErrorHandler:
    Set objA6A7A8Funcoes = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgValidaDataUtil", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Alterar dados para o sistema A6 de operações de volta
Public Function fgAlterarDadosA6Volta(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                      ByVal pstrDataTagOriginal As String, _
                                      ByVal pstrDataTagNova As String, _
                                      ByVal pblnAlterarData As Boolean) As Boolean

Dim strData                                 As String
Dim strValor                                As String

    On Error GoTo ErrorHandler

    strData = xmlRemessa.documentElement.selectSingleNode("//" & pstrDataTagOriginal).Text
    xmlRemessa.documentElement.selectSingleNode("//" & pstrDataTagOriginal).Text = xmlRemessa.documentElement.selectSingleNode("//" & pstrDataTagNova).Text
    xmlRemessa.documentElement.selectSingleNode("//" & pstrDataTagNova).Text = strData

    If pblnAlterarData Then
        Exit Function
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
        xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing Then
        If Val("0" & xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text) > 0 Then
            strValor = xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text
            xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text
            xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text = strValor
        End If
    End If

    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgAlterarDadosA6Volta", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter a quantidade de dias uteis de Expurgo
Public Function fgObterQtdDiasExpurgo() As Long

Dim strXMLParmGeral                         As String
Dim xmlParmGeral                            As DOMDocument40

On Error GoTo ErrorHandler

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
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterQtdDiasExpurgo", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verificar se a Operação é uma Operação BMA
Function fgTipoOperacaoBMA(ByVal lngTipoOperacao) As Boolean
    
    'Verifica se um 'TipoDeOperacao' pertence à BMA

    fgTipoOperacaoBMA = fgIN(lngTipoOperacao, _
        enumTipoOperacaoLQS.TransferenciaBMA, enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
        enumTipoOperacaoLQS.DefinitivaDescobertaBMA, enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
        enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA, enumTipoOperacaoLQS.LiquidacaoOperacaoTermoBMA, _
        enumTipoOperacaoLQS.LiquidacaoEventosJurosBMA, enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA, _
        enumTipoOperacaoLQS.OperacaoTermoInternaBMA, enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
        enumTipoOperacaoLQS.EspecDefinitivaCobertura, enumTipoOperacaoLQS.EspecTermoIntermediacao, _
        enumTipoOperacaoLQS.EspecTermoCobertura, enumTipoOperacaoLQS.DepositoBMA, _
        enumTipoOperacaoLQS.RetiradaBMA, enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA, _
        enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA, enumTipoOperacaoLQS.CancelamentoEspecificacaoBMA, _
        enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA, enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA, _
        enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA, enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA, _
        enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
        enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA, enumTipoOperacaoLQS.CompromissadaEspecificaTermo, _
        enumTipoOperacaoLQS.LiquidacaoCompromissadaEspecificaVolta, enumTipoOperacaoLQS.LiquidacaoEventosJurosTituloCompro, _
        enumTipoOperacaoLQS.OperacaoCompromissadaInternaBMA, enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA, _
        enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA)

End Function

'Verificar o erro oracle no MQ Series
Public Function fgVerificaErroOracleMQSeries(ByVal pstrErro As String) As Boolean

Dim xmlErro                                 As MSXML2.DOMDocument40
   
On Error GoTo ErrorHandler
   
   Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
   
   If xmlErro.loadXML(pstrErro) Then
       If xmlErro.selectSingleNode("//ErrorType") Is Nothing Then
           fgVerificaErroOracleMQSeries = False
       Else
           If xmlErro.selectSingleNode("//ErrorType").Text = "2" Then
               fgVerificaErroOracleMQSeries = True
               If xmlErro.selectSingleNode("//Number").Text = 36 Then
                   fgVerificaErroOracleMQSeries = True
               End If
           End If
       End If
   Else
       fgVerificaErroOracleMQSeries = True
   End If
    
   Set xmlErro = Nothing
   Exit Function
   
ErrorHandler:
   Set xmlErro = Nothing
   Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Obter o fluxo de mensagem SPB
Public Function fgObterTipoFluxoMensagemSPB(ByVal pstrCodigoMensagem As String) As Integer

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT SQ_TIPO_FLUX " & _
             "   FROM A8.TB_MENSAGEM A, " & _
             "        A8.TB_EVENTO B " & _
             "  WHERE A.SQ_EVEN = B.SQ_EVEN " & _
             "    AND A.CO_MESG = '" & Trim(pstrCodigoMensagem) & "'"
             
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        fgObterTipoFluxoMensagemSPB = objRS.fields("SQ_TIPO_FLUX")
    End If
    
    objRS.Close

    Exit Function
ErrorHandler:


    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS ", "fgObterTipoFluxoMensagemSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar se a mensagem SPB é externa
Public Function fgMensagemExterna(ByVal pstrCodigoMensagemSPB As String) As Boolean

Dim lngTipoFluxoMensagem                    As Long

On Error GoTo ErrorHandler

    If Right$(Trim$(pstrCodigoMensagemSPB), 2) = "R1" Or _
       Right$(Trim$(pstrCodigoMensagemSPB), 2) = "R2" Then
        fgMensagemExterna = True
    Else
        lngTipoFluxoMensagem = fgObterTipoFluxoMensagemSPB(Trim$(pstrCodigoMensagemSPB))
        
        If lngTipoFluxoMensagem = enumTipoFluxo.TipoFluxo5 Or _
           lngTipoFluxoMensagem = enumTipoFluxo.TipoFluxo7 Then
            
            fgMensagemExterna = True
        End If
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS ", "fgMensagemExterna Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Gravar arquivo de log de erro
Public Sub fgGravaArquivo(ByVal pstrNomeArquivo As String, _
                          ByVal pstrErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strMensagem                              As String

On Error GoTo ErrorHandler

    strMensagem = String$(50, "*") & vbCrLf
    strMensagem = strMensagem & pstrErro & vbCrLf
    strMensagem = strMensagem & String$(50, "*")

    'Log de Erro será gerado na pasta \ LogErro do diretório Server do SLCC.
    'strNomeArquivoLogErro = App.Path & "\log\" & pstrNomeArquivo & "_" & Format(Now, "yyyymmddHHmmss") & ".log"
    
    strNomeArquivoLogErro = App.Path & "\LogErro\" & pstrNomeArquivo & "_" & Format(Now, "yyyymmddHHmmss") & ".log"
    
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, strMensagem
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'Obter o codigo do banco - compensação
Public Function fgObterBanco(ByVal plngCodigoEmpresa As Long) As Long

On Error GoTo ErrorHandler

    
    Select Case plngCodigoEmpresa
        Case enumCodigoEmpresa.Banespa
            fgObterBanco = 33
        Case enumCodigoEmpresa.Bozano
            fgObterBanco = 351
        Case enumCodigoEmpresa.Meridional
            fgObterBanco = 8
        Case enumCodigoEmpresa.Santander
            fgObterBanco = 353
        Case Else
            fgObterBanco = plngCodigoEmpresa
    End Select
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7Server", "fgObterBanco Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgObterValorMensagemPelaTag(ByRef pxmlMensagem As MSXML2.DOMDocument40) As String

Dim strValorMensagem                        As String

    On Error GoTo ErrorHandler

    strValorMensagem = vbNullString
    
    If Not pxmlMensagem.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
        strValorMensagem = pxmlMensagem.selectSingleNode("//VA_OPER_ATIV").Text
    ElseIf Not pxmlMensagem.selectSingleNode("//VA_FINC") Is Nothing Then
        strValorMensagem = pxmlMensagem.selectSingleNode("//VA_FINC").Text
    ElseIf Not pxmlMensagem.selectSingleNode("//VlrLanc") Is Nothing Then
        strValorMensagem = pxmlMensagem.selectSingleNode("//VlrLanc").Text
    ElseIf Not pxmlMensagem.selectSingleNode("//VlrNLiqdant") Is Nothing Then
        strValorMensagem = pxmlMensagem.selectSingleNode("//VlrNLiqdant").Text
    ElseIf Not pxmlMensagem.selectSingleNode("//VlrResultLiqdNLiqdant") Is Nothing Then
        strValorMensagem = pxmlMensagem.selectSingleNode("//VlrResultLiqdNLiqdant").Text
    End If
    
    If fgVlrXml_To_Decimal(strValorMensagem) = 0 Then
        
        If Not pxmlMensagem.selectSingleNode("//CodMsg") Is Nothing Then
            If pxmlMensagem.selectSingleNode("//CodMsg").Text = "BMC0101" Then
                
                If Not pxmlMensagem.selectSingleNode("//VA_MOED_ESTR") Is Nothing Then
                    strValorMensagem = pxmlMensagem.selectSingleNode("//VA_MOED_ESTR").Text
                End If
        
            ElseIf pxmlMensagem.selectSingleNode("//CodMsg").Text = "BMC0102" Then
                
                If Not pxmlMensagem.selectSingleNode("//VlrME") Is Nothing Then
                    strValorMensagem = pxmlMensagem.selectSingleNode("//VlrME").Text
                End If
                
            End If
        End If
    
    End If
        
    If Not pxmlMensagem.selectSingleNode("//CodMsg") Is Nothing Then
        If pxmlMensagem.selectSingleNode("//CodMsg").Text = "LDL0004" Then
            If Not pxmlMensagem.selectSingleNode("//VlrLanc") Is Nothing Then
                strValorMensagem = pxmlMensagem.selectSingleNode("//VlrLanc").Text
            ElseIf Not pxmlMensagem.selectSingleNode("//VA_FINC") Is Nothing Then
                strValorMensagem = pxmlMensagem.selectSingleNode("//VA_FINC").Text
            ElseIf Not pxmlMensagem.selectSingleNode("//VlrNLiqdant") Is Nothing Then
                strValorMensagem = pxmlMensagem.selectSingleNode("//VlrNLiqdant").Text
            ElseIf Not pxmlMensagem.selectSingleNode("//VlrResultLiqdNLiqdant") Is Nothing Then
                strValorMensagem = pxmlMensagem.selectSingleNode("//VlrResultLiqdNLiqdant").Text
            End If
        End If
    End If
        
    fgObterValorMensagemPelaTag = strValorMensagem

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS ", "fgObterValorMensagemPelaTag Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgObterDataLiquidacaoPelaTag(ByRef pxmlMensagem As MSXML2.DOMDocument40) As String

Dim strDataLiquidacao                       As String

    On Error GoTo ErrorHandler

    strDataLiquidacao = vbNullString
    
    If Not pxmlMensagem.selectSingleNode("//DT_LIQU") Is Nothing Then
        strDataLiquidacao = pxmlMensagem.selectSingleNode("//DT_LIQU").Text
    ElseIf Not pxmlMensagem.selectSingleNode("//DT_LIQU_OPER_ATIV_MOED_ESTR") Is Nothing Then
        strDataLiquidacao = pxmlMensagem.selectSingleNode("//DT_LIQU_OPER_ATIV_MOED_ESTR").Text
    End If
    
    fgObterDataLiquidacaoPelaTag = strDataLiquidacao

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS ", "fgObterDataLiquidacaoPelaTag Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Sub fgWait(Optional ByVal intSeconds As Integer = 1)

Dim datDataInicial                          As Date

    datDataInicial = fgDataHoraServidor(DataHoraAux)
    
    Do
        If DateDiff("s", datDataInicial, fgDataHoraServidor(DataHoraAux)) > intSeconds Or intSeconds > 30 Then
            Exit Do
        End If
    Loop
    
    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS ", "fgWait Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Public Function fgAlterarCodigoProduto(ByVal plngCodigoProduto As Long) As Long

    fgAlterarCodigoProduto = IIf(plngCodigoProduto = 507, 529, IIf(plngCodigoProduto = 529, 507, _
                             IIf(plngCodigoProduto = 313, 315, IIf(plngCodigoProduto = 315, 313, _
                             IIf(plngCodigoProduto = 312, 314, IIf(plngCodigoProduto = 314, 312, _
                             IIf(plngCodigoProduto = 343, 344, IIf(plngCodigoProduto = 344, 343, _
                             IIf(plngCodigoProduto = 341, 342, IIf(plngCodigoProduto = 342, 341, _
                             IIf(plngCodigoProduto = 177, 244, IIf(plngCodigoProduto = 244, 177, _
                             IIf(plngCodigoProduto = 69, 243, IIf(plngCodigoProduto = 243, 69, _
                             IIf(plngCodigoProduto = 182, 340, IIf(plngCodigoProduto = 340, 182, _
                             IIf(plngCodigoProduto = 181, 339, IIf(plngCodigoProduto = 339, 181, _
                             IIf(plngCodigoProduto = 337, 338, IIf(plngCodigoProduto = 338, 337, _
                             IIf(plngCodigoProduto = 180, 179, IIf(plngCodigoProduto = 179, 180, _
                             IIf(plngCodigoProduto = 70, 178, IIf(plngCodigoProduto = 178, 70, _
                             IIf(plngCodigoProduto = 679, 680, IIf(plngCodigoProduto = 680, 679, _
                             IIf(plngCodigoProduto = 681, 682, IIf(plngCodigoProduto = 682, 681, _
                             IIf(plngCodigoProduto = 683, 684, IIf(plngCodigoProduto = 684, 683, _
                             IIf(plngCodigoProduto = 685, 686, IIf(plngCodigoProduto = 686, 685, plngCodigoProduto _
                             ))))))))))))))))))))))))))))))))

End Function


'Adicionar dias uteis a uma data
Public Function fgAdicionarDiasUteis(ByVal pdatData As Date, _
                                     ByVal pintQtdeDias As Integer, _
                                     ByVal plngMovimento As enumPaginacao) As Date


Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes

On Error GoTo ErrorHandler
    
    Set objA6A7A8Funcoes = CreateObject("A6A7a8.clsA6A7A8Funcoes")
    fgAdicionarDiasUteis = objA6A7A8Funcoes.AdicionarDiasUteis(pdatData, _
                                                               pintQtdeDias, _
                                                               plngMovimento)
                                                             
    Set objA6A7A8Funcoes = Nothing

    Exit Function
ErrorHandler:
    Set objA6A7A8Funcoes = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS ", "fgAdicionarDiasUteis Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter a Codigo SISBACEN da Camara
Public Function fgObterCodSISBACENCamr() As Long

'Dim strXMLParmGeral                         As String
'Dim xmlParmGeral                            As DOMDocument40

On Error GoTo ErrorHandler

    fgObterCodSISBACENCamr = 14805

'    strXMLParmGeral = fgSelectVarchar4000(0, False)
'    Set xmlParmGeral = CreateObject("MSXML2.DOMDocument.4.0")
'
'    If Not xmlParmGeral.loadXML(strXMLParmGeral) Then
'        GoTo ErrorHandler
'    End If
'
'    If xmlParmGeral.selectSingleNode("//BASE_HISTORICA") Is Nothing Then
'        'Valor default
'        fgObterCodSISBACENCamr = 40
'    Else
'        fgObterCodSISBACENCamr = CLng(xmlParmGeral.selectSingleNode("//BASE_HISTORICA").Text)
'    End If
'
'    Set xmlParmGeral = Nothing

    Exit Function
ErrorHandler:
'    Set xmlParmGeral = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterCodSISBACENCamr", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter a Codigo Cancelamento BMC0002
Public Function fgObterCodCancelBMC0002() As Long

'Dim strXMLParmGeral                         As String
'Dim xmlParmGeral                            As DOMDocument40

On Error GoTo ErrorHandler

    fgObterCodCancelBMC0002 = 1

'    strXMLParmGeral = fgSelectVarchar4000(0, False)
'    Set xmlParmGeral = CreateObject("MSXML2.DOMDocument.4.0")
'
'    If Not xmlParmGeral.loadXML(strXMLParmGeral) Then
'        GoTo ErrorHandler
'    End If
'
'    If xmlParmGeral.selectSingleNode("//BASE_HISTORICA") Is Nothing Then
'        'Valor default
'        fgObterCodSISBACENCamr = 40
'    Else
'        fgObterCodSISBACENCamr = CLng(xmlParmGeral.selectSingleNode("//BASE_HISTORICA").Text)
'    End If
'
'    Set xmlParmGeral = Nothing

    Exit Function
ErrorHandler:
'    Set xmlParmGeral = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterCodCancelBMC0002", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter a Codigo Banqueiro Swift para Layout de PJ Moeda Estrangeira
Public Function fgObterCodigoBanqueiroSwiftPJME() As String

On Error GoTo ErrorHandler

    fgObterCodigoBanqueiroSwiftPJME = "IRVTUS3N"

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterCodigoBanqueiroSwiftPJME", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter Nome do Cliente para layout PJ Moeda Estrangeira
Public Function fgObterNomeClientePJME() As String


On Error GoTo ErrorHandler

    fgObterNomeClientePJME = "The bank of NY"

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterNomeClientePJME", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter Conta Banqueiro para layout PJ Moeda Estrangeira
Public Function fgObterContaBanqueiroPJME(ByVal intTipoMensagem As Integer) As Double


On Error GoTo ErrorHandler
    
    If TipoMensagem = enumTipoMensagemBUS.EmissaoOperacaoCCR Or _
       TipoMensagem = enumTipoMensagemBUS.NegociacaoOperacaoCCR Or _
       TipoMensagem = enumTipoMensagemLQS.DevolucaoRecolhimentoEstornoReembolsoCCR Then
        fgObterContaBanqueiroPJME = Val("8033039558")
    Else
    
        fgObterContaBanqueiroPJME = Val("8900560681")
    
    End If
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgObterContaBanqueiroPJME", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
