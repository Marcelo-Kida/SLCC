Attribute VB_Name = "basRBLibrary"

' Este componente tem como objetivo agrupar métodos úteis a quaisquer sistemas desenvolvidos pela empresa.
' Biblioteca generalizada de funcionalidades.

Option Explicit

Public Const gstrCaixaSubReserva            As String = "Caixa Sub-reserva"
Public Const gstrCaixaFuturo                As String = "Caixa Futuro"
Public Const gstrItemGenerico               As String = "Item Genérico"

Public Const gstrSalvar                     As String = "Salvar"
Public Const gstrSair                       As String = "Sair"
Public Const gstrAtualizar                  As String = "Atualizar"

Public Const gstrDataVazia                  As String = "00:00:00"

Global gblnLog                              As Boolean
'------------------ API ----------------------------------------
'API'S para geração de Arquivo temporário no Windows \Temp
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'API para verificar o tempo de execução de uma rotina ou processo
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'API para abertura de arquivo Help
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Const HH_DISPLAY_TOPIC = &H0

'Constantes para formatação
Public Const gstrMaskInt                    As String = "##,###,###,###,###,##0;[vbRed]-##,###,###,###,###,##0"
Public Const gstrMaskDec                    As String = "##,###,###,###,###,##0.00;[vbRed] -##,###,###,###,###,##0.00"
Public Const gstrMaskDecN                   As String = "##,###,###,###,###,##0.CASAS;[vbRed] -##,###,###,###,###,##0.CASAS"  'para formatacao com N cadas decimais

'Este Enumerator esta duplicado para possibilitar a compilação do
'Componente A6A7A8CA - Adilson - 08/10/2003
Public Enum enumFormatoDataHoraAux
    DataAux = 1
    HoraAux = 2
    DataHoraAux = 3
End Enum

Private Type udtCodigoItemCaixa
    TP_CAIX                                 As String * 1
    CO_ITEM_CAIX_NIVE_01                    As String * 3
    CO_ITEM_CAIX_NIVE_02                    As String * 3
    CO_ITEM_CAIX_NIVE_03                    As String * 3
    CO_ITEM_CAIX_NIVE_04                    As String * 3
    CO_ITEM_CAIX_NIVE_05                    As String * 3
End Type

Private Type udtCodigoItemCaixaAux
    CO_ITEM_CAIX                            As String * 16
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

Public Type udtProtocoloErroNZ
    CodigoMensagem                As String * 9
    ControleRemessaNZ             As String * 20
    CodigoEmpresa                 As String * 5
    OrigemErro                    As String * 1
    DataRemessa                   As String * 8
    NomeDoCampo1                  As String * 80
    CodigoErro1                   As String * 9
    ConteúdoCampoErro1            As String * 30
    NomeDoCampo2                  As String * 80
    CodigoErro2                   As String * 9
    ConteúdoCampoErro2            As String * 30
    NomeDoCampo3                  As String * 80
    CodigoErro3                   As String * 9
    ConteúdoCampoErro3            As String * 30
    NomeDoCampo4                  As String * 80
    CodigoErro4                   As String * 9
    ConteúdoCampoErro4            As String * 30
    NomeDoCampo5                  As String * 80
    CodigoErro5                   As String * 9
    ConteúdoCampoErro5            As String * 30
End Type

Public Type udtProtocoloErroNZAux
    String                        As String * 638
End Type

Public Type udtProtocoloErroNZExt
    CodigoMensagem                As String * 9
    ControleRemessaPZ             As String * 20
    CodigoEmpresa                 As String * 5
    OrigemErro                    As String * 1
    DataRemessa                   As String * 8
    NomeDoCampo1                  As String * 80
    CodigoErro1                   As String * 9
    ConteúdoCampoErro1            As String * 30
    NomeDoCampo2                  As String * 80
    CodigoErro2                   As String * 9
    ConteúdoCampoErro2            As String * 30
    NomeDoCampo3                  As String * 80
    CodigoErro3                   As String * 9
    ConteúdoCampoErro3            As String * 30
    NomeDoCampo4                  As String * 80
    CodigoErro4                   As String * 9
    ConteúdoCampoErro4            As String * 30
    NomeDoCampo5                  As String * 80
    CodigoErro5                   As String * 9
    ConteúdoCampoErro5            As String * 30
    ControleRemessaNZ             As String * 20
End Type

Public Type udtProtocoloErroNZExtAux
    String                        As String * 658
End Type

'---------------------------------------------------------------------------------------------------
'INICIO INTEGRAÇÃO PZ
'----------------------------------------------------------------------------------------------------

'Devolução

Public Type udtSTR0010          'Layout PZW0004 ***
    Cod_Tipo_Reg                As String * 1
    BANCO                       As String * 3
    Pontovenda                  As String * 7
    SiglaSistema                As String * 3
    CodSistema                  As String * 4
    NumeroControleSistemaLegado As String * 23
    NDoc                        As String * 6
    BancoDest                   As String * 3
    IspbDest                    As String * 8
    DataContabil                As String * 8
    HoraAgen                    As String * 6
    CodAcatamento               As String * 1
    CodMeioTrans                As String * 1
    Flag                        As String * 1
    NumVerCtrl                  As String * 2
    FlagCC                      As String * 1
    CodHst                      As String * 5
    CDeb                        As String * 13
    AgeDeb                      As String * 5
    CodigoOri                   As String * 2
    Filler1                     As String * 39
    Erro1                       As String * 5
    Erro2                       As String * 5
    Erro3                       As String * 5
    CodPz                       As String * 50
    CodMsg                      As String * 9
    VlrLanc                     As String * 18
    DtMovto                     As String * 8
    AgCred                      As String * 5
    AgRem                       As String * 5
    Filler2                     As String * 40
    NumPagto                    As String * 15
    CodTransferencia            As String * 2
    HISTORICO                   As String * 200
    NivelPref                   As String * 1
    Filler3                     As String * 390
End Type
 
Public Type udtSTR0010aux
   MENSAGEM         As String * 900
End Type

'***** Inicio Layout  PZW0916 - Retorno HEADER NZ - PZ

Public Type udtPZW0916
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
    NuOP                          As String * 24
    IndicadorDupli                As String * 1
    IndicadorValorCorte           As String * 1
    IndicadorValorMinimo          As String * 1
    NumeroControleLegado          As String * 23
    CodSituTransfPagto            As String * 2
    FILLER                        As String * 15
End Type

Public Type udtPZW0916Aux
    String                        As String * 200
End Type

'****** Inicio Layout  PZO0140. - Retorno R1 PZ
Public Type udtPZO00140
    HeaderNZ_PZ         As String * 200  '<== udtPZW0916
    CodMsg              As String * 9    '   A   9
    NumCtrlIF           As String * 20   '   A   20
    ISPBIFDebtd         As String * 8    '   A   8
    NumCtrlSTR          As String * 20   '   A   20
    SitLancSTR          As String * 3    '   N   3
    DtHrSit             As String * 14   '   N   14
    DtMovto             As String * 8    '   N   8
    FILLER              As String * 316  '   A   316
End Type

Public Type udtPZO00140Aux
    String              As String * 600
End Type

'***** Inicio Layout  PZO0141. - Retorno Erro externo PZ

Public Type udtPZO00141
    Retorno             As String * 2
    HeaderNZ_PZ_ERRO    As String * 750
End Type

Public Type udtPZO00141Aux
    String             As String * 752
End Type

'**********************************


'***** Inicio Layout PZW0001
Public Type udtPZW0001
    TipoRegistro                            As String * 1
    BancoOrigem                             As String * 3
    UnidadeOrigem                           As String * 7
    SiglaSistemaOrigem                      As String * 3
    CodigoSistemaOrigem                     As String * 4
    ControleLegado                          As String * 23
    NumeroDocumento                         As String * 6
    BancoDestino                            As String * 3
    ISPBDestino                             As String * 8
    DataContabil                            As String * 8
    HoraAgendamento                         As String * 6
    CodigoAcao                              As String * 1
    MeioTransferencia                       As String * 1
    IdentificadorAlteracao                  As String * 1
    NumeroVersao                            As String * 2
    LancamentoCC                            As String * 1
    HistoricoCC                             As String * 5
    ContaDebitada                           As String * 13
    AgeciaDebitada                          As String * 5
    OrigemContabilidade                     As String * 2
    Filler1                                 As String * 39
    Erro1                                   As String * 5
    Erro2                                   As String * 5
    Erro3                                   As String * 5
    CodigoPZ                                As String * 50
    CodigoMensagem                          As String * 9
    ValorLancamento                         As String * 18
    DataMovimento                           As String * 8
    AgenciaCreditada                        As String * 5
    AgenciaRemetente                        As String * 5
    Filler2                                 As String * 40
End Type

Public Type udtPZW0001Aux
   String                                   As String * 292
End Type


'Layout STR0008R2

Public Type udtSTR0008R2
    NU_RECB_PGTO                            As String * 15 'PIC X(015)
    CO_EMPR                                 As String * 4  'PIC 9(004)
    DT_RECB_PGTO                            As String * 8  'PIC 9(008)
    HO_RECB_MESG                            As String * 6  'PIC 9(006)
    VA_RECB_PGTO                            As String * 18 'PIC 9(018)
    CO_MESG                                 As String * 9  'PIC X(009)
    CO_ISPB_REMT                            As String * 8  'PIC X(008)
    CO_BANC_REMT                            As String * 4  'PIC 9(004)
    NO_REMT                                 As String * 80 'PIC X(080)
    NU_CNPJ_CPF_REMT                        As String * 15 'PIC 9(015)
    TP_CNTA_REMT                            As String * 2  'PIC X(002)
    CO_AGEN_REMT                            As String * 9  'PIC 9(009)
    DG_AGEN_REMT                            As String * 1  'PIC X(001)
    NU_CNTA_REMT                            As String * 12 'PIC 9(012)
    DG_CNTA_REMT                            As String * 1  'PIC X(001)
    NO_DEST                                 As String * 80 'PIC X(080)
    NU_CNPJ_CPF_DEST                        As String * 15 'PIC 9(015)
    IN_TIPO_PESS_DEST                       As String * 1  'PIC X(001)
    TP_CNTA_DEST                            As String * 2  'PIC X(002)
    CO_AGEN_DEST                            As String * 9  'PIC 9(009)
    NU_CNTA_DEST                            As String * 13 'PIC 9(013)
    SQ_TIPO_TAG_FIND                        As String * 4  'PIC 9(004)
    CO_DOMI_FIND                            As String * 20 'PIC X(020)
    TX_HIST_COMP_RECB                       As String * 200 'PIC X(200)
    CO_SITU_RECB_PGTO                       As String * 4  'PIC 9(004)
    NU_CTRL_REME                            As String * 20 'PIC X(020)
    NU_CTRL_EXTE_RECB                       As String * 20 'PIC X(020)
    NU_DOCT_CRED                            As String * 9  'PIC 9(009)
End Type

Public Type udtSTR0008R2Aux
    String                                  As String * 589
End Type

Public Type udtConsultaPZ
    BANCO                                   As String * 3  'PIC 9(003)
    AGENCIA                                 As String * 5  'PIC 9(005)
    CONTA                                   As String * 13 'PIC 9(013)
    DT_MENSAGE                              As String * 8  'PIC 9(008)
    Hora                                    As String * 8  'PIC X(008)
    CO_MENSAGEM                             As String * 9  'PIC X(009)
    SG_SIST                                 As String * 3  'PIC X(003)
    FILLER                                  As String * 20
    QT_ITEM                                 As String * 5  'PIC 9(005)
    RC_ROTINA                               As String * 5  'PIC 9(005)
    MSG_RC_ROTINA                           As String * 80  'PIC X(080)
End Type

Public Type udtConsultaPZAux
    String                                  As String * 159
End Type


'---------------------------------------------------------------------------------------------------
'FIM INTEGRAÇÃO PZ
'----------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------
'INICIO INTEGRAÇÃO ARQUIVOS CBLC AMDF - ALCO
'-----------------------------------------------------------------------------------------------------------
Public Type udtAMDF_ALCOHeader
    TipoRegistro                            As String * 2
    CodigoArquivo                           As String * 4   'AMDF
    CodigoUsuario                           As String * 4   'Usuario BOVESPA/CBLC
    CodigoOrigem                            As String * 8   'BOV/CBLC
    CodigoDestino                           As String * 4
    NumeroMovimento                         As String * 9
    DataGeracao                             As String * 8
    FILLER                                  As String * 261
End Type

Public Type udtAMDF_ALCOHeaderAux
    String                                  As String * 300
End Type


Public Type udtAMDF_ALCOLancamento
    TipoRegistro                            As String * 2
    DataEfetivacao                          As String * 8
    CodigoEmpresaResponsavel                As String * 4
    DescricaoEmpresaResponsavel             As String * 15
    CodigoGrupo                             As String * 5
    DescricaoGrupo                          As String * 15
    CodigoLancamentoFinanceriro             As String * 5
    DescricaoLancementoFincanceiro          As String * 60
    TipoLancamento                          As String * 1
    CodigoClienteQualificado                As String * 7
    FILLER                                  As String * 26
    CodigoCorretora                         As String * 5
    NomeCorretora                           As String * 30
    TipoMoeda                               As String * 5
    DescricaoMoeda                          As String * 15
    ValorLancamento                         As String * 18
    LancamentoParaQualificado               As String * 1
    FormaPagamento                          As String * 5
    SituacaoLancamento                      As String * 1
    DescricaoSituLancamento                 As String * 15
    BancoLiquidante                         As String * 8
    IdentificacaoLancamento                 As String * 26
    Filler2                                 As String * 23
End Type

Public Type udtAMDF_ALCOLancamentoAux
    String                                  As String * 300
End Type

Public Type udtAMDF_ALCOTrailer
    TipoRegistro                            As String * 2
    CodigoArquivo                           As String * 4
    CodigoUsuario                           As String * 4
    CodigoOrigem                            As String * 8
    CodigoDestino                           As String * 4
    NumeroMovimento                         As String * 9
    DataGeracao                             As String * 8
    TotalRegistros                          As String * 9
    FILLER                                  As String * 252
End Type

Public Type udtAMDF_ALCOTrailerAux
    String                                  As String * 300
End Type

'-----------------------------------------------------------------------------------------------------------
'FIM INTEGRAÇÃO ARQUIVOS CBLC AMDF - ALCO
'-----------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------
'INICIO INTEGRAÇÃO ARQUIVOS BMF T025
'-----------------------------------------------------------------------------------------------------------
Public Type udtT025_Header
    IdentificaoTransacao                    As String * 6
    ComplementoTransacao                    As String * 3
    TipoRegistro                            As String * 2
    DataGeracao                             As String * 8
    DataNetting                             As String * 8
    MembroCompensacao                       As String * 6
    CodigoCorretora                         As String * 6
    CodigoCliente                           As String * 6
    CodigoMercadoria                        As String * 3
End Type

Public Type udtT025_HeaderAux
    String                                  As String * 48
End Type

Public Type udtT025_Lancamento
    IdentificaoTransacao                    As String * 6
    ComplementoTransacao                    As String * 3
    TipoRegistro                            As String * 2
    CodigoLancamento                        As String * 3
    DebitoCredito                           As String * 1
    ValorLancamento                         As String * 15
End Type

Public Type udtT025_LancamentoAux
    String                                  As String * 30
End Type

'-----------------------------------------------------------------------------------------------------------
'FIM INTEGRAÇÃO ARQUIVOS BMF T025
'-----------------------------------------------------------------------------------------------------------

Public Enum enumFilasEntrada
    enumA7QEENTRADA = 1
    enumA7QEMENSAGEMRECEBIDA = 2
    enumA7QEREPORT = 3
    enumA8QEENTRADA = 4
    enumA6QEREMESSASUBRESERVA = 5
    enumA6QEREMESSAFUTURO = 6
End Enum


' Retorna o tipo de componente utilizado no momento.

Public Function fgObterTipoComponente() As String

    Select Case UCase$(App.Title)
        Case "A6", "A7BUSCLIENT", "A8"
            fgObterTipoComponente = "I"
        Case "A6SUBRESERVA", "A7SERVER", "A8LQS"
            fgObterTipoComponente = "R"
        Case "A6MIU", "A7MIU", "A8MIU"
            fgObterTipoComponente = "C"
        Case "A6A7A8"
            fgObterTipoComponente = "RC"
        Case "A6A7A8MIU"
            fgObterTipoComponente = "CC"
        Case "A6A7A8CA"
            fgObterTipoComponente = "D"
    End Select

End Function

' Retorna dependências de componentes.

Public Function fgObterDependecias() As String

    Select Case UCase$(App.Title)
        Case "A8"
            fgObterDependecias = "A8MIU"
        Case "A8LQS"
            fgObterDependecias = "A6A7A8;A6A7A8CA"
        Case "A8MIU"
            fgObterDependecias = "A8LQS"
        Case "A6"
            fgObterDependecias = "A6MIU"
        Case "A6MIU"
            fgObterDependecias = "A6SUBRESERVA"
        Case "A6SUBRESERVA"
            fgObterDependecias = "A6A7A8;A6A7A8CA"
        Case "A7"
            fgObterDependecias = "A7MIU"
        Case "A7MIU"
            fgObterDependecias = "A7SERVER"
        Case "A7SERVER"
            fgObterDependecias = "A6A7A8;A6A7A8CA"
        Case "A6A7A8"
            fgObterDependecias = "A6A7A8CA"
        Case "A6A7A8MIU"
            fgObterDependecias = "A6A7A8"
        Case "A6A7A8CA"
            fgObterDependecias = vbNullString
    End Select

End Function

' Adiona um node a um objeto XML.

Public Function fgAppendNode(ByRef xmlDocument As MSXML2.DOMDocument40, _
                             ByVal pstrNodeContext As String, _
                             ByVal pstrNodeName As String, _
                             ByVal pstrNodeValue As String, _
                    Optional ByVal pstrNodeRepetName As String = "") As Boolean

Dim objDomNodeAux                           As MSXML2.IXMLDOMNode
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    If pstrNodeName = vbNullString Then
        'Parâmetro pstrNodeName deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrNodeName deve ser diferente de vbNullString"
    End If
    
    If pstrNodeContext = "" Then
        'Se a Tag for o root passar pstrNodeContext = Nome Tag Principal
        Set objDomNodeContext = xmlDocument
    Else
        'Para criar um Grupo (Ex.:Grupo_Usuario) dentro de Uma Repeticao (Ex. Repeat_Usuario)
        'Passar o argumento pNomeRepet = Repet_Usuario
        If pstrNodeRepetName <> "" Then
            Set objDomNodeContext = xmlDocument.selectSingleNode("//" & pstrNodeRepetName).childNodes.Item(xmlDocument.selectSingleNode("//" & pstrNodeRepetName).childNodes.length - 1)
        Else
            Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
        End If
    End If
    
    Set objDomNodeAux = xmlDocument.createElement(pstrNodeName)
    objDomNodeAux.Text = pstrNodeValue
    objDomNodeContext.appendChild objDomNodeAux

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

' Adiona um atributo a um objeto XML.

Public Function fgAppendAttribute(ByRef xmlDocument As MSXML2.DOMDocument40, _
                                  ByVal pstrNodeContext As String, _
                                  ByVal pstrNomeAtributo As String, _
                                  ByVal pvntValorAtributo As Variant) As String

Dim xmlAttrib                            As MSXML2.IXMLDOMAttribute
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrNomeAtributo = vbNullString Then
        'Parâmetro pstrNomeAtributo deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrXMLFilho deve ser diferente de vbNullString"
    ElseIf pstrNodeContext = vbNullString Then
        'Parâmetro pstrNodeContext deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrNodeContext deve ser diferente de vbNullString"
    End If

    Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
    
    Set xmlAttrib = xmlDocument.createAttribute(pstrNomeAtributo)
    
    xmlAttrib.Text = pvntValorAtributo
    
    objDomNodeContext.attributes.setNamedItem xmlAttrib

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

' Adiona uma string XML estruturado a um objeto XML.

Public Sub fgAppendXML(ByRef objMapaNavegacao As MSXML2.DOMDocument40, _
                       ByVal pstrNodeContext As String, _
                       ByVal pstrXMLFilho As String, _
              Optional ByVal pstrNodeRepetName As String = "")
              
Dim xmlFilho                                As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrXMLFilho = vbNullString Then
        'Parâmetro pstrXMLFilho deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrXMLFilho deve ser diferente de vbNullString"
    ElseIf pstrNodeContext = vbNullString Then
        'Parâmetro pstrNodeContext deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrNodeContext deve ser diferente de vbNullString"
    End If
 
    'Leitura do XML Filho
    Set xmlFilho = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlFilho.loadXML(pstrXMLFilho) Then
        '100 - Documento XML Inválido.
        'lngCodigoErroNegocio = 100
        fgErroLoadXML objMapaNavegacao, App.EXEName, "basRBLibrary", "fgAppendXML"
        GoTo ErrorHandler
    End If
       
    'Setar o nivel que deverá entrar o XML Filho
    
    'Para criar um Grupo (Ex.:Grupo_Usuario) dentro de Uma Repeticao (Ex. Repeat_Usuario)
    'Passar o argumento pNomeRepet = Repet_Usuario
    If pstrNodeRepetName <> "" Then
        Set objDomNode = objMapaNavegacao.selectSingleNode("//" & pstrNodeRepetName)
    Else
        Set objDomNode = objMapaNavegacao.documentElement.selectSingleNode("//" & pstrNodeContext)
        'Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
    End If

    'Adicionar XML Filho na Saida
    objDomNode.appendChild xmlFilho.childNodes.Item(0)
    
    Set xmlFilho = Nothing

    Exit Sub

ErrorHandler:
    
    Set xmlFilho = Nothing
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd

End Sub

' Retorna o nível (categoria) do item de caixa passado como parâmetro.

Public Function fgObterNivelItemCaixa(ByVal pstrCodItemCaixa As String) As Integer
Dim udtCodigoItemCaixa                      As udtCodigoItemCaixa
Dim udtCodigoItemCaixaAux                   As udtCodigoItemCaixaAux
    
    If InStr(1, Format(pstrCodItemCaixa, "0,000,000,000,000,000"), "000") <> 0 Then
        fgObterNivelItemCaixa = (InStr(1, Format(pstrCodItemCaixa, "0,000,000,000,000,000"), "000") \ 4)
        If fgObterNivelItemCaixa = 0 Then
            fgObterNivelItemCaixa = 1
        End If
    ElseIf Len(pstrCodItemCaixa) = 0 Then
        fgObterNivelItemCaixa = 1
    Else
        fgObterNivelItemCaixa = 5
    End If
    
End Function

' Remove um node de um objeto XML.

Public Function fgRemoveNode(ByRef xmlDocument As MSXML2.DOMDocument40, _
                             ByVal pstrNodeNameRemove As String) As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrNodeNameRemove = vbNullString Then
        'Parâmetro pstrNodeNameRemove deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrNodeNameRemove deve ser diferente de vbNullString"
    End If

    If xmlDocument.documentElement.selectSingleNode("//" & pstrNodeNameRemove) Is Nothing Then
        Exit Function
    End If
    
    Set objDomNode = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeNameRemove)
    
    xmlDocument.getElementsByTagName(objDomNode.parentNode.nodeName).Item(0).removeChild objDomNode
    
    Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Function

' Cria um CDATA Section em um objeto XML.

Public Function fgCreateCDATASection(ByVal psConteudo As String) As MSXML2.IXMLDOMCDATASection

Dim xmlDoc                                  As MSXML2.DOMDocument40
Dim objnodeCDATA                            As IXMLDOMCDATASection


On Error GoTo ErrorHandler

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objnodeCDATA = xmlDoc.createCDATASection(psConteudo)

    Set fgCreateCDATASection = objnodeCDATA
    
    Set xmlDoc = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlDoc = Nothing

    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Function

' Converte uma data em formato 'YYYYMMDD' em uma string compatível com uma instrução SQL do Oracle.

Public Function fgDtXML_To_Oracle(ByVal pstrYYYYMMDD As String) As String
On Error GoTo ErrorHandler

    If Len(pstrYYYYMMDD) <> 8 Then
        If Trim(pstrYYYYMMDD) = vbNullString Or _
            pstrYYYYMMDD = gstrDataVazia Then
            fgDtXML_To_Oracle = "Null"
            Exit Function
        Else
            'Parâmetro pstrYYYYMMDD deve ser informado com 8 dígito, no formato "YYYYMMDD"
            Err.Raise vbObjectError + 513, , "Parâmetro pstrYYYYMMDD deve ser informado com 8 dígito, no formato YYYYMMDD"
        End If
    End If

    fgDtXML_To_Oracle = "TO_DATE('" & Format(fgDtXML_To_Date(pstrYYYYMMDD), "YYYYMMDD") & "','YYYYMMDD')"

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Converte uma data/hora no formato 'YYYYMMDDHHMMSS' em uma string compatível com uma instrução SQL do Oracle.

Public Function fgDtHrXML_To_Oracle(ByVal pstrYYYYMMDDHHMMSS As String) As String
On Error GoTo ErrorHandler

    If Len(pstrYYYYMMDDHHMMSS) <> 14 Then
        If Trim(pstrYYYYMMDDHHMMSS) = vbNullString Or _
            pstrYYYYMMDDHHMMSS = gstrDataVazia Then
            fgDtHrXML_To_Oracle = "Null"
            Exit Function
        Else
            'Parâmetro pstrYYYYMMDDHHMMSS deve ser informado com 8 dígito, no formato "YYYYMMDDHHMMSS"
            Err.Raise vbObjectError + 513, , "Parâmetro pstrYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato dd/mm/yyyyy HH:MM:SS"
        End If
    End If
         
    fgDtHrXML_To_Oracle = "TO_DATE('" & pstrYYYYMMDDHHMMSS & "','yyyymmddhh24miss')"

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Converte uma string data no formato YYYYMMDD em um objeto Date obedecendo as configurações regionais
' selecionadas pelo usuário.

Public Function fgDtXML_To_Date(ByVal strYYYYMMDD As String) As Date

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long

On Error GoTo ErrorHandler

    If Len(strYYYYMMDD) <> 8 Then
        If Trim(strYYYYMMDD) = "" Then
            fgDtXML_To_Date = gstrDataVazia
            Exit Function
        Else
            'Parâmetro strYYYYMMDD deve ser informado com 8 dígito, no formato "yyyyymmdd
            Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Date", "Parâmetro strYYYYMMDD deve ser informado com 8 dígito, no formato yyyyymmdd"
        End If
    End If

    lngAno = Mid(strYYYYMMDD, 1, 4)
    lngMes = Mid(strYYYYMMDD, 5, 2)
    lngDia = Mid(strYYYYMMDD, 7, 2)

    fgDtXML_To_Date = DateSerial(lngAno, lngMes, lngDia)

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Converte uma string data/hora no formato YYYYMMDDHHMMSS em um objeto Date obedecendo as configurações regionais
' selecionadas pelo usuário.

Public Function fgDtHrStr_To_DateTime(ByVal strYYYYMMDDHHMMSS As String) As Date

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long
Dim intHora                                 As Integer
Dim intMinuto                               As Integer
Dim intSegundo                              As Integer
Dim dthAux                                  As Date

On Error GoTo ErrorHandler

    If Len(strYYYYMMDDHHMMSS) <> 14 Then
        If Trim(strYYYYMMDDHHMMSS) = "" Then
            fgDtHrStr_To_DateTime = gstrDataVazia
            Exit Function
        Else
            'Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss
             Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Date", "Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss"
        End If
    End If

    lngAno = Mid(strYYYYMMDDHHMMSS, 1, 4)
    lngMes = Mid(strYYYYMMDDHHMMSS, 5, 2)
    lngDia = Mid(strYYYYMMDDHHMMSS, 7, 2)

    intHora = Mid(strYYYYMMDDHHMMSS, 9, 2)
    intMinuto = Mid(strYYYYMMDDHHMMSS, 11, 2)
    intSegundo = Mid(strYYYYMMDDHHMMSS, 13, 2)

    dthAux = DateSerial(lngAno, lngMes, lngDia)
    
    dthAux = DateAdd("H", intHora, dthAux)
    dthAux = DateAdd("n", intMinuto, dthAux)
    dthAux = DateAdd("s", intSegundo, dthAux)
        
    fgDtHrStr_To_DateTime = dthAux

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Converte uma string hora no formato HHMMSS em um objeto Date obedecendo as configurações regionais
' selecionadas pelo usuário.

Public Function fgHrStr_To_Time(ByVal strHHMMSS As String) As Date

Dim intHora                                 As Integer
Dim intMinuto                               As Integer
Dim intSegundo                              As Integer
Dim dthAux                                  As Date

On Error GoTo ErrorHandler

    If Len(strHHMMSS) <> 6 Then
        If Trim(strHHMMSS) = "" Then
            fgHrStr_To_Time = gstrDataVazia
            Exit Function
        Else
            'Parâmetro strHHMMSS deve ser informado com 6 dígitos, no formato HHmmss
             Err.Raise vbObjectError + 513, App.EXEName & "-fgHrStr_To_Time", "Parâmetro strHHMMSS deve ser informado com 6 dígitos, no formato HHmmss"
        End If
    End If
    
    intHora = Mid(strHHMMSS, 1, 2)
    intMinuto = Mid(strHHMMSS, 3, 2)
    intSegundo = Mid(strHHMMSS, 5, 2)

    'dthAux = DateSerial(lngAno, lngMes, lngDia)
    
    dthAux = DateAdd("H", intHora, dthAux)
    dthAux = DateAdd("n", intMinuto, dthAux)
    dthAux = DateAdd("s", intSegundo, dthAux)
        
    fgHrStr_To_Time = dthAux

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Converte uma string data no formato YYYYMMDD em uma string formatada obedecendo as configurações regionais
' selecionadas pelo usuário.

Public Function fgDtXML_To_Interface(ByVal strYYYYMMDD As String) As String

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long

On Error GoTo ErrorHandler
    
    If Len(strYYYYMMDD) <> 8 Then
        If Trim(strYYYYMMDD) = "" Then
            Exit Function
        Else
            'Parâmetro strYYYYMMDD deve ser informado com 8 dígito, no formato "yyyyymmdd
            Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Interface", "Parâmetro strYYYYMMDD deve ser informado com 8 dígito, no formato yyyyymmdd"
        End If
    End If

    If Trim(strYYYYMMDD) = "00:00:00" Then
        fgDtXML_To_Interface = vbNullString
        Exit Function
    End If

    lngAno = Mid(strYYYYMMDD, 1, 4)
    lngMes = Mid(strYYYYMMDD, 5, 2)
    lngDia = Mid(strYYYYMMDD, 7, 2)

    fgDtXML_To_Interface = DateSerial(lngAno, lngMes, lngDia)

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Converte uma string data/hora no formato YYYYMMDDHHMMSS em uma string formatada obedecendo as configurações regionais
' selecionadas pelo usuário.

Public Function fgDtHrXML_To_Interface(ByVal strYYYYMMDDHHMMSS As String) As String

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long
Dim intHora                                 As Integer
Dim intMinuto                               As Integer
Dim intSegundo                              As Integer
Dim dthAux                                  As Date

On Error GoTo ErrorHandler

    If Len(strYYYYMMDDHHMMSS) <> 14 Then
        If Trim(strYYYYMMDDHHMMSS) = "" Then
            Exit Function
        Else
            'Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss
             Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Date", "Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss"
        End If
    End If

    lngAno = Mid(strYYYYMMDDHHMMSS, 1, 4)
    lngMes = Mid(strYYYYMMDDHHMMSS, 5, 2)
    lngDia = Mid(strYYYYMMDDHHMMSS, 7, 2)

    intHora = Mid(strYYYYMMDDHHMMSS, 9, 2)
    intMinuto = Mid(strYYYYMMDDHHMMSS, 11, 2)
    intSegundo = Mid(strYYYYMMDDHHMMSS, 13, 2)

    dthAux = DateSerial(lngAno, lngMes, lngDia)
    
    dthAux = DateAdd("H", intHora, dthAux)
    dthAux = DateAdd("n", intMinuto, dthAux)
    dthAux = DateAdd("s", intSegundo, dthAux)
        
    fgDtHrXML_To_Interface = dthAux

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Converte uma string data no formato DD_MM_YYYY em um objeto Date obedecendo as configurações regionais
' selecionadas pelo usuário.

Public Function fgDtString_To_Date(ByVal psDD_MM_YYYY As String) As Date

Dim intDia                                  As Integer
Dim intMes                                  As Integer
Dim intAno                                  As Integer

On Error GoTo ErrorHandler

    If Len(psDD_MM_YYYY) < 10 Then
        'Parâmetro psDD_MM_YYYY deve ser informado com 10 dígito, no formato "dd/mm/yyyyy"
        Err.Raise vbObjectError + 513, , "Parâmetro psDD_MM_YYYY deve ser informado com 10 dígitos, no formato dd/mm/yyyyy"
    End If
     
    intDia = Mid(psDD_MM_YYYY, 1, 2)
    intMes = Mid(psDD_MM_YYYY, 4, 2)
    intAno = Mid(psDD_MM_YYYY, 7, 4)
    
    fgDtString_To_Date = DateSerial(intAno, intMes, intDia)
    
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Converte um objeto Date em uma string no formato data.

Public Function fgDt_To_Xml(ByVal pdtmData As Date) As String

On Error GoTo ErrorHandler

    If Not IsDate(pdtmData) Then
        'Parâmetro pdtmData deve ser do tipo Date
        Err.Raise vbObjectError + 513, , "Parâmetro pdtmData deve ser do tipo 'Date'"
    End If
    
    fgDt_To_Xml = Format(pdtmData, "YYYYMMDD")
    
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Converte um objeto Date em uma string no formato data/hora.

Public Function fgDtHr_To_Xml(ByVal pdtmData As Date) As String

On Error GoTo ErrorHandler

    If Not IsDate(pdtmData) Then
        'Parâmetro pdtmData deve ser do tipo Date
        Err.Raise vbObjectError + 513, , "Parâmetro pdtmData deve ser do tipo 'Date'"
    End If

    fgDtHr_To_Xml = Format(pdtmData, "YYYYMMDDHHMMSS")

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Converte um objeto Date em uma string no formato hora.

Public Function fgHr_To_Xml(ByVal pdtmData As Date) As String

On Error GoTo ErrorHandler

    If Not IsDate(pdtmData) Then
        'Parâmetro pdtmData deve ser do tipo Date
        Err.Raise vbObjectError + 513, , "Parâmetro pdtmData deve ser do tipo 'Date'"
    End If

    fgHr_To_Xml = Format(pdtmData, "HHMMSS")

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Retorna a hora de uma variant passada como parâmetro no formato YYYYMMDDHHMMSS.

Public Function fgDtHrXml_To_Time(ByVal vntYYYYMMDDHHMMSS As Variant) As String
    
    If Trim$(vntYYYYMMDDHHMMSS) = vbNullString Then Exit Function
    
    vntYYYYMMDDHHMMSS = Right$(vntYYYYMMDDHHMMSS, 6)
    fgDtHrXml_To_Time = Left$(vntYYYYMMDDHHMMSS, 2) & ":" & Mid$(vntYYYYMMDDHHMMSS, 3, 2) & ":" & Right$(vntYYYYMMDDHHMMSS, 2)
    
End Function

' Recebe como parâmetro um valor com o separador decimal vírgula.
' Retorna o valor formatado no padrão do painel de controle.

Public Function fgVlrXml_To_Interface(ByVal pvntValor As Variant, _
                             Optional ByVal pblnDecimal As Boolean = True)
                             
    fgVlrXml_To_Interface = fgVlrXml_To_InterfaceDecimais(pvntValor, IIf(pblnDecimal, 2, 0))
    
End Function

' Recebe como parâmetro um valor com o separador decimal vírgula.
' Retorna o valor formatado no padrão do painel de controle.

Public Function fgVlrXml_To_InterfaceDecimais(ByVal pvntValor As Variant, _
                                    Optional ByVal pintDecimais As Integer = 2) As String

Dim varValorAux                             As Variant
Dim strNegativeSign                         As String
Dim strDecimal                              As String
Dim strMaskAux                              As String


On Error GoTo ErrorHandler

    If pintDecimais = 0 Then
        strMaskAux = gstrMaskInt
    Else
        strMaskAux = Replace(gstrMaskDecN, "CASAS", String(pintDecimais, "0"))
    End If


    If IsEmpty(pvntValor) Then
        fgVlrXml_To_InterfaceDecimais = Format(0, strMaskAux)
        Exit Function
    ElseIf Trim(pvntValor) = "" Then
        fgVlrXml_To_InterfaceDecimais = Format(0, strMaskAux)
        Exit Function
    End If
    
    strDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    
    If strDecimal <> "," Then
        pvntValor = Replace(pvntValor, ",", strDecimal)
    End If
    
    'If pblnDecimal Then
        fgVlrXml_To_InterfaceDecimais = Format$(pvntValor, strMaskAux)
    'End If

    Exit Function
    
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

' Recebe como parâmetro um valor com o separador decimal em qualquer formato.
' Retorna um valor com o separador decimal no formato do Servidor

Public Function fgVlrXml_To_Decimal(ByVal pvntValor As Variant) As Variant

Dim intPOS                                  As Integer
Dim strDecimal                              As String
Dim strThousand                             As String
Dim strNegativeSign                         As String
Dim vntArrayDecimal                         As Variant
Dim vntArrayThousand                        As Variant
Dim vntArrayPonto                           As Variant
Dim vntArrayVirgula                         As Variant
Dim blnValorNegativo                        As Boolean

On Error GoTo ErrorHandler

    If IsEmpty(pvntValor) Then
        fgVlrXml_To_Decimal = CDec(0)
        Exit Function
    ElseIf Trim(pvntValor) = "" Then
        fgVlrXml_To_Decimal = CDec(0)
        Exit Function
    End If

    strDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    strThousand = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sThousand")
    strNegativeSign = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sNegativeSign")
    
    If strNegativeSign = "" Then strNegativeSign = "-"

    vntArrayDecimal = Split(pvntValor, strDecimal)
    vntArrayThousand = Split(pvntValor, strThousand)
    
    vntArrayPonto = Split(pvntValor, ".")
    vntArrayVirgula = Split(pvntValor, ",")
    
    If (UBound(vntArrayDecimal) + UBound(vntArrayThousand)) > 1 Then
        'Parametro pvntValor não pode estar formatado
'        GoTo ErrorHandler
        pvntValor = Replace(pvntValor, strThousand, vbNullString)
    End If
        
    blnValorNegativo = False

    If InStr(pvntValor, strNegativeSign) > 0 Then
        blnValorNegativo = True
    End If

    If blnValorNegativo Then
        pvntValor = Replace(pvntValor, strNegativeSign, vbNullString)
    End If

    intPOS = InStr(1, pvntValor, strDecimal)

    If intPOS > 0 Then
        fgVlrXml_To_Decimal = CDec(pvntValor)
    Else
        'Verificar se número tem separador decimal diferente da máquina
        If UBound(vntArrayPonto) Then
            fgVlrXml_To_Decimal = CDec(Replace(pvntValor, ".", strDecimal))
        ElseIf UBound(vntArrayVirgula) Then
            fgVlrXml_To_Decimal = CDec(Replace(pvntValor, ",", strDecimal))
        Else
            fgVlrXml_To_Decimal = CDec(pvntValor)
        End If
        
    End If
    
    If blnValorNegativo Then
        fgVlrXml_To_Decimal = fgVlrXml_To_Decimal * -1
    End If

    Exit Function
    
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

' Recebe como parâmetro um valor com o separador decimal no formato corrente.
' Retorna um valor com o separador decimal no formato do Servidor (Vírgula)

Public Function fgVlr_To_Xml(ByVal pvntValor As Variant) As String

Dim strValor                                As String
Dim strDecimal                              As String
Dim strThousand                             As String
Dim strNegativeSign                         As String

On Error GoTo ErrorHandler
    
    If IsEmpty(pvntValor) Then
        fgVlr_To_Xml = CDec(0)
        Exit Function
    ElseIf Trim(pvntValor) = "" Then
        fgVlr_To_Xml = CDec(0)
        Exit Function
    End If
    
    strValor = pvntValor
    
    'Ver Valor negativo
    '.00; ""; empty
    
    strDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    strThousand = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sThousand")
    strNegativeSign = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sNegativeSign")

    strValor = Replace(strValor, strThousand, vbNullString)
    strValor = Replace(strValor, strDecimal, ",")
    
    fgVlr_To_Xml = strValor

    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

' Retorna data e hora do servidor.

Public Function fgDataHoraServidor(ByVal pFormato As enumFormatoDataHoraAux) As Date

Dim objMIU                  As Object
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
Dim strDataHoraServidor     As String

On Error GoTo ErrorHandler
    
    If App.EXEName = "A6" Or _
       App.EXEName = "A7" Or _
       App.EXEName = "A8" Then
    
        #If EnableSoap = 1 Then
            Set objMIU = fgCriarObjetoMIU(App.EXEName & "Miu.clsMIU")
        #Else
            Set objMIU = CreateObject(App.EXEName & "Miu.clsMIU")
        #End If

        strDataHoraServidor = objMIU.DataHoraServidor(pFormato, _
                                                      vntCodErro, _
                                                      vntMensagemErro)
                                                     
        Select Case pFormato
            Case enumFormatoDataHoraAux.DataAux
                fgDataHoraServidor = CDate(Format(strDataHoraServidor, "DD/MM/YYYY"))
            Case enumFormatoDataHoraAux.HoraAux
                fgDataHoraServidor = CDate(Format(strDataHoraServidor, "HH:mm:ss"))
            Case enumFormatoDataHoraAux.DataHoraAux
                fgDataHoraServidor = CDate(Format(strDataHoraServidor, "DD/MM/YYYY HH:mm:ss"))
        End Select
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objMIU = Nothing
    
    Else
       Select Case pFormato
            Case enumFormatoDataHoraAux.DataAux
                fgDataHoraServidor = Date
            Case enumFormatoDataHoraAux.HoraAux
                fgDataHoraServidor = Time
            Case enumFormatoDataHoraAux.DataHoraAux
                fgDataHoraServidor = Now
        End Select
    End If

Exit Function
ErrorHandler:
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Retorna data e hora do servidor formatados para instrução Oracle.

Public Function fgDataHoraServidor_To_Oracle() As String

Dim objMIU                  As Object
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    If App.EXEName = "A6" Or _
       App.EXEName = "A7" Or _
       App.EXEName = "A8" Then
    
        #If EnableSoap = 1 Then
            Set objMIU = fgCriarObjetoMIU(App.EXEName & "Miu.clsMIU")
        #Else
            Set objMIU = CreateObject(App.EXEName & "Miu.clsMIU")
        #End If
        
        fgDataHoraServidor_To_Oracle = "TO_DATE('" & Format(objMIU.DataHoraServidor(enumFormatoDataHoraAux.DataHoraAux, vntCodErro, vntMensagemErro), "YYYYMMDDHHmmss") & "','yyyymmddHH24miss')"
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objMIU = Nothing
    Else
        fgDataHoraServidor_To_Oracle = "TO_DATE('" & Format(Now, "YYYYMMDDHHmmss") & "','yyyymmddHH24miss')"
    End If

Exit Function
ErrorHandler:
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Retorna data e hora do servidor formatados para instrução Oracle.

Public Function fgDataServidor_To_Oracle() As String

Dim objMIU                  As Object
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    If App.EXEName = "A6" Or _
       App.EXEName = "A7" Or _
       App.EXEName = "A8" Then
    
        #If EnableSoap = 1 Then
            Set objMIU = fgCriarObjetoMIU(App.EXEName & "Miu.clsMIU")
        #Else
            Set objMIU = CreateObject(App.EXEName & "Miu.clsMIU")
        #End If
        
        fgDataServidor_To_Oracle = "TO_DATE('" & Format(objMIU.DataHoraServidor(enumFormatoDataHoraAux.DataAux, vntCodErro, vntMensagemErro), "YYYYMMDD") & "','yyyymmdd')"
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objMIU = Nothing
    Else
        fgDataServidor_To_Oracle = "TO_DATE('" & Format(Now, "YYYYMMDD") & "','yyyymmdd')"
    End If

Exit Function
ErrorHandler:
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Deve ser utilizada no Evento Form_KeyPress
'A propriedade KeyPreview do Form deve ser setada para True em Design Time.
'Caso seja necessário bloquear outro caracter, é só acresentá-lo no case
'** Atenção ** Será impossível digitar estes caracteres nos Fomrs (em qualquer controle).
'Adilson - 27/09/2003

Public Function fgBloqueiaCaracterEspecial(ByVal plngKeyAscii As Long)
    
    Select Case plngKeyAscii
        Case 39, 34
            Beep
            fgBloqueiaCaracterEspecial = 0
        Case Else
            fgBloqueiaCaracterEspecial = plngKeyAscii
    End Select
    
End Function

' Limpa caracteres especiais de uma string passada como parâmetro.

Public Function fgLimpaCaracterEspecial(ByVal pstrTexto As String)

Dim strRetorno                                As String
    
Const CAR_APOSTROFE = 39     ' Caracter '
Const CAR_CASP1 = 96         ' `
Const CAR_CASP2 = 180        ' ´
Const CAR_ASPAS1 = 145       ' ‘
Const CAR_ASPAS2 = 146       ' ’
    
On Error GoTo ErrorHandler

    strRetorno = ""
    strRetorno = Replace(pstrTexto, Chr(CAR_APOSTROFE), "")
    strRetorno = Replace(strRetorno, Chr(CAR_CASP1), "")
    strRetorno = Replace(strRetorno, Chr(CAR_CASP2), "")
    strRetorno = Replace(strRetorno, Chr(CAR_ASPAS1), "")
    strRetorno = Replace(strRetorno, Chr(CAR_ASPAS2), "")
    
    fgLimpaCaracterEspecial = strRetorno

    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Verifica se o conteúdo passado como parâmetro tem características de Hora (time) ou não.

Public Function IsTime(ByVal psTime As String) As Boolean

Dim lsTimeSeparator                         As String

    lsTimeSeparator = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sTime")
    
    If InStr(1, psTime, lsTimeSeparator) > 0 Then
        If Trim(psTime) = Format(psTime, "HH:MM") Then
            IsTime = True
        End If
    End If

End Function

' Obtém código de erro de negócio a partir de XML de erro.

Public Function fgObterCodigoDeErroDeNegocioXMLErro(ByVal pstrXMLErro As String) As Long

Dim xmlErro                                 As MSXML2.DOMDocument40
Dim lngErrorNumber                          As Long

    lngErrorNumber = Err.Number

On Error GoTo ErrorHandler

    If Len(pstrXMLErro) = 0 Then Exit Function
    
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlErro.loadXML(pstrXMLErro) Then
        If lngErrorNumber <> 0 Then
            fgObterCodigoDeErroDeNegocioXMLErro = lngErrorNumber
            Exit Function
        Else
            fgErroLoadXML xmlErro, "basRBLibrary", "", "fgObterCodigoDeErroDeNegocioXMLErro"
        End If
    End If
    
    fgObterCodigoDeErroDeNegocioXMLErro = xmlErro.documentElement.selectSingleNode("Grupo_ErrorInfo/Number").Text
    
    Exit Function
    
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

' Função genérica para tratamento de erro no Load de uma string XML.

Public Function fgErroLoadXML(ByRef pxmlDocument As MSXML2.DOMDocument40, _
                              ByVal pstrComponente As String, _
                              ByVal pstrClasse As String, _
                              ByVal pstrMetodo As String)


    Err.Raise pxmlDocument.parseError.errorCode, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, pxmlDocument.parseError.reason

End Function

' Completa uma string com determinado caracter, no tamanho especificado.

Public Function fgCompletaString(ByVal pstrTexto As String, pstrCaractere As String, pintNumCasas As Integer, Optional ByVal pblnEsquerda As Boolean) As String

    If IsNull(pblnEsquerda) Then
        fgCompletaString = Trim(pstrTexto) + String(pintNumCasas - Len(Trim(pstrTexto)), pstrCaractere)
    Else
        If pblnEsquerda Then
            fgCompletaString = String(pintNumCasas - Len(Trim(pstrTexto)), pstrCaractere) + Trim(pstrTexto)
        Else
            fgCompletaString = Trim(pstrTexto) + String(pintNumCasas - Len(Trim(pstrTexto)), pstrCaractere)
        End If
    End If

End Function

'Função para verificar o tempo de execução de uma rotina ou processo, devolvendo o resultado em Nanossegundos

Public Function fgTempoDecorrido(Optional ByVal plngTempoAnterior As Long = 0) As Long

    If plngTempoAnterior > 0 Then
        fgTempoDecorrido = GetTickCount - plngTempoAnterior
    Else
        fgTempoDecorrido = GetTickCount
    End If
    
End Function

' Converte e desconverte strings para a Base 64.

Public Function fgBase64Decode(ByVal pstrBase64EncodeString As String) As String

Dim vntFileIn()                              As Byte
Dim vntOut()                                 As Byte
Dim vntinp(3)                                As Byte
Dim lngTemp                                  As Long
Dim lngCont                                  As Long
Dim lngLenght                                As Long
Dim bytRemaining                             As Byte
Dim lngBytesOut                              As Long
Dim lngTamanho                               As Long
Dim vntBase64Tab(63)                         As Byte
Dim vntDecodeTable(233)                      As Byte

On Error GoTo ErrorHandler
    
    If Trim(pstrBase64EncodeString) = vbNullString Then Exit Function
    
    flInicializaBase64 vntBase64Tab, vntDecodeTable
    
    lngTamanho = Len(pstrBase64EncodeString)
    
    vntFileIn = StrConv(pstrBase64EncodeString, vbFromUnicode)
    
    If vntFileIn(UBound(vntFileIn)) = 61 Then
        bytRemaining = 1
        If vntFileIn(UBound(vntFileIn) - 1) = 61 Then
            bytRemaining = 2
        End If
    End If
    
    lngLenght = UBound(vntFileIn) + 1 'lngLenght of the string
    
    lngBytesOut = ((lngLenght / 4) * 3) - bytRemaining ' how many bytes will the decoded string have
    
    ReDim vntOut(lngBytesOut - 1)
    
    For lngCont = 0 To lngLenght Step 4
        If lngCont >= lngTamanho Then
            GoTo Fim
        End If
        vntinp(0) = vntDecodeTable(vntFileIn(lngCont))
        vntinp(1) = vntDecodeTable(vntFileIn(lngCont + 1))
        vntinp(2) = vntDecodeTable(vntFileIn(lngCont + 2))
        vntinp(3) = vntDecodeTable(vntFileIn(lngCont + 3))
        If vntinp(3) = 64 Or vntinp(2) = 64 Then
            If vntinp(3) = 64 And Not (vntinp(2) = 64) Then
                vntinp(0) = vntDecodeTable(vntFileIn(lngCont))
                vntinp(1) = vntDecodeTable(vntFileIn(lngCont + 1))
                vntinp(2) = vntDecodeTable(vntFileIn(lngCont + 2))
                '2 bytes out
                vntOut(lngTemp) = (vntinp(0) * 4) Or ((vntinp(1) \ 16) And &H3)
                vntOut(lngTemp + 1) = ((vntinp(1) And &HF) * 16) Or ((vntinp(2) \ 4) And &HF)
                GoTo Fim
            ElseIf vntinp(2) = 64 Then
                vntinp(0) = vntDecodeTable(vntFileIn(lngCont))
                vntinp(1) = vntDecodeTable(vntFileIn(lngCont + 1))
                '1 byte out
                vntOut(lngTemp) = (vntinp(0) * 4) Or ((vntinp(1) \ 16) And &H3)
                GoTo Fim
            End If
        End If
        '3 bytes out
        vntOut(lngTemp) = (vntinp(0) * 4) Or ((vntinp(1) \ 16) And &H3)
        vntOut(lngTemp + 1) = ((vntinp(1) And &HF) * 16) Or ((vntinp(2) \ 4) And &HF)
        vntOut(lngTemp + 2) = ((vntinp(2) And &H3) * 64) Or vntinp(3)
        lngTemp = lngTemp + 3
    Next

Fim:
    
    fgBase64Decode = StrConv(vntOut, vbUnicode)
    
    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Function

' Converte strings para a Base 64.

Public Function fgBase64Encode(ByVal pstrStringToEncode As String) As String

Dim vntFileIn()                              As Byte
Dim vntOut()                                 As Byte
Dim vntbin(2)                                As Byte
Dim lngTemp                                  As Long
Dim lngCont                                  As Long
Dim lngLenght                                As Long
Dim bytRemaining                           As Byte
Dim lngBytesOut                              As Long
Dim vntBase64Tab(63)                         As Byte
Dim vntDecodeTable(233)                      As Byte

On Error GoTo ErrorHandler
    
    If Trim(pstrStringToEncode) = vbNullString Then Exit Function
    
    flInicializaBase64 vntBase64Tab, vntDecodeTable

    vntFileIn = StrConv(pstrStringToEncode, vbFromUnicode)
    
    lngLenght = UBound(vntFileIn) + 1 'lngLenght of the string
    
    bytRemaining = ((lngLenght) Mod 3)
    
    If bytRemaining = 0 Then
        lngBytesOut = ((lngLenght / 3) * 4)  ' how many bytes will the encoded string have
    Else
        lngBytesOut = (((lngLenght + (3 - bytRemaining)) / 3) * 4) ' how many bytes will the encoded string have
    End If
    
    ReDim vntOut(lngBytesOut - 1)
    
    For lngCont = 0 To lngLenght - bytRemaining - 1 Step 3
        '3 bytes in
        vntbin(0) = vntFileIn(lngCont)
        vntbin(1) = vntFileIn(lngCont + 1)
        vntbin(2) = vntFileIn(lngCont + 2)
        '4 bytes out
        vntOut(lngTemp) = vntBase64Tab((vntbin(0) \ 4) And &H3F)
        vntOut(lngTemp + 1) = vntBase64Tab((vntbin(0) And &H3) * 16 Or (vntbin(1) \ 16) And &HF)
        vntOut(lngTemp + 2) = vntBase64Tab((vntbin(1) And &HF) * 4 Or (vntbin(2) \ 64) And &H3)
        vntOut(lngTemp + 3) = vntBase64Tab(vntbin(2) And &H3F)
        lngTemp = lngTemp + 4
    Next
    If bytRemaining = 1 Then ' if there is 1 byte bytRemaining
        'read 1 byte, the second in 0
        vntbin(0) = vntFileIn(UBound(vntFileIn))
        vntbin(1) = 0
        vntOut(UBound(vntOut) - 3) = vntBase64Tab((vntbin(0) \ 4) And &H3F)
        vntOut(UBound(vntOut) - 2) = vntBase64Tab((vntbin(0) And &H3) * 16 Or (vntbin(1) \ 16) And &HF)
        vntOut(UBound(vntOut) - 1) = 61
        vntOut(UBound(vntOut)) = 61
    ElseIf bytRemaining = 2 Then 'if there are 2 bytes bytRemaining
        'read 2 bytes, the third is 0
        vntbin(0) = vntFileIn(UBound(vntFileIn) - 1)
        vntbin(1) = vntFileIn(UBound(vntFileIn))
        vntbin(2) = 0
        vntOut(UBound(vntOut) - 3) = vntBase64Tab((vntbin(0) \ 4) And &H3F)
        vntOut(UBound(vntOut) - 2) = vntBase64Tab((vntbin(0) And &H3) * 16 Or (vntbin(1) \ 16) And &HF)
        vntOut(UBound(vntOut) - 1) = vntBase64Tab((vntbin(1) And &HF) * 4 Or (vntbin(2) \ 64) And &H3)
        vntOut(UBound(vntOut)) = 61
    End If

    fgBase64Encode = StrConv(vntOut, vbUnicode)

    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Function

' Inicializa base 64 para conversão de strings.

Private Sub flInicializaBase64(ByRef pvntBase64Tab() As Byte, _
                               ByRef pvntDecodeTable() As Byte)

Dim lngCont                                  As Long
Dim vntDecodeTable                           As Variant

On Error GoTo ErrorHandler

    'initialize the base64 table
    Erase pvntBase64Tab
    Erase pvntDecodeTable

    vntDecodeTable = Array("255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "62", "255", "255", "255", "63", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", _
                          "255", "255", "255", "64", "255", "255", "255", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                          "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", _
                          "255", "255", "255", "255", "255", "255", "26", "27", "28", "29", "30", "31", "32", "33", "34", _
                          "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", _
                          "51", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", _
                          "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255")
    
    For lngCont = LBound(vntDecodeTable) To UBound(vntDecodeTable)
        pvntDecodeTable(lngCont) = vntDecodeTable(lngCont)
    Next
    
    For lngCont = 65 To 90
        pvntBase64Tab(lngCont - 65) = lngCont
    Next
    
    For lngCont = 97 To 122
        pvntBase64Tab(lngCont - 71) = lngCont
    Next
    
    For lngCont = 0 To 9
        pvntBase64Tab(lngCont + 52) = 48 + lngCont
    Next
    
    pvntBase64Tab(62) = 43
    pvntBase64Tab(63) = 47
    
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Sub

' Função genérica para tratamentos de erro SOAP e HTTP.

Public Function fgTrataErroSoapHttp(ByVal plngErrNumber As Long, _
                                    ByVal pstrErrDescription As String) As String

'-----------------------------------------------------------------------------------
'>>> Erros SOAP
'-----------------------------------------------------------------------------------
'5050    - SOAP - BAD_REQUEST
'5051    - SOAP - ACCESS_DENIED
'5053    - SOAP - NOT_FOUND
'5054    - SOAP - BAD_METHOD
'5055    - SOAP - REQ_TIMEOUT
'5056    - SOAP - CONFLICT
'5058    - SOAP - TOO_LARGE
'5100    - SOAP - SERVER_ERROR
'5101    - SOAP - SRV_NOT_SUPPORTED
'5102    - SOAP - BAD_GATEWAY
'5103    - SOAP - NOT_AVAILABLE
'5104    - SOAP - SRV_TIMEOUT
'5105    - SOAP - VER_NOT_SUPPORTED
'5200    - SOAP - BAD_CONTENT
'5300    - SOAP - CONNECTION_ERROR
'5301    - SOAP - BAD_CERTIFICATE_NAME
'-----------------------------------------------------------------------------------
'>>> Erros HTTP
'-----------------------------------------------------------------------------------
'5400    - HTTP - HTTP_UNSPECIFIED
'5402    - HTTP - HTTP_BAD_REQUEST
'5403    - HTTP - HTTP_BAD_RESPONSE
'5404    - HTTP - HTTP_BAD_URL
'5405    - HTTP - HTTP_DNS_FAILURE
'5406    - HTTP - HTTP_CONNECT_FAILED
'5407    - HTTP - HTTP_SEND_FAILED
'5408    - HTTP - HTTP_RECV_FAILED
'5409    - HTTP - HTTP_HOST_NOT_FOUND
'5415    - HTTP - HTTP_TIMEOUT
'5416    - HTTP - HTTP_CANNOT_USE_PROXY
'5417    - HTTP - HTTP_BAD_CERTIFICATE
'5418    - HTTP - HTTP_BAD_CERT_AUTHORITY
'5419    - HTTP - HTTP_SSL_ERROR
'-----------------------------------------------------------------------------------

Dim strRetorno                              As String

Const PREF_SOAP                             As String = "SOAP - "
Const PREF_HTTP                             As String = "HTTP - "
Const RNG_INICIO_HTTP                       As Integer = 5400
Const RNG_FIM_HTTP                          As Integer = 5419
Const MSG_SUFIXO                            As String = " Tente novamente, ao persistir o erro contate o suporte."

    'Verifica se é um erro proviniente do protocolo SOAP, neste caso,
    'pesquisa a STRING << WSDLREADER >>, que indica que um dos arquivos
    'obrigatórios para a utilização do SOAP não foi encontrado (.WSDL, .WSML)...
    If InStr(1, UCase(pstrErrDescription), "WSDLREADER", vbTextCompare) > 0 Then
        strRetorno = PREF_SOAP & "Serviço IIS indisponível, ou " & _
                                 "arquivos obrigatórios não encontrados no diretório virtual." & _
                                 MSG_SUFIXO
    Else
        '...se não encontrou, então pesquisa pela STRING << HRESULT >>,
        'que indica << retorno >> no padrão de linguagem IDL (Interface Definition Language)...
        If InStr(1, UCase(pstrErrDescription), "HRESULT", vbTextCompare) > 0 Then
            Select Case plngErrNumber
                Case 5050
                    strRetorno = PREF_SOAP & "Solicitação mal formada."
                Case 5051
                    strRetorno = PREF_SOAP & "Acesso Negado ao serviço."
                Case 5053
                    strRetorno = PREF_SOAP & "Serviço IIS indisponível."
                Case 5054
                    strRetorno = PREF_SOAP & "Método mal formado."
                Case 5055
                    strRetorno = PREF_SOAP & "Tempo de resposta para a solicitação excedido (TIMEOUT)."
                Case 5056
                    strRetorno = PREF_SOAP & "Identificação de conflito na chamada do serviço."
                Case 5058
                    strRetorno = PREF_SOAP & "Solicitação muito extensa (XML muito grande)."
                Case 5100
                    strRetorno = PREF_SOAP & "Erro de utilização do servidor."
                Case 5101
                    strRetorno = PREF_SOAP & "Serviço não suportado."
                Case 5102
                    strRetorno = PREF_SOAP & "Conexão com GATEWAY mal formada."
                Case 5103
                    strRetorno = PREF_SOAP & "Serviço indisponível."
                Case 5104
                    strRetorno = PREF_SOAP & "Tempo de resposta do serviço excedido (TIMEOUT)."
                Case 5105
                    strRetorno = PREF_SOAP & "Versão não suportada."
                Case 5200
                    strRetorno = PREF_SOAP & "Conteúdo da mensagem mal formado."
                Case 5300
                    strRetorno = PREF_SOAP & "Erro de conexão com servidor."
                Case 5301
                    strRetorno = PREF_SOAP & "Nome do certificado mal formado."
                Case 5400
                    strRetorno = PREF_HTTP & "Erro não especificado."
                Case 5402
                    strRetorno = PREF_HTTP & "Solicitação mal formada."
                Case 5403
                    strRetorno = PREF_HTTP & "Resposta mal formada."
                Case 5404
                    strRetorno = PREF_HTTP & "Caminho (URL) mal formado."
                Case 5405
                    strRetorno = PREF_HTTP & "Falha de utilização do DNS."
                Case 5406
                    strRetorno = PREF_HTTP & "Falha de conexão."
                Case 5407
                    strRetorno = PREF_HTTP & "Falha no envio da mensagem."
                Case 5408
                    strRetorno = PREF_HTTP & "Falha no recebimento da mensagem."
                Case 5409
                    strRetorno = PREF_HTTP & "Servidor não encontrado."
                Case 5415
                    strRetorno = PREF_HTTP & "Tempo de resposta excedido (TIMEOUT)."
                Case 5416
                    strRetorno = PREF_HTTP & "Impossível a utilização do PROXY."
                Case 5417
                    strRetorno = PREF_HTTP & "Certificado mal formado."
                Case 5418
                    strRetorno = PREF_HTTP & "Autoridade certificadora mal formada."
                Case 5419
                    strRetorno = PREF_HTTP & "Erro de utilização de conexão segura (SSL)."
                Case Else
                    If plngErrNumber >= RNG_INICIO_HTTP And plngErrNumber <= RNG_FIM_HTTP Then
                        strRetorno = PREF_HTTP & "Erro não especificado."
                    Else
                        strRetorno = PREF_SOAP & "Erro não especificado."
                    End If
            End Select
        
             strRetorno = strRetorno & MSG_SUFIXO
        Else
            '...se não encontrou, então é um erro de outra natureza, neste caso,
            '   retorna a própria STRING de erro fornecida
            strRetorno = pstrErrDescription
        End If
    End If

    fgTrataErroSoapHttp = strRetorno

End Function

' Adiciona atributos ao último node de um objeto XML.

Public Function fgAppendAttributeLastNode(ByRef xmlDocument As MSXML2.DOMDocument40, _
                                            ByVal pstrNodeContext As String, _
                                            ByVal pstrNomeAtributo As String, _
                                            ByVal pvntValorAtributo As Variant) As String

Dim xmlAttrib                               As MSXML2.IXMLDOMAttribute
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    If pstrNomeAtributo = vbNullString Then
        'Parâmetro pstrNomeAtributo deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrXMLFilho deve ser diferente de vbNullString"
    ElseIf pstrNodeContext = vbNullString Then
        'Parâmetro pstrNodeContext deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrNodeContext deve ser diferente de vbNullString"
    End If

    'Set objDomNodeContext = xmlDocument.selectSingleNode("//" & pstrNodeContext).childNodes.Item(xmlDocument.selectSingleNode("//" & pstrNodeContext).childNodes.length - 1)
    
    Set objDomNodeContext = xmlDocument.selectNodes("//field").Item(xmlDocument.selectNodes("//field").length - 1)
    'Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
    
    
    Set xmlAttrib = xmlDocument.createAttribute(pstrNomeAtributo)
    
    xmlAttrib.Text = pvntValorAtributo
    
    objDomNodeContext.attributes.setNamedItem xmlAttrib

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

' Formata CNPJ para o formato 00.000.000/0000-00 (apresentacao em tela)

Public Function fgFormataCnpj(ByVal pstrCNPJ As Variant) As String

Dim strRetorno                              As String
    
    If Trim(pstrCNPJ) = "" Then
        fgFormataCnpj = ""
        Exit Function
    End If
    
    strRetorno = pstrCNPJ
    strRetorno = Format(strRetorno, String(14, "0"))
    strRetorno = Format(strRetorno, "@@.@@@.@@@/@@@@-@@")
    
    fgFormataCnpj = strRetorno

End Function

' Faz um Dump na DebugWindow das formatações de coluns do ListView.

Public Function fgDumpFormatacaoListView(lst As Control)

    Dim i, al As String

    Debug.Print ""
    Debug.Print "With " & lst.Name & ".ColumnHeaders"
    For i = 1 To lst.ColumnHeaders.Count
        Select Case lst.ColumnHeaders(i).Alignment
            Case 2
                al = "lvwColumnCenter"
            Case 0
                al = ""
            Case 1
                al = "lvwColumnRight"
        End Select
        Debug.Print "    .Add " & i & ", " & lst.ColumnHeaders(i).Key & ", """ & lst.ColumnHeaders(i).Text & """, " & Int(lst.ColumnHeaders(i).Width) & IIf(al <> "", ", " & al, "")
    Next
    Debug.Print "End With"
    Debug.Print ""

'    With lstMensagem.ColumnHeaders
'        .Clear
'        .Add , , "Conciliar Mensagem", 1600
'        .Add , , "Número Operação", 1600
'        .Add , , "Data da Operação", 1500
'        .Add , , "Data da Liquidação", 1500
'        .Add , , "D/C", 500
'        .Add , , "ID do Título", 1440
'        .Add , , "Data Vencimento", 1500
'        .Add , , "Quantidade", 1440, lvwColumnRight
'        .Add , , "PU", 1440, lvwColumnRight
'        .Add , , "Valor", 1440, lvwColumnRight
'        .Add , , "Veículo Legal", 1440
'        .Add , , "CNPJ Contraparte", 1440
'    End With

End Function

' Simula a função DECODE do Oracle

Public Function fgDECODE(Valor, ParamArray arr())

    
    Dim bTemElse As Boolean
    Dim i As Integer
    
    bTemElse = ((UBound(arr) - LBound(arr)) Mod 2 = 0)

    For i = LBound(arr) To (UBound(arr) + IIf(bTemElse, -1, 0)) Step 2
        If arr(i) = Valor Then
            fgDECODE = arr(i + 1)
            Exit Function
        End If
    Next
    
    If bTemElse Then
        fgDECODE = arr(UBound(arr))
    Else
        fgDECODE = Null
    End If
        
End Function

' Utilizada para dar uma visualizada rapida em um XML, em tempo de desenvolvimento

Public Function fgDumpXML(ByRef xmlAux As Variant)

    Dim strDir As String * 255, i As Integer
    Dim strFullPath, f As Long
    Dim strXML As String
        
    i = GetTempPath(255, strDir)
    strFullPath = Left(strDir, i)
    If Right(strFullPath, 1) <> "\" Then
        strFullPath = strFullPath & "\"
    End If
    strFullPath = strFullPath & "DumpXML.xml"
    
    f = FreeFile
    Open strFullPath For Output Access Write As #f
    
    If TypeName(xmlAux) = "String" Then
        strXML = xmlAux
    ElseIf TypeName(xmlAux) = "DOMDocument" Then
        strXML = xmlAux.xml
    End If
    
    If InStr(1, strXML, "<?xml version=", vbTextCompare) = 0 Then
        Print #f, "<?xml version='1.0' encoding='ISO-8859-1'?>"
    End If
    
    Print #f, strXML
    Close #f
    
    Shell """C:\Arquivos de Programas\Internet Explorer\iexplore.exe"" " & strFullPath, vbMaximizedFocus
'    Shell """C:\Program Files\Internet Explorer\iexplore.exe"" " & strFullPath, vbMaximizedFocus
    
End Function

' Utilizada para dar uma visualizada rapida em um XML, em tempo de desenvolvimento

Public Function fgDumpHTML(ByRef htmlAux As Variant)

    Dim strDir As String * 255, i As Integer
    Dim strFullPath, f As Long
        
    i = GetTempPath(255, strDir)
    strFullPath = Left(strDir, i)
    If Right(strFullPath, 1) <> "\" Then
        strFullPath = strFullPath & "\"
    End If
    strFullPath = strFullPath & "DumpHTML.htm"
    
    f = FreeFile
    Open strFullPath For Output Access Write As #f

    Print #f, htmlAux
    Close #f
    
    Shell """C:\Arquivos de programas\Internet Explorer\iexplore.exe"" " & strFullPath, vbMaximizedFocus
    
End Function

' Executa uma funcao XPath em um XML (m.garcia/g.amaral)
' pstrExpression: uma expressão a ser interpretada. Ex: "sum(//VA_OPER_ATIV)"

Public Function fgFuncaoXPath(pxml As MSXML2.DOMDocument40, pstrExpression As String)

On Error GoTo ErrorHandler
                               
    Dim strXSL                      As String
    Dim xmlXSL          As MSXML2.DOMDocument40
    Dim xmlResultado    As MSXML2.DOMDocument40

    Set xmlXSL = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlResultado = CreateObject("MSXML2.DOMDocument.4.0")

    strXSL = ""
    strXSL = strXSL & "<?xml version=""1.0"" ?>"
    strXSL = strXSL & "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">"
    strXSL = strXSL & "<xsl:template match=""/"">"
    strXSL = strXSL & "<xsl:element name=""resultado"">"
    strXSL = strXSL & "<xsl:value-of select=""" & pstrExpression & """ />"
    strXSL = strXSL & "</xsl:element>"
    strXSL = strXSL & "</xsl:template>"
    strXSL = strXSL & "</xsl:stylesheet>"
    
    xmlXSL.loadXML strXSL
    xmlResultado.loadXML pxml.transformNode(xmlXSL)
    
    fgFuncaoXPath = xmlResultado.selectSingleNode("//resultado").Text
    
    Set xmlXSL = Nothing
    Set xmlResultado = Nothing

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

' Executa uma group function em um XML, agrupando um campo (similar ao GROUP BY do SQL)
' pstrExpression: uma expressão a ser interpretada. Ex: "sum(//VA_OPER_ATIV)"

Public Function fgXSTLGroupFunction(pxml As MSXML2.DOMDocument40, _
                                    pstrCampo As String, _
                                    pstrExpression As String) As MSXML2.DOMDocument40

'On Error GoTo ErrorHandler
'
'
'    Dim strXSL                      As String
'    Dim xmlXSL          As MSXML2.DOMDocument40
'    Dim xmlResultado    As MSXML2.DOMDocument40
'
'    Set xmlXSL = CreateObject("MSXML2.DOMDocument.4.0")
'    Set xmlResultado = CreateObject("MSXML2.DOMDocument.4.0")
'
'    strXSL = ""
'    strXSL = strXSL & "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
'    strXSL = strXSL & "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">"
'    strXSL = strXSL & "   <xsl:output method=""html"" />"
'    strXSL = strXSL & ""
'    strXSL = strXSL & "   <xsl:key name=""chave"" match=""RelProdPmt"" use=""CPF"" />"
'    strXSL = strXSL & ""
'    strXSL = strXSL & "   <xsl:template match=""/"">"
'    strXSL = strXSL & "      <xsl:element name=""results"">"
'    strXSL = strXSL & "         <xsl:apply-templates mode=""tpl_grp"">"
'    strXSL = strXSL & "            <xsl:sort select=""@group"" />"
'    strXSL = strXSL & "         </xsl:apply-templates>"
'    strXSL = strXSL & "      </xsl:element>"
'    strXSL = strXSL & "   </xsl:template>"
'    strXSL = strXSL & ""
'    strXSL = strXSL & "   <xsl:template match=""RelProdPmts"" mode=""tpl_grp"">"
'    strXSL = strXSL & "      <xsl:for-each select=""RelProdPmt[count(. | key('chave', CPF)[1])=1]"">"
'    strXSL = strXSL & "         <xsl:variable name=""CPF"" select=""CPF"" />"
'    strXSL = strXSL & ""
'    strXSL = strXSL & "         <xsl:element name=""result"">"
'    strXSL = strXSL & "            <xsl:attribute name=""group"">"
'    strXSL = strXSL & "               <xsl:value-of select=""$CPF"" />"
'    strXSL = strXSL & "            </xsl:attribute>"
'    strXSL = strXSL & ""
'    strXSL = strXSL & "            <xsl:attribute name=""value"">"
'    strXSL = strXSL & "               <xsl:value-of select=""sum(//ValorOperacao[../CPF=$CPF])"" />"
'    strXSL = strXSL & "            </xsl:attribute>"
'    strXSL = strXSL & "         </xsl:element>"
'    strXSL = strXSL & "      </xsl:for-each>"
'    strXSL = strXSL & "   </xsl:template>"
'    strXSL = strXSL & "</xsl:stylesheet>"
'
'    xmlXSL.loadXML strXSL
'    xmlResultado.loadXML pxml.transformNode(xmlXSL)
'
'    fgFuncaoXPath = xmlResultado.selectSingleNode("//resultado").Text
'
'    Set xmlXSL = Nothing
'    Set xmlResultado = Nothing
'
'Exit Function
'ErrorHandler:
'    Err.Raise Err.Number, Err.Description, Err.Description
End Function

' Faz um 'selectNodes' em um XML, e retorna os XML´s dos nós encontrados.

Public Function fgSelectNodesTexto(xml As MSXML2.DOMDocument40, pstrContext As String) As String

    Dim n                                   As MSXML2.IXMLDOMNode
    Dim strRetorno                          As String
    
    For Each n In xml.selectNodes(pstrContext)
        strRetorno = strRetorno & n.xml
    Next
    fgSelectNodesTexto = strRetorno
    
End Function

' Verifica se valor está na lista

Public Function fgIN(ByVal Valor, ParamArray Lista()) As Boolean
    
    Dim i
    
    fgIN = False
    If IsMissing(Lista) Then
        Exit Function
    End If
    
    For i = LBound(Lista) To UBound(Lista)
        If Valor = Lista(i) Then
            fgIN = True
            Exit For
        End If
    Next
    
End Function

' Classifica o conteúdo de uma array.

Public Function fgSortArray(ByRef DArray(), Element As Integer)
    Dim gap As Integer, doneflag As Integer, SwapArray()
    Dim Index As Integer, acol As Long, cnt As Long
    ReDim SwapArray(2, UBound(DArray, 1), UBound(DArray, 2))
    'Gap is half the records
    gap = Int(UBound(DArray, 2) / 2)
    Do While gap >= 1
        Do
            doneflag = 1
            For Index = 0 To (UBound(DArray, 2) - (gap + 1))
                'Compare 1st 1/2 to 2nd 1/2
                If DArray(Element, Index) > DArray(Element, (Index + gap)) Then
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        SwapArray(0, acol, Index) = DArray(acol, Index)
                        SwapArray(1, acol, Index) = DArray(acol, Index + gap)
                    Next
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        DArray(acol, Index) = SwapArray(1, acol, Index)
                        DArray(acol, Index + gap) = SwapArray(0, acol, Index)
                    Next
                    cnt = cnt + 1
                    doneflag = 0
                End If
            Next
        Loop Until doneflag = 1
        gap = Int(gap / 2)
    Loop
End Function

' Mostra o conteúdo de um array

Public Function fgDumpArray(a(), Optional ByVal ColSize As Integer = 20)
    Dim i, j, k, n
    Dim nArq, sFormato As String, sAux As String
    Const Arquivo = "c:\array.txt"

    'se está vazio, não faz nada
    On Error Resume Next
    i = UBound(a)
    If Err.Number > 0 Then
        Exit Function
    End If
    On Error GoTo 0

    'se é um array com 1 dimensão, transforma em duas dimensoes
    On Error Resume Next
    i = UBound(a, 2)
    If Err.Number > 0 Then
        ReDim Preserve a(0, LBound(a) To UBound(a))
    End If
    On Error GoTo 0

    nArq = FreeFile
    Open Arquivo For Output Access Write As #nArq

    sFormato = String(ColSize, "@")
    
    'print no cabeçalho
    sAux = "     |"
    For i = LBound(a, 2) To UBound(a, 2)
        sAux = sAux & Format(i, sFormato & "!") & "|"
    Next
    Print #nArq, sAux
    sAux = "     |"
    For i = LBound(a, 2) To UBound(a, 2)
        sAux = sAux & String(ColSize, "-") & "|"
    Next
    Print #nArq, sAux
    
    'print do corpo
    For i = LBound(a) To UBound(a)
        sAux = Format(i, "@@@@@") & "|"
        For j = LBound(a, 2) To UBound(a, 2)
            If IsNull(a(i, j)) Then
                sAux = sAux & Mid(Format("<null>", sFormato & "!"), 1, ColSize) & "|"
            ElseIf IsEmpty(a(i, j)) Then
                sAux = sAux & Mid(Format("<empty>", sFormato & "!"), 1, ColSize) & "|"
            ElseIf IsNumeric(a(i, j)) And Not IsEmpty(a(i, j)) Then
                sAux = sAux & Mid(Format(a(i, j), sFormato), 1, ColSize) & "|"
            Else
                sAux = sAux & Mid(Format(a(i, j), sFormato & "!"), 1, ColSize) & "|"
            End If
        Next
        Print #nArq, sAux
    Next
    
    'mostra na tela
    Close nArq
    i = Shell("notepad " & Arquivo, vbMaximizedFocus)
    
End Function

' Se NULL, retorna o segundo parametro.

Public Function fgNVL(pvntValor, pvntIF)
    If IsNull(pvntValor) Then
        fgNVL = pvntIF
    Else
        fgNVL = pvntValor
    End If
End Function

' Se EMPTY, retorna o segundo parametro.

Public Function fgEVL(pvntValor, pvntIF)
    If IsEmpty(pvntValor) Then
        fgEVL = pvntIF
    Else
        fgEVL = pvntValor
    End If
End Function

' Se ZERO, retorna o segundo parametro.

Public Function fgZVL(pvntValor, pvntIF)
    If pvntValor = 0 Then
        fgZVL = pvntIF
    Else
        fgZVL = pvntValor
    End If
End Function

' Retorna o maior item do array.

Public Function fgMax(ParamArray arr())

    Dim lb As Long, ub As Long, i As Long
    Dim result As Variant
    
    lb = LBound(arr)
    ub = UBound(arr)
    
    If ub < lb Then
        Exit Function
    End If
    
    result = arr(lb)
    For i = lb + 1 To ub
        If arr(i) > result Then
            result = arr(i)
        End If
    Next
    
    fgMax = result

End Function

' Tenta acessar um node de objeto XML (prevenção contra Object variable with block no set.).

Public Function fgSelectSingleNode(ByRef xmlDocument As MSXML2.IXMLDOMNode, _
                                   ByVal pstrNodeContext As String) As IXMLDOMNode

    On Error GoTo Error_Handler
    
    Dim strAux              As String
    
    'Tenta acessar o nó
    strAux = xmlDocument.selectSingleNode(pstrNodeContext).xml
    
    Set fgSelectSingleNode = xmlDocument.selectSingleNode(pstrNodeContext)

Exit Function
Error_Handler:

    If Err.Number = 91 Then
        'Object Variable ...
        Err.Raise Err.Number, "basRBLibrary", "Tag não encontrada: [" & pstrNodeContext & "]. (" & Err.Description & ")"
    Else
        Err.Raise Err.Number, "basRBLibrary", Err.Description & " [" & pstrNodeContext & "]"
    End If

End Function

' Tenta acessar um node de objeto XML (prevenção contra Object variable with block no set.).

Public Function fgSelectSingleNodeText(ByRef xmlDocument As MSXML2.IXMLDOMNode, _
                                   ByVal pstrNodeContext As String) As String

    On Error GoTo Error_Handler
    
    Dim strAux              As String
    
    'Tenta acessar o nó
    fgSelectSingleNodeText = xmlDocument.selectSingleNode(pstrNodeContext).Text

Exit Function
Error_Handler:

    If Err.Number = 91 Then
        'Object Variable ...
        Err.Raise Err.Number, "basRBLibrary", "Tag não encontrada: [" & pstrNodeContext & "]. (" & Err.Description & ")"
    Else
        Err.Raise Err.Number, "basRBLibrary", Err.Description & " [" & pstrNodeContext & "]"
    End If

End Function

' Verifica se as contas passadas como parâmetro pertencem ao mesmo local de liquidação.

Public Function fgContaMesmaCamara(ByVal pstrConta1 As String, _
                                   ByVal pstrConta2 As String, _
                                   ByVal plngCodigoLocalLiquidacao As Long) As String

On Error GoTo ErrorHandler

'71.90   GARANTIA DAS CAMARAS LDL - CBLC
'72.90   GARANTIA DAS CAMARAS LDL - BMF AT
'73.90   GARANTIA DAS CAMARAS LDL - BMF DR
'74.90   GARANTIA DAS CAMARAS LDL - BMF CB
'75.90   GARANTIA DAS CAMARAS LDL - CETIP
'76.90   GARANTIA DAS CAMARAS LDL - CIP
'77.90   GARANTIA DAS CAMARAS LDL - TECBAN
'81.90   NEGOCIACAO CAMARAS LDL - CBLC
'82.90   DEPOSITO DAS CAMARAS LDL - BMF AT
    
'    pstrConta1 = fgCompletaString(Right$(pstrConta1, 9), "0", 9, True)
'    pstrConta2 = fgCompletaString(Right$(pstrConta2, 9), "0", 9, True)
'
'    Select Case plngCodigoLocalLiquidacao
'
'        Case enumLocalLiquidacao.BMA
'
'        Case enumLocalLiquidacao.BMC
'
'        Case enumLocalLiquidacao.BMD
'
'        Case enumLocalLiquidacao.CETIP
'
'        Case enumLocalLiquidacao.CIP
'
'        Case enumLocalLiquidacao.CLBCAcoes, _
'             enumLocalLiquidacao.CLBCTpPriv, _
'             enumLocalLiquidacao.CLBCTPub
'
'    End Select
'

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

' Transforma uma Data do Oracle para uma expressão para ser usada com o operador BETWEEN (gus)
' Serve para evitar a expressão WHERE TRUNC(A.CAMPO_DATA) = TO_DATE(...)
' pois esta construção faz com que o Oracle não utilize um índice no plano de acesso.
' A primeira data é truncada, e a segunda é truncada, depois acrescenta-se um dia, e subtrai 1 segundo.

Public Function fgDtOracle_To_OracleBetween(ByVal pstrData As String) As String
On Error GoTo ErrorHandler

    fgDtOracle_To_OracleBetween = " TRUNC(" & pstrData & ") AND (TRUNC(" & pstrData & " + 1) - 1/86400)"
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub fgAberturaTreeViewSet(Node As Object, xml As MSXML2.DOMDocument40)

    On Error GoTo ErrorHandler
        
    
    Dim strNodeName                     As String
    
    strNodeName = "n" & Node.Index
    
    If xml Is Nothing Then
        Set xml = CreateObject("MSXML2.DOMDocument.4.0")
    End If
    If xml.xml = vbNullString Then
        fgAppendNode xml, "", "ITEMS", ""
    End If
    
    If xml.selectNodes("//" & strNodeName).length = 0 Then
        fgAppendNode xml, "ITEMS", strNodeName, ""
    End If
    
    xml.selectSingleNode("//" & strNodeName).Text = Node.Expanded
    
Exit Sub
ErrorHandler:
    
    Err.Raise Err.Number, "basRBLibrary - fgAberturaTreeViewSet"
    
End Sub

Public Sub fgAberturaTreeViewRefresh(tvw As Control, xml As MSXML2.DOMDocument40)

    On Error GoTo ErrorHandler

    Dim xmlNode             As IXMLDOMNode
    Dim Node                As Object
    Dim strNodeName         As String
    

    For Each Node In tvw.Nodes
        strNodeName = "n" & Node.Index
    
        If Not xml.selectNodes("//" & strNodeName).length = 0 Then
            Node.Expanded = CBool(xml.selectSingleNode("//" & strNodeName).Text)
        End If
    Next
    
Exit Sub
ErrorHandler:
    
    Err.Raise Err.Number, "basRBLibrary - fgAberturaTreeViewRefresh"

End Sub

Public Function fgLimpaCaracteresCNPJ(ByVal pCNPJ) As String
    fgLimpaCaracteresCNPJ = Replace(Replace(Replace(pCNPJ, ".", ""), "/", ""), "-", "")
End Function

Public Function fgValidaCNPJ_CPF(ByVal pstrCNPJ_CPF As String) As Boolean
    Dim soma As Long
    Dim ind As Integer
    Dim pos As Integer
    Dim resto As Long
    Dim digito As Integer
    Dim dv As String
    Dim x As Integer

    If IsNull(pstrCNPJ_CPF) Then
        fgValidaCNPJ_CPF = False
        Exit Function
    End If

    pstrCNPJ_CPF = Replace(Replace(Replace(pstrCNPJ_CPF, ".", ""), "/", ""), "-", "")

    If Not IsNumeric(pstrCNPJ_CPF) Then
        fgValidaCNPJ_CPF = False
        Exit Function
    End If

    digito = Right(pstrCNPJ_CPF, 2)
    dv = ""
    pstrCNPJ_CPF = Left(pstrCNPJ_CPF, Len(pstrCNPJ_CPF) - 2)

    For x = 1 To 2
        soma = 0
        ind = 2
        For pos = Len(pstrCNPJ_CPF) To 1 Step -1
            soma = soma + (CLng(Mid(pstrCNPJ_CPF, pos, 1)) * ind)
            ind = ind + 1

            If Len(pstrCNPJ_CPF) > 11 And ind > 9 Then
                ind = 2
            End If
        Next

        resto = soma - ((soma \ 11) * 11)

        If resto < 2 Then
            pstrCNPJ_CPF = pstrCNPJ_CPF & "0"
            dv = dv & "0"
        Else
            pstrCNPJ_CPF = pstrCNPJ_CPF & CStr(11 - resto)
            dv = dv & CStr(11 - resto)
        End If
    Next

    If dv = digito Then
        fgValidaCNPJ_CPF = True
    Else
        fgValidaCNPJ_CPF = False
    End If
End Function
Public Function fgAjustaPath(ByVal pstrPath As String) As String

    If Right(pstrPath, 1) <> "\" Then
        pstrPath = pstrPath & "\"
    End If
    fgAjustaPath = pstrPath

End Function

Public Function fgSimpleText(ByVal pstrText As String)
    fgSimpleText = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
                        pstrText, _
                        vbTab, " "), "  ", " "), "  ", " "), "  ", " "), "  ", " "), vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function

Public Function fgLogText(ByVal strLog As String, Optional ByRef dblTempo As Double = 0)

On Error GoTo ErrorTrap

    Dim lngFile As String
    Dim strFile As String
    Dim strTime As String
    Dim strThread  As String
    Dim strApp     As String
    Dim strCtx     As String

    
    ''Object context
    'On Error Resume Next
    'Dim objContext                          As COMSVCSLib.ObjectContext
    'Set objContext = COMSVCSLib.GetObjectContext()
    'strCtx = objContext.ContextInfo.GetTransactionId & "-" & objContext.ContextInfo.GetContextId
    'On Error GoTo ErrorTrap

    If flVariavelAmbiente("SLCC_LOG") = "" Then
        'Só loga se variável de ambiente estiver configurada
        Exit Function
    End If

    strFile = fgAjustaPath(App.Path) & "log"

    On Error Resume Next
    MkDir strFile
    On Error GoTo ErrorTrap

    strFile = strFile & "\SLCC-" & Format(Now, "yyyy-mm-dd") & ".log"
    strApp = Format(App.EXEName, "@@@@@@@@@@!")
    strThread = "[" & strApp & "|" & Replace(Format(Hex(GetCurrentThreadId()), "@@@"), " ", 0) & "|" & strCtx & "]"

    If dblTempo = 0 Then
        strTime = Format(Now, "hh:mm:ss.") & Format((Timer - Int(Timer)) * 100000, "00000")
        dblTempo = Timer
    Else
        strTime = Format(Format(Timer - dblTempo, "0.00000"), "@@@@@@@@@@@@@@")
        'dblTempo = Timer
    End If

    lngFile = FreeFile
    Open strFile For Append Access Write As #lngFile
    Print #lngFile, strTime & " " & strThread & " " & strLog
    Close #lngFile

Exit Function
ErrorTrap:

    On Error Resume Next
    Close #lngFile
    App.LogEvent "SLCC - Erro em fgLogText: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "]"

End Function
Public Function flVariavelAmbiente(ByVal pstrNomeVarAmbiente As String) As String

Dim lngReturnCode                            As Long
Dim strValorVarAmbiente                      As String

    On Error Resume Next
    
    strValorVarAmbiente = String(2000, Chr(0))
        
    lngReturnCode = GetEnvironmentVariable(UCase(pstrNomeVarAmbiente), strValorVarAmbiente, Len(strValorVarAmbiente))
    
    If lngReturnCode <> 0 Then
        flVariavelAmbiente = Mid(strValorVarAmbiente, 1, InStr(1, strValorVarAmbiente, Chr(0)) - 1)
    Else
        flVariavelAmbiente = ""
    End If
            
End Function

Public Function fgGetPrivateProfileString(ByVal pstrSecao As String, _
                                          ByVal pstrChave As String, _
                                          ByVal pstrNomeArquivo As String) As String

Dim lngReturnCode                           As Long
Dim strValorKey                             As String

On Error GoTo ErrHandler
    
    If Dir(pstrNomeArquivo) = vbNullString Then
        Err.Raise 513, , "Arquivo de Configuração não encontrado."
    Else
    
        strValorKey = String(1000, Chr(0))
        lngReturnCode = GetPrivateProfileString(pstrSecao, pstrChave, "", strValorKey, Len(strValorKey), pstrNomeArquivo)
    
        If lngReturnCode <> 0 Then
            fgGetPrivateProfileString = Mid(strValorKey, 1, InStr(1, strValorKey, Chr(0)) - 1)
        End If
    End If
    
    Exit Function
ErrHandler:
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


