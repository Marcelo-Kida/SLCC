VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA8Dominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objetivo       : Publicar dominios comuns para os sistemas A8

Option Explicit

'07/02/2008 - Kida - damanda SGC

Public Enum enumCanalDeVenda
    Nenhum = 0
    SGC = 1
    SGM = 2
End Enum

Public Enum enumTipoAcessoMensagemSPB
    Consultar = 1
    Enviar = 2
End Enum

Public Enum enumSituacaoLote
    Liberado = 1
    ALiberar = 2
    ALiberarManual = 3
    Rejeitado = 4
End Enum

Public Enum enumTipoSituacaoSistema
    OperacaoComAtivo = 1
    MensagemSPB = 2
    LancamentoContaCorrente = 3
End Enum

Public Enum enumSituacaoAceite
    Aceito = 1
    Rejeitado = 2
End Enum

Public Enum enumSituacaoAtivo
    Aceito = 1
    Rejeitado = 2
    ErroInconsistencia = 3
    PendenteD0 = 4
    PendenteD1 = 5
    RejeitadoSELIC = 6
    RejeitadoCETIP = 7
    RejeitadoEmitente = 8
End Enum

Public Enum enumFuncionalidade
    RecebimentoBoleto = 1
    RecebimentoOperacao = 2
    RecebimentoEventosDespesas = 3
    RecebimentoCancelamento = 4
    Confirmacao = 5
    Liberacao = 6
    Conciliacao = 7
    LiberacaoMensagem = 8
    Concordancia = 9
    RejeicaoConcordancia = 10
    LiberacaoRecebimento = 11
    ConcordanciaPrevia = 12
    RejeicaoConcordanciaPrevia = 13
    LiberacaoPrevia = 14
    LiberacaoMensagemSTR = 15
    LiquidarEventos = 16
    RegularizarContingencia = 17
    LiberacaoPagamento = 18
    ConcordanciaBMC0101 = 19
    ConcordanciaBMC0103 = 20
    ConcordanciaLDL0005R2 = 21
    PagamentoContingencia = 22
    Discordancia = 23
End Enum

Public Enum enumFuncaoSistema
    Confirmar = 1
    Conciliar = 2
    Liberar = 3
    Pagar = 4
    Concordar = 5
    PagarContingencia = 6
    PagarSTR = 7
    PagarBACEN = 8
    Discordar = 9
    LberarCancelamentoEspecificacaoCompromissada = 10
    LiberarLiquidacaoLeilaoBMA = 11
    LiberarPagamento = 12
    Receber = 13
    ConfirmarReativacao = 14
    LiberarReativacao = 15
    IntegracaoCC = 16
    LiberarCAM0054 = 17
End Enum

Public Enum enumTipoMensagemBUS
    CompromissadaIda = 3
    TermoLiquidacao = 9
    VinculoDesvinculoTransferencia = 13
    EventosSelic = 15
    DespesasSelic = 17
    Redesconto = 19
    PagamentoRedesconto = 21
    ConversaoRedesconto = 23
    TransferenciaBMA = 30
    EspecificacaoOperacaoBMA = 38
    IntermediacaoOperInternas = 40
    LiquidacaoFisicaOperacaoBMA = 42
    OperacoesCorretoras = 50
    MovimentacaoInstrumentoFinanceiro = 52
    EspecifQtdeCotasFundosInvestimento = 62
    OperacaoDefinitivaCETIP = 64
    OperacaoCompromissadaCETIP = 66
    OperacaoRetornoAntecipacaoCETIP = 68
    RegistroContratoSWAP = 72
    RegistroOperacaoesCETIP = 74
    RegistroContratoTermoCETIP = 76
    ExercicioOpcaoContratoSwapCETIP = 78
    AntecipacaoResgateContratoDerivativoCETIP = 80
    LancamentoPU_CETIP = 82
    EventosJurosCETIP = 86
    DespesasCETIP = 88
    LivreMovimentacao = 90
    RegistroContratoSWAP_CETIP21 = 94
    CargaVeiculoLegal = 110
    RegistroLiquidacaoMultilateralCBLC = 120
    RegistroLiquidacaoBrutaCBLC = 122
    RegistroLiquidacaoEventoCBLC = 124
    RegistroLiquidacaoMultilateralBMF = 126
    RegistroOperacoesBMC = 130
    LiquidacaoMultilateralBMC = 132
    TransferenciasBMC = 134
    DespesasBMC = 136
    RegistroOperacoesRodaDolar = 138
    
    '19/03/07 - Tratamento de EnvioTEDClientes
    EnvioTEDClientes = 150
    ''''''''''''''''''''''''''''''''''''''''''
    
    '08/02/08 Automatizaco Pag Despesas
    EnvioPagDespesas = 154
    ''''''''''''''''''''''''''''''''''''''''''
    
    LancamentoContaCorrenteBG = 156
        
    MensagemSPB = 1000
    MensagemSPBNaoTratada = 1002
    MensagemSPBErro = 1003
    RetornoContaCorrente = 1006
    RetornoContaCorrenteDV = 1006
    RetornoContaCorrenteBG = 1008
    MensagemPZErro = 1009
    MensagemPZR1 = 1010
    MensagemPZR2 = 1011
    MensagemSTR0010R2PZA8 = 1012
    
    'CCR
    ConsultaOperacaoCCR = 158
    EmissaoOperacaoCCR = 160
    NegociacaoOperacaoCCR = 162
    DevolucaoRecolhimentoEstornoReembolsoCCR = 164
    ConsultaLimitesImportacaoCCR = 166
    
    'CAM
    ContratacaoMercadoPrimario = 168
    EdicaoContratacaoMercadoPrimario = 171
    ConfirmacaoEdicaoContratacaoMercadoPrimario = 174
    AlteracaoContrato = 177
    EdicaoAlteracaoContrato = 180
    ConfirmacaoEdicaoAlteracaoContrato = 183
    LiquidacaoMercadoPrimario = 186
    BaixaValorLiquidar = 189
    RestabelecimentoBaixa = 192
    CancelamentoValorLiquidar = 195
    EdicaoCancelamentoValorLiquidar = 198
    ConfirmacaoEdicaoCancelamentoValorLiquidar = 201
    VinculacaoContratos = 204
    AnulacaoEvento = 207
    CorretoraRequisitaClausulasEspecificas = 210
    IFInformaClausulasEspecificas = 213
    ManutencaoCadastroAgenciaCentralizadoraCambio = 216
    CredenciamentoDescredenciamentoDispostoRMCCI = 218
    IncorporacaoContratos = 220
    AceiteRejeicaoIncorporacaoContratos = 223
    ConsultaContratosEmSer = 227
    ConsultaEventosUmDia = 229
    ConsultaDetalhamentoContratoInterbancario = 231
    ConsultaEventosContratoMercadoPrimario = 233
    ConsultaEventosContratoIntermediadoMercadoPrimario = 235
    ConsultaHistoricoIncorporacoes = 237
    ConsultaContratosIncorporacao = 239
    ConsultaCadeiaIncorporacoesContrato = 241
    ConsultaPosicaoCambioMoeda = 243
    AtualizaçãoInclusãoInstrucoesPagamento = 245
    ConsultaInstrucoesPagamento = 247
    RegistroOperacaoInterbancaria = 249
    RegistroOperacaoArbitragem = 253
    ComplementoInformacoesContratacaoInterbancarioViaLeilao = 257
    IFInformaLiquidacaoInterbancaria = 258
    IFCamaraConsultaContratosCambioMercadoInterbancario = 260
    
    'CAM TIR 30/10/2015 - Nick
    IFInformaTIRemContrapartidaaRagadorouRecebedorPaís = 262
    IFInformaTIRemContrapartidaOutraCDE = 264
    IFInformaTIRemContrapartidaOperacaoCambialPropria = 266
    IFRequisitaInclusaoemCadastroCDE = 268
    IFRequisitaAlteracaoCadastroCDE = 270
    IFRequisitaExclusaoCadastroCDE = 272
    IFInformaAnulacaoRegistroTIR = 274
    IFConsultaCDE = 276
    IFConsultaTIRUmDia = 278
    IFConsultaDetalhamentoTIR = 280
    
    

End Enum

'Este Enum inicia a partir de 51, pois as informações de status são comuns
'para OPERAÇÃO e MENSAGEM. A faixa de 01 a 50 está reservada para OPERAÇÃO.
Public Enum enumStatusMensagem
    ManualEmSer = 51
    Concordancia = 52
    EnviadaBUS = 53
    LidaBUS = 54
    ErrosNaoTratados = 55
    GEN0004 = 56
    GEN0009 = 57
    ErroNegocioSPB = 58
    SistemasInternosNZ_PK = 59
    R1 = 60
    R2 = 61
    Aviso = 62
    Informação = 63
    CanceladaOrigem = 64
    Respondida = 65
    Encerrada = 66
    AConciliar = 67
    Conciliada = 68
    MensagemAgendada = 69
    MensagemInconsistente = 70
    MensagemLiquidada = 71
    Liberada = 72
    MensagemEmLancamento = 73
    MensagemPendente = 74
    MensagemExpirada = 75
    MensagemRejeitada = 76
    PagamentoLib = 77
    ConcordanciaBackoffice = 78
    ConcordanciaAdmArea = 79
    DiscordanciaBackoffice = 80
    DiscordanciaAdmArea = 81
    Confirmada = 82
    PagamentoBackoffice = 83
    MensagemCancelada = 84
    ConcordanciaLib = 85
    DiscordanciaLib = 86
    ConcordanciaBackofficePrevia = 87
    LiquidadaFisicamente = 88
    Discordada = 89
    ConcordanciaBackofficeAutomatico = 90
    PagamentoBackofficeAutomatico = 91
    PagamentoLibAutomatico = 92
    ConcordanciaAutomatica = 93
    ConcordanciaLibAuto = 94
    PendenteLibAlcadaAdmArea = 95
    PendenteLibAlcadaAdmGeral = 96
    RecebimentoLib = 97
    Inativa = 98
    ConciliadaAutomatica = 100
    Registrada = 300
    CanceladaCamara = 301
    ConciliadaContingencia = 302
    Devolvida = 303
    Excluida = 307
    Recolhida = 308
    Reembolsada = 309
    PendenteDeAceite = 312
    PendenteDeRegistro = 313
    EnviadaSTR = 315
    Lancada = 316
    Estornada = 317
    EnviadaLegado = 323
    Reativada = 324
    ConfirmadaContingencia = 325
    MensagemLiquidadaAutomatica = 326
End Enum

Public Enum enumMensagemEntradaManual
    TratadaSLCC = 1
    NaoTratadaSLCC = 2
End Enum

Public Enum enumTipoAcao
    AlteracaoHorarioAgendamento = 1
    AlteracaoTipoCompromisso = 2
    CancelamentoSolicitado = 3
    CancelamentoEnviado = 4
    EstornoSolicitado = 5
    EstornoEnviado = 6
    RejeicaoConcordancia = 7            '<- Rejeicao de Concordancia do BackOffice
    RejeicaoDiscordancia = 8            '<- Rejeicao de Discordancia do BackOffice
    EnviadaLDL1002 = 9
    EnviadaSEL1023 = 10
    EnviadaLDL1016 = 11
    RejeicaoConcordanciaAdmArea = 12    '<- Rejeicao de Discordancia do Adm de Area
    RejeicaoDiscordanciaAdmArea = 13    '<- Rejeicao de Discordancia do Adm de Area
    EnviadaLDL0003Concordancia = 14
    EnviadaLDL0003Discordancia = 15
    EnviadoPagamento = 16
    EnviadoPagamentoContingencia = 17
    AjusteValor = 18
    EnviadaLTR0002Concordancia = 19
    EnviadaLTR0002Discordancia = 20
    EnviadoPagamentoSTR = 21
    EnviadoPagamentoBACEN = 22
    EnviadaLDL1006 = 23
    EnviadaLTR0008Concordancia = 24
    EnviadaLTR0008Discordancia = 25
    ConcordanciaEmSer = 26
    ConcordanciaManualEmSer = 27
    RejeicaoConcordanciaEmSer = 28
    RejeicaoConcordanciaManualEmSer = 29
    Liberacao = 30
    LiberacaoAntecipada = 31
    Liquidacao = 32
    EnviadaSTR0004Pagamento = 33
    EnviadaSTR0007 = 34
    EnviadaLDL1003 = 35
    Concordancia = 36
    EnviadaSEL1007 = 37
    CancelamentoPendente = 38
    CancelamentoRejeitado = 39
    CancelamentoSolicitadoComMensagem = 40
    EnviadoRecebimento = 41
    EnviadaBMC0102 = 42
    EnviadaBMC0012 = 43
    RegistroContingencia = 44
    DiscordanciaAdmBO = 45
    LTR0001ComISPBJaExistente = 46
    LTR0007ComISPBJaExistente = 47
    RejeicaoPorDuplicidade = 48
    EnviadaLTR0004Pagamento = 49
    EnviadaLTR0003Pagamento = 50
    EnviadaBMC0001 = 51
    EnviadaCAM0002 = 52
    EnviadaBMC0002 = 53
    EnviadaBMC0003 = 54
    EnviadaConfirmacaoContingencia = 55
    PreviaLiquidada = 56
    EnviadaCAM0054 = 57
    EnviadaCAM0006 = 58
    EnviadaCAM0009 = 59
    EnviadaCAM0007 = 60
    EnviadaCAM0010 = 61
    EnviadaCAM0014 = 62
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// ATENCAO: quando criar um tipo acao novo, colocar tambem a descricao na funcao 'fgDescricaoTipoAcao'
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
End Enum

Public Enum enumTipoFiltroA8
    frmConsultaOperacao = 1
    frmConsultaMensagem = 2
    frmRemessaRejeitada = 3
    frmConfirmacaoOperacao = 4
    frmLiberacaoOperacaoMensagem = 5
    frmReenvioCancelamentoEstornoMsg = 6
    frmSuspenderDisponibilizarLancamentoCC = 7
    frmConsultaContaCorrente = 8
    frmIntegrarCCOnLine = 9
    frmIntegrarCCOnLineEstorno = 10
    frmConsultaVeiculoLegal = 11
    frmConsultaMovimentacao = 12
    frmConsultaMovimentacaoRejeitada = 13
    frmConsultaMovimentacaoCC_HA = 14
    frmConsultaMensagemSisbacen = 15
End Enum

Public Enum enumAcaoMensagemSPB
    CancelamentoSolicitado = 1
    Cancelamento = 2
    Estorno = 3
End Enum

Public Enum enumIndicadorEntradaSaida
    ENTRADA = 1
    Saida = 2
End Enum

Public Enum enumTipoMovimentoPJ
    Previsto = 100
    Realizado = 200
    EstornoPrevisto = 300
    EstornoRealizado = 400
End Enum

Public Enum enumTipoCaixaPJ
    Reserva = 100
    Futuro = 200
End Enum

Public Enum enumTipoDebitoCreditoEstorno
    EstornoDebito = 3
    EstornoCredito = 4
End Enum

Public Enum enumTipoEnvioPJ
    Total = 1
    Parcial = 2
End Enum

Public Enum enumTipoProcessamentoPJ
    OnLine = 1
    Batch = 2
End Enum

Public Enum enumIdentificadorFila
    BUS = 1
    SubReserva = 2
    Liberacao = 3
End Enum

Public Enum enumJustificativa
    SistemaEmContingencia = 1
    RegraWorkflow = 2
    GradeHorario = 3
    CadastroGrupoUsuario = 4
    OperacaoRetroativa = 5
    EntradaManual = 6
    CompromissadaIda = 7
    VeicluloLegal = 8
    CNPJContraparte = 9
    IdentificadorTitulo = 10
    DataVencimentoTitulo = 11
    NumeroComando = 12
    DataOperacao = 13
    DebitoCredito = 14
    DataLiquidacao = 15
    QuantidadeTitulo = 16
    PU = 17
    ValorFinanceiro = 18
    ContaCustodia = 19
    TaxaNegociada = 20
    TitulaCustodiante = 21
    TipoNegociacao = 22
    MensagemOperacaoNaoInformado = 23
    SequenciaOperacao = 24
    OperacaoJaAlterada = 25
    MensagemNaoEncontrada = 26
    MensagemJaAlterada = 27
    OperacaoNaoEncontrada = 28
    IndentificadorContraparteCamara = 29
    CodigoOperacaoCETIP = 30
    DescricaoAtivo = 31
    ModalidadeLiquidacao = 32
    TipoMensagem = 33
    OperacaoNaoLiquidada = 34
    DataRetorno = 35
    PrazoDiasRetorno = 36
    ValorRetorno = 37
    IndentificadorParticipanteCamara = 38
    CodigoPraca = 39
    MoedaEstrangeira = 40
    ValorMoedaEstrangeira = 41
    CodigoContratacaoSISBACEN = 42
    CanalSISBACENCorretora = 43
    ComponenteMBSNaoEncontrado = 44
    ErroAoAcessarFuncaoMBS = 45
    ErroNaoIdentificadoMBS = 46
    NumeroIdentNegociacaoBMC = 47
    SemMensagemCCR0006 = 48
    CodigoAssociacaoCambio = 49
End Enum

Public Enum enumTipoAgendamento
    Operacao = 1
    MENSAGEM = 2
End Enum

Public Enum enumTipoConfirmacao
    Operacao = 1
    MENSAGEM = 2
    Integracao = 3
End Enum

Public Enum enumOpcaoNZPJ
    PJ = 1
    NZ = 2
End Enum

Public Enum enumTipoInformacao
    TipoBackOffice = 1
    GrupoVeiculoLegal = 2
    LocalLiquidacao = 3
    GrupoUsuario = 4
    TodasInformacoes = 5
End Enum

Public Enum enumTipoGradeHorario
    Padrao = 1
    Eventual = 2
End Enum

'Substituiu a rotina de abertura
Public Enum enumCodigoOperacaoSelic
    Resgate = 1012
    Amortizacao = 1010
    Juros = 1060
    RetornoCompromissosTitulosVencemDia = 1056
    LiquidacaoTermosDataLiquidacao = 1052
    LiquidacaoLeilaoVendaPrimario = 1002
    LiquidacaoLeilaoVendaSecundario = 1005
    LiquidacaoLeilaoCompraSecundario = 1006
    Despesas = 1069
    TransferenciaTitulo = 1554
    EstornoComprRetornoDefinido = 3144
    EstornoComprRetornoAberto = 3147
    EstornoTermoIncertaDefinitiva = 3152
    EstornoTermoRetornoDefinido = 3154
    EstornoTermoRetornoAberto = 3157
    EstornoTermoCompromisRetornoDefinido = 4144
    EstornoTermoCompromisRetornoAberto = 4147
    EstornoTermoCertaDefinitiva = 4152
    EstornoTermoLiquCertaRetornoDefinido = 4154
    EstornoTermoLiquCertaRetornoAberto = 4157
End Enum

Public Enum enumStatusIntegracao
    Inicial = 101
    Disponível = 102
    Suspenso = 103
    Cancelado = 104
    Estornado = 105
    IntegradoCC = 106
    Integrado = 107
    ErroCC = 108
    EnviadoCC = 109
    Antecipado = 110
    CanceladoOrigem = 111
    EnviadoEstornoCC = 112
    ErroEstornoCC = 113
    ErroSaldoCC = 114
    EnviadoCCAutomatico = 115
End Enum

Public Enum enumCodigoEmpresa
    Santander = 523
    Meridional = 558
    Bozano = 493
    Banespa = 701
    REAL_ABN = 356
    SUDAMERIS = 215
    BANDEPE = 24
End Enum

Public Enum enumTipoLancamentoIntegracao
    Normal = 1
    Estorno = 2
End Enum

Public Enum enumTipoFuncaoDV
    ConsisteMovimento = 1
    AtualizaMovimento = 2
    ConsisteAtualizaMovimento = 3
    ConsisteAgendamento = 4
    ConsultaSaldo = 5
End Enum

Public Enum enumIntegrarCCContabilidade
    ContaCorrente = 1
    Contabilidade = 2
End Enum

Public Enum enumTipoEnvioRetorno
    Envio = 1
    Retorno = 2
End Enum

Public Enum enumTipoLiquidacao
    SemModalidade = 0
    Multilateral = 1
    Bruta = 2
    BilateralSTR = 3
    BilateralBT = 4
    BrutaSTR = 5
    BrutaBT = 6
    BrutaBTAUT = 7
    Bilateral = 8
End Enum

Public Enum enumTipoMensagemSPB
    MensagemR1 = 1
    MensagemR2 = 2
    MensagemInformacao = 3
    MensagemAviso = 4
    MensagemInativa = 5
End Enum

Public Enum enumTipoFluxo
    TipoFluxo1 = 1 '.  Requisição de Serviço
    TipoFluxo2 = 2 '.  Requisição de Transferência
    TipoFluxo3 = 3 '.  Requisição de transferência com notificação
    TipoFluxo4 = 4 '.  Consulta
    TipoFluxo5 = 5 '.  Informação à IF
    TipoFluxo6 = 6 '.  Informação ao provedor com resposta
    TipoFluxo7 = 7 '.  Aviso à IF
    TipoFluxo8 = 8 '.  Informação da IF para prestador de serviço
    TipoFluxo9 = 9 '.  Requisição de Serviço à IF
    TipoFluxo10 = 10 'Requisição de transferência de arquivos
End Enum

Public Enum enumPerfilAcesso
    Nenhum = 0
    BackOffice = 1
    AdmArea = 2
    AdmGeral = 3
    AdmGeralPrevia = 4
End Enum

Public Enum enumTipoConciliacao
    MultilateralTitulos = 1
    MultilateralFinanceira = 2
    Bilateral = 3
    Bruta = 4
    Registro = 5
End Enum

Public Enum enumAcaoConciliacao
    BOConcordar = 1
    BODiscordar = 2
    AdmAreaLiberar = 3
    AdmAreaRejeitar = 4
    AdmGeralTransferirContingencia = 5
    AdmGeralPagamento = 6
    AdmGeralPagamentoBACEN = 7
    AdmGeralPagamentoSTR = 8
    AdmGeralPagamentoContingencia = 9
    AdmGeralRegularizar = 10
    AdmGeralEnviarConcordancia = 11
    AdmGeralEnviarDiscordancia = 12
    AdmGeralRejeitar = 13
    AdmAreaPagamentoContingencia = 14
    AdmAreaRegularizar = 15
    BOPagamento = 16
    AdmAreaLiberarAntecipada = 17
    AdmAreaLiberarContingencia = 18
    AdmGeralRecebimento = 19
    BOEnviarConcordancia = 20
    BOEnviarConcordanciaContingencia = 21
    AdmAreaPagamento = 22
    BORegularizar = 23
    
    'KIDA - CBLC
    AdmGeralRegularizarRecebimento = 24
End Enum

Public Enum enumSitOpBMA
    OperacaoLiquidada = 29
    ResultadoLiquidoApurado = 20
    OperacaoCancelada = 16
    EspecificacaoTotal = 15
    EspecificacaoDeConta = 14
    PreRegistroDeOperacao = 12
    OperacaoRegistrada = 11
End Enum

Public Enum enumSitResultLiqdBMA
    previo = 1
    final = 2
    definitivo = 3
End Enum

Public Enum enumTIpoFiltroExibicaoMsg
    FiltroCPNJ = 1
    FiltroConta = 2
End Enum

Public Enum enumTipoEventoSEL1611
    OperacaoProcessadaAutomaticamente = 1
    LancamentoRegistradoContraparte = 2
    PreRegistroIntermediacaoLote = 3
End Enum

Public Enum enumTipoContaLancCC
    ContaTributada = 1
    ContaContraparte = 2
    ContaIsenta = 3
End Enum

Public Enum enumTipoNegociacao
    BalcaoMensageria = 1
    NegociacaoEletronica = 2
    BalcaoPCAM = 3
End Enum

