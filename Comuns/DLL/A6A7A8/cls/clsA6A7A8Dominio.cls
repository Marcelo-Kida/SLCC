VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA6A7A8Dominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objetivo       : Publicar dominios comuns para os sistemas A6, A7 e A8

Option Explicit

Public Enum enumAcaoControleInterno
    ErrorHandlerBypass = 1
End Enum

Public Enum enumTipoCaixa
    CaixaSubReserva = 1
    CaixaFuturo = 2
End Enum

Public Enum enumStatusOperacao
    Inicial = 1
    EmSer = 2
    AComplementar = 3
    Concordancia = 4
    ConcordanciaAutomatica = 5
    Liberada = 9
    LiberadaAutomatica = 10
    EmLancamento = 12
    LiquidadaFisicamente = 14
    Liquidada = 15
    Cancelada = 16
    CanceladaOrigem = 17
    Estornada = 18
    Rejeitada = 19
    Pendencia = 20
    Expirada = 21
    AConciliar = 22
    ConciliadaAutomatica = 23
    Inconsistencia = 24
    ManualEmSer = 25
    Conciliada = 26
    BaixadaViaContingencia = 27
    LiquidadaViaContingencia = 28
    LiquidadaConvertida = 29
    Registrada = 30
    RegistradaAutomatica = 31
    ConcordanciaBackoffice = 32
    ConcordanciaAdmArea = 33
    DiscordanciaBackoffice = 34
    DiscordanciaAdmArea = 35
    PagamentoLib = 36
    LiquidadaFisicamenteAutomatica = 37
    Confirmada = 40
    PagamentoBackoffice = 41
    ConcordanciaLib = 42
    DiscordanciaLib = 43
    LiberadaCliente1 = 44
    LiquidadaCliente1 = 45
    ConcordanciaBackofficePrevia = 46
    LiberadaAntecipada = 47
    RejeitadaPiloto = 48
    AConciliarBMA0013 = 49
    ConcordanciaBackofficeBMA0013 = 50
    ConcordanciaBackofficeAutomatico = 200
    PagamentoLiberadoAutomatico = 201
    PagamentoBackofficeAutomatico = 202
    LiberadaTQ = 203    'status transitório, antes de Liquidada, somente para o periodo de integraçao com o sistema TQ (liquidacao CETIP)
    ConcordanciaLibAuto = 204
    RecebimentoLib = 205
    LiquidacaoFutura = 206
    DebitoMoedaNacionalLiquidado = 207
    DebitoMoedaEstrangeiraLiquidado = 208
    Inativa = 209
    RejeitadaLiquidacao = 210
    PendenteLiquidacao = 211
    ConcordanciaBalcao = 212
    ConcordanciaBalcaoAutomatica = 213
    AConciliarRegistro = 214
    AConciliarAceite = 215
    CanceladaCamara = 216
    CancelamentoSolicitado = 217
    ReativacaoSolicitada = 218
    ConcordanciaReativacao = 219
    ConcordanciaReativacaoAutomatica = 220
    LiberadaReativacao = 221
    LiberadaReativacaoAutomatica = 222
    ConcordanciaAceite = 223
    ConcordanciaAceiteAutomatica = 224
    'CCR
    Excluida = 304
    Recolhida = 305
    Reembolsada = 306
    PendenteDeAceite = 310
    PendenteDeRegistro = 311
    
    EnviadaSTR = 314
    
    'SGC
    PendenteEstrategia = 320
    'Sisbacen Interbancario/Arbitragem
    Reativada = 321
    Respondida = 322
End Enum

Public Enum enumTipoOperacaoLQS
    
    'SELIC
    Definitiva = 1
    CompromissadaIda = 2
    CompromissadaVolta = 3
    CompromissadaVoltaConciliacao = 4
    Termo = 5
    TermoDataLiquidacaoCerta = 6
    LeilaoPrimario = 7
    Vinculo = 8
    Desvinculo = 9
    TransferenciaCamaras = 10
    TransferenciaSemFinanceiro = 11
    DespesaSelic = 12
    EventosResgate = 13
    EventosAmortização = 14
    EventosJuros = 15
    RedescontoIntradia = 16
    RedescontoPrazoUmDiaUtil = 17
    RedescontoIntradiaAssociadoAquisicao = 18
    ConversaoRedesconto = 19
    PagamentoRedesconto = 20
    PagamentoRedescontoAssociadoVenda = 21
    RedescontoIntradiaAssocLiberacaoGarantia = 22
    ConsolidacaoRedescontosIntradia = 23
    LeilaoSecundario = 24
    LeilaoSecundarioCompra = 25
    LeilaoInformalDefinitiva = 92
    LeilaoInformalCompromissada = 93
    TermoDataLiquidacaoIncerta = 114
    
    'BMA
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '                               !!!!!!!!!!!! ATENCAO!!!!!!!!!!!!!!!
    '  Quando inserir um novo tipo de operacao para BMA, incluir também na rotina "fgTipoOperacaoBMA"
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    TransferenciaBMA = 26
    DefinitivaCobertaBMA = 27
    DefinitivaDescobertaBMA = 28
    OperacaoTermoCobertaBMA = 29                   '==> Operações a termo coberta(1 - Termo de Papel Decorrido com Correção ;3 - termo de Leilão;5 - Termo de Papel decorrido sem Correção)
    OperacaoTermodesCobertaBMA = 30                '==> Operações a termo descoberta(1 - Termo de Papel Decorrido com Correção ;3 - termo de Leilão;5 - Termo de Papel decorrido sem Correção)
    LiquidacaoOperacaoTermoBMA = 31
    LiquidacaoEventosJurosBMA = 32
    OperacaoDefinitivaInternaBMA = 33
    OperacaoTermoInternaBMA = 34
    EspecDefinitivaIntermediacao = 36
    EspecDefinitivaCobertura = 37
    EspecTermoIntermediacao = 38
    EspecTermoCobertura = 39
    DepositoBMA = 40
    RetiradaBMA = 41
    MovimentacaoEntreCamarasBMA = 42
    LiquidacaoFisicaOperacaoBMA = 69
    CancelamentoEspecificacaoBMA = 70
    CompromissadaEspecificaCobertaBMA = 77
    CompromissadaEspecificaDescobertaBMA = 78
    CompromissadaMigracaoIdaBMA = 79
    CompromissadaMigracaoVoltaBMA = 80
    EspecCompromissadaIntermediacao = 81
    EspecCompromissadaCobertura = 82
    CancelamentoEspecificacaoCompromissadaBMA = 83
    CompromissadaEspecificaTermo = 84
    LiquidacaoCompromissadaEspecificaVolta = 85
    LiquidacaoEventosJurosTituloCompro = 86
    OperacaoCompromissadaInternaBMA = 87
    LeilaoVendaPrimarioBMA = 94
    CompromissadaGenericaAVista = 95
    CompromissadaGenericaATermo = 96
    NETEntradaManualMultilateralBMA = 124

    'CETIP
    LiquidacaoComCorretoras = 35
    MovimentacaoInstrumentoFinanceiro = 45
    RetiradaCustodia = 46
    VincDesvInstrumentoFinanceiro = 47
    TransferenciaCustodia = 48
    ResgateFundoInvestimento = 49
    MovimExercicioDireitosDebentures = 50
    ConversaoPermutaValorMobiliario = 51
    EspecifQtdeCotasFundoInvestimento = 52
    OperacaoDefinitivaCETIP = 53
    OperacaoCompromissadaCETIP = 54
    OperacaoAntecipacaoCompromissadaCETIP = 55
    OperacaoRetornoCompromissadaCETIP = 56
    OperacaoRetencaoIR = 57
    RegistroContratoSWAP = 58
    RegDadosComplemContratoSWAP = 59
    RegistroContratoTermo = 60
    ExercicioOpcaoContratoSWAP = 61
    AntecipacaoResgateContratoSWAP = 62            '<- antigo AntecipacaoResgateContratoDerivativo
    LanctoPUFatorContratoDerivativo = 63
    OperacaoCessaoContratoDerivativo = 64
    OperAnuenciaCessaoContratoDerivativo = 65
    OperIntermediacaoContratoDerivativo = 66
    EventosJurosSWAP = 67                          '<- ANTIGO EventosJurosCETIP = 67
    DespesasCETIP = 68
    MovInstrumentoFinanceiroConciliacao = 71
    RetiradaCustodiaConciliacao = 72
    TransferCustodiaConciliacao = 73
    OperacaoDefinitivaVendaOrdem = 74
    MovimExercDireitosDebenturesConciliacao = 75
    ConversaoPermutaValorMobiliarioConciliacao = 76
    AplicacaoFundoInvestimentoCETIP = 97
    DepositoFundoInvestimentoCETIP = 98
    DepositoFundoInvestimentoConciliacaoCETIP = 99
    AlteracaoContaCorrente = 100
    AntecipacaoResgateContratoTERMO = 101
    EventosJurosTERMO = 102
    EventosCETIP = 105
    NETEntradaManualMultilateralCETIP = 122
    NETEntradaManualBilateralCETIP = 123
    RegistroContratoSWAPSemOpcaoBarreira = 162      '<- Layout 94
    RegistroContratoSWAPComOpcaoBarreira = 163      '<- Layout 94
    AntecipacaoResgateContratoSWAP_CETIP21 = 164
    LanctoPUFatorContratoDerivativo_CETIP21 = 165
    MovimentacaoInstrumentoFinanceiroCTP4001 = 171

    'Corretoras
    LiqCorretoraInternaCTributacaoCBLC = 43
    LiqCorretoraInternaSTributacaoCBLC = 44
    LiqCorretoraExternaCTributacaoCBLC = 88
    LiqCorretoraExternaSTributacaoCBLC = 89
    LiqCorretoraInternaCTributacaoBMF = 90
    LiqCorretoraInternaSTributacaoBMF = 91
    LiqCorretoraExternaCTributacaoBMF = 103
    LiqCorretoraExternaSTributacaoBMF = 104
    LiqCorretoraInternaCTributacaoSTR = 107
    LiqCorretoraInternaSTributacaoSTR = 108
    LiqCorretoraExternaCTributacaoSTR = 109
    LiqCorretoraExternaSTributacaoSTR = 110

    'CBLC
    RegistroLiquidacaoMultilateralCBLC = 111
    RegistroLiquidacaoBrutaCBLC = 112
    RegistroLiquidacaoEventoCBLC = 113
    NETEntradaManualMultilateralCBLC = 125

    'BMD
    RegistroLiquidacaoMultilateralBMF = 115
    NETEntradaManualMultilateralBMD = 126

    'BMC
    RegistroOperacoesBMC = 116
    LiquidacaoMultilateralBMC = 117
    TransferenciasBMCDeposito = 118
    DespesasBMC = 119
    TransferenciasBMCRetirada = 120
    RegistroOperacoesRodaDolar = 121
    NETEntradaManualMultilateralBMC = 127
    RegistroOperacaoBMCBalcao = 144
    RegistroOperacaoBMCEletronica = 145

    'TED                                            ***- Tipos de Operações geradas pelo layout 150 -***
    EnvioTEDSTR0006Clientes = 129
    EnvioTEDSTR0007Clientes = 130
    EnvioTEDSTR0008Clientes = 131
    EnvioTEDSTR0009Clientes = 132
    EnvioTEDSTR0025Clientes = 133
    EnvioTEDSTR0034Clientes = 134
    
    EnvioPAG0105Clientes = 135
    EnvioPAG0106Clientes = 136
    EnvioPAG0108Clientes = 137
    EnvioPAG0109Clientes = 138
    EnvioPAG0121Clientes = 139
    EnvioPAG0134Clientes = 140

    EmissaoTEDSTR0007FdosFIDC = 175
    EmissaoTEDSTR0008FdosFIDC = 176
    EmissaoTEDPAG0106FdosFIDC = 177
    EmissaoTEDPAG0108FdosFIDC = 178


    'Automatização Pag Despesas
    EnvioSTR0006PagDespesas = 141
    EnvioSTR0007PagDespesas = 142
    EnvioSTR0008PagDespesas = 143
    EnvioPagDespesasContaCorrente = 146
    EnvioPagDespesasBoleto = 147
    EnvioPagDespesasTributos = 148
    EnvioSTR0006PagDespesasTrib = 149
    EnvioSTR0007PagDespesasTrib = 150
    EnvioSTR0008PagDespesasTrib = 151
    EnvioPagDespesasContaCorrenteTrib = 152
    EnvioPagDespesasBoletoTrib = 153
    EnvioPagDespesasTributosTrib = 154
    EnvioSTR0006PagDespesasIsenta = 155
    EnvioSTR0007PagDespesasIsenta = 156
    EnvioSTR0008PagDespesasIsenta = 157
    EnvioPagDespesasContaCorrenteIsenta = 158
    EnvioPagDespesasBoletoIsenta = 159
    EnvioPagDespesasTributosIsenta = 160
    
    'Lancamento Conta Corrente Operacoes Manuais e Automaticas
    LancamentoContaCorrenteOperacoesManuais = 161
    LancamentoCCCashFlow = 172
    LancamentoCCSwapJuros = 173
    LancamentoCCCashFlowStrikeFixo = 174
    
    'CCR
    ConsultaOperacaoCCR = 166
    EmissaoOperacaoCCR = 167
    NegociacaoOperacaoCCR = 168
    DevolucaoRecolhimentoEstornoReembolsoCCR = 169
    ConsultaLimitesImportacaoCCR = 170
    
    'CAM
    ContratacaoMercadoPrimario = 179
    EdicaoContratacaoMercadoPrimario = 181
    ConfirmacaoEdicaoContratacaoMercadoPrimario = 183
    AlteracaoContrato = 185
    EdicaoAlteracaoContrato = 187
    ConfirmacaoEdicaoAlteracaoContrato = 189
    LiquidacaoMercadoPrimario = 191
    BaixaValorLiquidar = 193
    RestabelecimentoBaixa = 195
    CancelamentoValorLiquidar = 197
    EdicaoCancelamentoValorLiquidar = 199
    ConfirmacaoEdicaoCancelamentoValorLiquidar = 201
    VinculacaoContratos = 203
    AnulacaoEvento = 205
    CorretoraRequisitaClausulasEspecificas = 207
    IFInformaClausulasEspecificas = 209
    ManutencaoCadastroAgenciaCentralizadoraCambio = 211
    CredenciamentoDescredenciamentoDispostoRMCCI = 212
    IncorporacaoContratos = 213
    AceiteRejeicaoIncorporacaoContratos = 215
    ConsultaContratosEmSer = 218
    ConsultaEventosUmDia = 219
    ConsultaDetalhamentoContratoInterbancario = 220
    ConsultaEventosContratoMercadoPrimario = 221
    ConsultaEventosContratoIntermediadoMercadoPrimario = 222
    ConsultaHistoricoIncorporacoes = 223
    ConsultaContratosIncorporacao = 224
    ConsultaCadeiaIncorporacoesContrato = 225
    ConsultaPosicaoCambioMoeda = 226
    AtualizaçãoInclusãoInstrucoesPagamento = 227
    ConsultaInstrucoesPagamento = 228
    InformaContratacaoCamaraSemTelaCega = 230
    InformaConfirmacaoOperacaoCamaraSemTelaCega = 231
    InformaContratacaoInterbancarioSemCamara = 232
    InformaConfirmacaoOperacaoInterbancarioSemCamara = 233
    InformaConfirmacaoContrCamaraTelaCega = 234
    InformaContrArbitParceiroExteriorPaisPropriaIF = 235
    InformaOperacaoArbitragemParceiroPais = 236
    InformaConfirmacaoOperArbitragemParceiroPais = 237
    CAMInformaContratacaoInterbancarioViaLeilao = 238
    InformaLiquidacaoInterbancaria = 239
    ConsultaContratosCambioMercadoInterbancario = 240
    
    'CAM TIR 30/10/2015 - Nick
    InformaTIRemContrapartidaaRagadorouRecebedorPaís = 247
    InformaTIRemContrapartidaOutraCDE = 248
    InformaTIRemContrapartidaOperacaoCambialPropria = 249
    RequisitaInclusaoemCadastroCDE = 250
    RequisitaAlteracaoCadastroCDE = 251
    RequisitaExclusaoCadastroCDE = 252
    InformaAnulacaoRegistroTIR = 253
    ConsultaCDE = 254
    ConsultaTIRUmDia = 255
    ConsultaDetalhamentoTIR = 256
    
    
End Enum

Public Enum enumSistemaNegociacaoBMA
    NegociacaoEletronica = 4
    OperacoesBalcao = 5
End Enum

Public Enum enumTipoAtivoBMA
    Fianca = 1
    TituloPrivadoCETIP = 2
    TituloPrivadoCBLC = 3
    TituloPublicoFederal = 4
    MoedaEstrangeira = 5
    Outros = 99
End Enum

Public Enum enumOperacaoCETIP_CTP0011
    DepositoEmissaoSemFinaceiro = 1
    DepositoEmissaoAplicacaoFinanceiro = 2
    RetiradaCustodia = 11
End Enum

Public Enum enumOperacaoCETIP_CTP0012
    DesistenciaExercicioDireitoVendaEmissor = 5
    DesistenciaExercicioDireitoNaoRepactuacao = 6
    LancamentoFinanceiroResgate = 12
    ExercicioDireitoVendaEmissor = 15
    ExercicioDireitaNaoRepactuacao = 16
End Enum

Public Enum enumTipoInstituicaoCreditadaDebitada
    CorretoraExterna = 1
    AgenteCompensacao = 2
    CorretoraInterna = 3
End Enum

Public Enum enumTipoLeilao
    FormalPrimario = 1
    FormalSecundarioVenda = 2
    FormalSecundarioCompra = 3
    InformalTituloVenda = 4
    InformalTituloCompra = 5
    InformalDinheiroVenda = 6
    InformalDinheiroCompra = 7
    InformalLivreMoviment = 8
    InformalLivreMovtPUemAberto = 9
    InformalPUemAberto = 10
End Enum

Public Enum enumOperacaoCETIP_CTP0013
    DesvinculacaoCaucao = 3
    Caucao = 13
    VinculacaoReservaTecnica = 23
    AntecipacaoDesvinculacaoReservaTecnica = 35
    VinculacaoDesvinculacaoContaMargem = 71
End Enum

Public Enum enumOperacaoCETIP_CTP0017
    RetiradaCustodiaConversaoAcoes = 17
End Enum

Public Enum enumIndicadorDeslocamento
    D0 = 0
    DM1 = 1
    DM2 = 2
End Enum

Public Enum enumOperacaoCETIP_CTP0021
    TransferenciaCustodiaCamaraNegociacao = 21
    TransferenciaCDPParaINSSSemFinanceiro = 24
    MovimentacaoGarantiaCamaras = 73
    TransferenciaGarantiasContaMargemCamara = 75
    TransferenciaCustodiaEntreContasCliente = 77
End Enum

Public Enum enumOperacaoCETIP_CTP0022
    EspecificacaoLancamentoFinancAplic = 22
    EspecificacaoLancamentoFinancResgate = 44
End Enum

Public Enum enumOperacaoCETIP_CTP0052
    DistribuicaoPrimariaValorImobiliario = 7
    ResgateAntecipado = 14
    IntermediacaoEmissaoSistemaCETIP = 51
    CompraVendaDefinitiva = 52
    RegistroRetroativaCompraVendaDefinitiva = 352
End Enum

Public Enum enumIndicadorPeriodicidade
    Diaria = 1
    final = 2
End Enum

Public Enum enumOperacaoCETIP_CTP0054
    CompraVendaCompromissadaRentaPreFix = 54
    CompraVendaCompromissadaRentaPosFix = 57
    CessaoRetrocessao = 81
    RegistroRetroativoCompraVendaComprRentaPreFix = 354
    RegistroRetroativoCompraVendaComprRentaPosFix = 357
    RegistroDocumCompraVendaComprRentaPreFix = 554
End Enum

Public Enum enumOperacaoCETIP_CTP0055
    AntecipacaoCompraVendaCompromissada = 55
    AntecipacaoCessaoRetrocessao = 82
    RegistroRetroativoAntecipacaoCompraVendaCompr = 355
End Enum

Public Enum enumOperacaoCETIP_CTP0056
    RetornoCompraVendaComprRentPreFix = 56
    RetornoCompraVendaComprRentPosFix = 58
End Enum

Public Enum enumOperacaoCETIP_CTP0064
    RetencaoIR = 64
End Enum

Public Enum enumOperacaoCETIP_CTP0100
    EstornoFinalizado = 43
    EstornoDepositoEmissaoSemFinanceiro = 101
    EstornoDepositoEmissaoComFinanceiro = 102
    EstornoDesvinculacaoCalcao = 103
    EstornoDistribuicaoValoresMobiliarios = 107
    EstornoRetiradaCustodia = 111
    EstornoLancamentoFinanceiroResgate = 112
    EstornoCaucao = 113
    EstornoResgateAntecipado = 114
    EstornoLancamentoPUFatorContrato = 115
    EstornoRetiradaCustodiaConversaoAcoes = 117
    EstornoCessaoContrato = 120
    EstornoEspecificacaoLancamentoFinancAplic = 122
    EstornoEspecificaoLancamentoFinancResgate = 144
    EstornoIntermediacaoEmissaoSistemaCETIP = 151
    EstornoCompraVendaDefinitiva = 152
    EstornoCompraVendaComprRentPreFix = 154
    EstornoAntecipacaoCompraVendaCompr = 155
    EstornoCompraVendaComprRentPosFix = 157
    EstornoPuVoltaOperRetnCompVendComprPosFix = 158
    EstornoRetencaoIR = 164
    EstornoVinculacaoDesvinculacaoContaMargem = 171
    EstornoMovimentacaoGarantiaCamara = 173
    EstornoTransferenciaGarantiasContaMargemCamara = 175
    EstornoTransferenciaCustodiaEntreContasCliente = 177
    EstornoPagamentoIntermedContrato = 181
    EstornoRegistroIntermedContratoSwap = 182
    EstornoPostecipadoCompraVendaDefinitiva = 252
    EstornoPostecipadoCompraVendaComprRentPreFix = 254
    EstornoPostecipadoCompraVendaComprRentPosFix = 257
    
    'SF - KIDA - 28/05/2009
    EstornoVinculacaoReservaTecnica = 123
    EstornoAntecipacaoDesvinculacaoReservaTecnica = 135
    
End Enum

Public Enum enumOperacaoCETIP_CTP9001
    RegistroContratoSemModalidade = 1
    RegistroContratoComModalidade = 2
    RegistroRetroativoContratoSemModalidade = 301
    RegistroRetroativoContratoComModalidade = 302
End Enum

Public Enum enumOperacaoCETIP_CTP9002
    RegistroContratoComModalidade = 2
    RegistroRetroativoContratoComModalidade = 302
End Enum

Public Enum enumOperacaoCETIP_CTP9003
    RegistroContratoSemModalidade = 1
    RegistroRetroativoContratoSemModalidade = 301
End Enum

Public Enum enumOperacaoCETIP_CTP9010
    ExercicioOpcaoSWAP = 10
End Enum

Public Enum enumOperacaoCETIP_CTP9014
    ResgateAntecipado = 14
    RegistroRetroativoResgateAntecipado = 314
End Enum

Public Enum enumOperacaoCETIP_CTP9015
    ExercicioDireitoVendaEmissor = 15
End Enum

Public Enum enumOperacaoCETIP_CTP9020
    CessaoContrato = 20
End Enum

Public Enum enumOperacaoCETIP_CTP9021
    ConfirmacaoCessaoContrato = 21
End Enum

Public Enum enumModalidadeLiquidacao
    BilateralSTR = 3
End Enum

Public Enum enumOperacaoCETIP_CTP9082
    PagamentoIntermediacaoContrato = 81
    RegistroIntermediacaoContratoSwap = 82
End Enum

Public Enum enumTipoRequisicaoBMA
    Requisitado = 1
    Voluntario = 2
    Simulacao = 7
End Enum

Public Enum enumTipoTransferenciaLDL
    TrasnfEntreClearings = 1
    Depósito = 2
    Retirada = 3
End Enum

Public Enum enumFormatoDataHora
    Data = 1
    Hora = 2
    DataHora = 3
End Enum

Public Enum enumSistemaSLCC
    BUS = 1
    LQS = 2
    SBR = 3
End Enum

Public Enum enumSLCCOwner
    OwnerA6 = 1
    OwnerA7 = 2
    OwnerA8 = 3
    OwnerA6Coli = 4
End Enum

Public Enum enumTipoBackOffice
    FundosProprios = 1
    Tesouraria = 2
    FundosTerceiros = 3
    Corretoras = 4
    GBM = 5
    GrandesEmpresas = 6
    Empresas = 7
    Comex = 8
    Todos = 9       '<-- Alterado para 9 para possibilitar a inclusão de novos tipos de backoffice em ordem
    Corretora = 10
End Enum

Public Enum enumTipoEntradaSaida
    ENTRADA = 1
    Saida = 2
End Enum

Public Enum enumTipoVinculo
    Vinculo = 1
    Desvinculo = 2
    TransferenciaCamaraLDL = 3
    TransferenciaSemMovFinanceira = 4
    TransferenciaCessaoFiduciaria = 5
End Enum

Public Enum enumTipoRemessaCETIP
    Previa = 1
    Definitiva = 2
End Enum

Public Enum enumTipoRemessa
    Remessa = 1
    Ajuste = 2
End Enum

Public Enum enumTipoMovimento
    Previsto = 1
    RealizadoSolicitado = 2
    RealizadoConfirmado = 3
    EstornoPrevisto = 4
    EstornoRealizadoSolicitado = 5
    EstornoRealizadoConfirmado = 6
    PrevistoCompromIda = 7
    EstornoPrevistoCompromIda = 8
End Enum

Public Enum enumEvento
    EstatisticaCaixaFuturo = 100
    DetalheOperacaoCaixaFuturo = 101
    CaixaSubReservaSaldoFechamento = 102
    MovimentoSubReserva = 103
    InformacoesCadastraisVeiculoLegal = 110
End Enum

Public Enum enumTipoDebitoCredito
    Debito = 1
    Credito = 2
End Enum

Public Enum enumTipoOperacao
    None = 0
    Alterar = 1
    Excluir = 2
    Incluir = 3
End Enum

Public Enum enumTipoSelecao
    MarcarTodas = 98
    DesmarcarTodas = 99
End Enum

Public Enum enumTipoOperacaoBMA
    TermoNaDataLiquidacao = 1
    RetornoCompromissada = 2
End Enum

Public Enum enumTipoNegociacaoBMA
    TermoPapelDecorridoComCorrecao = 1
    Compromissada = 2
    TermoLeilao = 3
    DefinitivaD0 = 4
    TermoPapelDecorridoSemCorrecao = 5
    MigracaoVoltaCamara = 6
    MigracaoIdaCamara = 7
    Emprestimo = 8
    Troca = 9
End Enum

Public Enum enumSubTipoNegociacaoBMA
    EspecificaAVista = 1
    EspecificaATermo = 2
    GenericaAVista = 3
    GenericaATermo = 4
End Enum


Public Enum enumCobertaDescobertaBMA
    Coberta = 1
    Descoberta = 2
End Enum

Public Enum enumLocalLiquidacao
    SSTR = 1
    CIP = 2
    COMPE_NOTURNO = 3
    CONTA_CORRENTE = 4
    CETIP = 5
    CLBCAcoes = 6
    CLBCTpPriv = 7
    CLBCTPub = 8
    SELIC = 13
    BMC = 10
    BMD = 11
    COMPE_DIURNO = 15
    BMA = 17
    PAG = 19
    Compuls = 20
    CCR = 21
    CAM = 22
End Enum

Public Enum enumTipoPagamentoLDL
    PagamentoJuros = 1
    Resgates = 4
    Amortizacoes = 5
End Enum

Public Enum enumTipoMensagemLQS
    Definitiva = 1
    Compromissada = 3
    RetornoCompromissada = 5
    TermoD0 = 7
    TermoDataLiquidacao = 9
    Leilao = 11
    VinculoDesvinculoTransf = 13
    EventosSelic = 15
    DespesasSelic = 17
    Redesconto = 19
    PgtoRedesconto = 21
    ConversaoRedesconto = 23
    AlteracaoDadosContaCorrente = 27
    TransferenciaLDL_BMA = 30
    RegistroOperacaoBMA = 32
    LiquidacaoOperacoesBMA = 34
    LiquidacaoEventosBMA = 36
    EspecificacaoOperacoesBMA = 38
    IntermediacaoOperacoesInternasBMA = 40
    LiquidacaoFisicaOperacaoBMA = 42
    OperacoesComCorretorasCETIP = 50
    MovimentacoesInstFinancCETIP = 52
    MovimentacoesCustodiaCETIP = 54
    ResgateFundoInvestimentoCETIP = 56
    ExercicioDesistenciaCETIP = 58
    ConversaoPermutaValorImobCETIP = 60
    EspecificacaoQuantidadesCotasCETIP = 62
    OperacaoDefinitivaCETIP = 64
    OperacaoCompromissadaCETIP = 66
    OperacaoRetornoAntecipacaoCETIP = 68
    OperacaoRetencaoIRF_CETIP = 70
    RegistroContratoSWAP = 72
    RegistroOperacaoesCETIP = 74
    RegistroContratoTermoCETIP = 76
    ExercicioOpcaoContratoSwapCETIP = 78
    AntecipacaoResgateContratoDerivativoCETIP = 80
    LancamentoPU_CETIP = 82
    MovimentacoesContratoDerivativo = 84
    EventoJurosCETIP = 86
    DespesasCETIP = 88
    LivreMovimentacao = 90
    RegistroContratoSWAPCetip21 = 94        '<- Layout 94
    
    'CBLC
    RegistroLiquidacaoMultilateralCBLC = 120
    RegistroLiquidacaoBrutaCBLC = 122
    RegistroLiquidacaoEventoCBLC = 124

    'BMF
    RegistroLiquidacaoMultilateralBMF = 126

    'BMC
    RegistroOperacoesBMC = 130
    RegistroOperacoesBMCRetorno = 131
    LiquidacaoMultilateralBMC = 132
    TransferenciasBMC = 134
    DespesasBMC = 136
    RegistroOperacoesRodaDolar = 138

    'TED
    EnvioTEDClientes = 150
    
    'Automatizaco Pag Despesas
    EnvioPagDespesas = 154
    
    'LancamentoContaCorrenteBG
    LancamentoContaCorrenteBG = 156
    
    'BACEN
    MensagemNZ = 1000
    MensagemErroNZ = 1003
    MensagemCCDV = 1006
    MensagemCCBG = 1008
    MensagemErroPZ = 1009
    MensagemR1PZ = 1010
    MensagemR2PZ = 1011
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

' #1::Envio TED Clientes
Public Enum enumTEDStrPag
    
    'STR
    strTransfContaClienteIF = 6
    strTransfIFContaCliente = 7
    strTransfContasDiferentesTitula = 8
    strTransfContasMesmaTitula = 9
    strTransfReservasBancDepositoJudi = 25
    strTransfEnvolvContasInvestimento = 34
    
    'PAG
    pagTransfContaClienteIF = 105
    pagTransfIFContaCliente = 106
    pagTransfContasDiferentesTitula = 108
    pagTransfContasMesmaTitula = 109
    pagTransfReservasBancDepositoJudi = 121
    pagTransfEnvolvContasInvestimento = 134

End Enum

Public Enum enumTipoEspecificacao
    Cobertura = 1
    Intermediacao = 2
    Cancelamento = 3
End Enum

Public Enum enumTipoSolicitacao
    Inclusao = 1
    Complementacao = 2
    Cancelamento = 3
    CancelamentoComMensagem = 4
    RetornoLegado = 5 'Utilizado somente internamente não irá vir do BUS
    Alteracao = 6
    LivreMovimentacao = 7
    CancelamentoPorLastro = 8
    Reativacao = 9
    Confirmacao = 10
End Enum

Public Enum enumIndicadorSimNao
    Sim = 1
    Nao = 2
    Ambos = 3
End Enum

Public Enum enumPaginacao
    proximo = 1
    Anterior = 2
End Enum

Public Enum enumFormaLiquidacao
    ContaCorrente = 1
    Contabil = 2
    ContaCorrenteComTributacao = 3
    Str = 4
    Boleto = 5
    Tributos = 6
End Enum

Public Enum enumTipoRedesconto
    RedescIntradia = 1
    RedescComPrazo1Dia = 2
    RedescIntradiaAquisicao = 3
    RedescIntradiaGarantia = 4
    ConsolidacaoRedescIntradia = 5
End Enum

Public Enum enumTipoPagtoRedesconto
    PagamentoRedesconto = 1
    PagamentoRedescontoAssocAVenda = 2
End Enum

Public Enum enumTipoContabil
    Ativo = 1
    Passivo = 2
    Derivativo = 3
End Enum

Public Enum enumISPB
    IspbSANTANDER = 61472676
    IspbBANESPA = 61411633
    IspbBOZZANO = 33517640
    IspbMERIDIONAL = 90400888
    IspbSELIC = 38121
    IspbBMA = 49334758
    IspbBACEN = 281666
    IspbBMC = 60934221
    IspbBMD = 54641030
    IspbCBLC = 60777661
    IspbCETIP = 28719664
    IspbREAL_ABN = 33066408
    IspbSUDAMERIS = 61230165
    IspbBANDEPE = 10866788
End Enum

Public Enum enumTipoCompraVendaCetip
    Compra = 2
    VendaCarteiraPropria = 1
    VendaCarteiraRepasse = 0
End Enum

'TP_CNPT
Public Enum enumTipoContraparte
    Externo = 1
    Interno = 2
    Cliente1 = 3
End Enum


Public Enum enumCodigoErroMBS

    NADA_ENCONTRADO = 1000
    SEQUENCIA_VAZIA = 1001
    FALHA_ABERTURA_ARQUIVO = 1002
    PARAMETRO_NAO_ENCONTRADO = 1003
    ASSOC_CADASTRADA = 1004
    PARAMETROS_ENTRADA_INCONSISTENTE = 1005
    ERRO_CONEXAO = 1006
    SEQUENCIA_DE_TRANSACOES_NAO_CADASTRADA = 1100
    TRANSACAO_NAO_CADASTRADA = 1101
    HIERARQUIA_CADASTRADA = 1102
    TRANSACAO_JA_E_FILHA = 1103
    TRANSACAO_NIVEL_ZERO = 1104
    US_TRANSACAO_AMARRADA_COM_SUBGRUPO_FAVORITOS = 1105
    HIERARQUIA_NAO_CADASTRADA = 1106
    SISTEMA_CADASTRADO = 1107
    TRANSACAO_CADASTRADA = 1108
    USUARIO_CADASTRADO = 1109
    GRUPO_DE_USUARIOS_CADASTRADO = 1110
    USUARIO_NAO_CADASTRADO = 1113

End Enum


Public Enum enumTipoIdentificadorParticipnateCamara

    ContaEmissora = 1
    ContaPropriaBancoMandatario = 2
    ContaCaucao = 3
    ContaAnuenteContrato = 4
    ContaCedenteContrato = 5
    ContaAdquirenteContrato = 6

End Enum


Public Enum enumTipoRentabilidade
    Pre = 1
    pos = 2
End Enum

Public Enum enumFatorGeradorAlerta
    Erro = 1
    SelicExpiradoPorTempo = 2
    SelicPendenteTitulo = 3
    SelicPendenteOperacao = 4
    SelicRejeitadaPeloSTR = 5
    RDCExpiradaporTempo = 6
    RDCPendente = 7
    RDCRejeitada = 8
    RDCCanceladaMensagem = 9
    RDCConversaoMensagem = 10
    RDCRetornoAutomatico = 11
    RemessaLegadoDadosInvalidos = 12
    AvisoDebitoCreditoRB = 13
    LancamentoCCSuspensoD1 = 14
    MensagensBloqueadasGradeHorario = 15
    MensagemLDL0001 = 16
    MensagemLDL0005R2 = 17
    MensagemLDL0006R2 = 18
    MensagemLTR0001 = 19
    MensagemLTR0007 = 20
    MensagemLTR0005R2 = 21
    MensagemBMA0015 = 22
    MensagemBMA0013 = 23
    MensagemBMA0014 = 24
    MensagemCTP1002 = 26
    MensagemLTR0006R2 = 27
    MensagemLDL0009R2 = 28
    MensagemLTR0009R1 = 32
    MensagemLDL1027 = 37
    DivergenciaNaConciliacao = 38
    AusenciaDeMensagemNaConciliacao = 39
    MensagemBMA0003 = 40
    MensagemBMA0002 = 41
    OperacaoNaoConciliadaAutomatica = 42
    LDLRejeitada = 43
    BMARejeitada = 44
    OperacaoNaoPermiteCancelamento = 45
    OperacaoNaoLocalizada = 46
    AusenciaDeOperacaoNaConciliacao = 47
    LTRRejeitada = 48
    LTRPendente = 49
    CTPRejeitada = 50
    CTPPendente = 51
    MensagemLDL0026R1 = 52
    RemessaRejeitadaA6 = 53
    MensagemBMC0101 = 54
    MensagemBMC0012 = 55
    MensagemBMC0005 = 56
    MensagemLDL0003R1 = 57
    MensagemLTR0003R2 = 58
    MensagemBMC0011 = 59
    MensagemNaoLocalizada = 60
    MensagemBMC0013 = 61
    MensagemBMC0002R1 = 62
    MensagemBMC0002E = 63
    MensagemBMC0003R1 = 64
    MensagemBMC0003E = 65
    MensagemBMC0015SemOperacao = 66
    MensagemBMC0015DivergOperacao = 67
    MensagemBMC0005SemOperacao = 68
    MensagemBMC0015NaoLocalizada = 69
    MensagemBMC0005NaoLocalizada = 70
    MensagemSTRE = 71
    MensagemDevolucaoTEDIndevida = 72
End Enum

Public Enum enumTipoArquivoCBLC
    AMDF = 1
    ALCO = 2
End Enum

Public Enum enumCodigoCorretora
    SantanderBrasilSA = 622
    SantanderCCVM = 78
End Enum

'Automatizaco Pag Despesas
Public Enum enumOrigemRecurso
    TransfIsentaTributada = 1
    Tributada = 2
    Isenta = 3
End Enum

Public Enum enumTipoIncondicionalSaldoCC
    Sim = 1
    Nao = 2
End Enum

Public Enum enumTipoUnilateralidade
    NenhumaParte = 1
    QualquerParte = 2
    SomenteCessionario = 3
    SomenteCedente = 4
End Enum

Public Enum enumTipoPagadorResponsavel
    VendedorComprador = 0
    Vendedor = 1
    Comprador = 2
End Enum

Public Enum enumTipoNegociacaoInterbancaria
    SemCamara = 1
    InterbancarioEletronico = 2
    SemTelaCega = 3
End Enum
        
Public Enum enumTipoNegociacaoArbitragem
    ParceiroPais = 1
    ParceiroExteriorPaisIF = 2
End Enum

Public Enum enumTipoNegociacaoCambial
    Interbancaria = 1
    Arbitragem = 2
End Enum
