using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace A8NET.Comum
{
    public static class Comum
    {
        #region <<< Enumeradores >>>
        public enum EnumStatusOperacao
        {
            Inicial = 1,
            EmSer = 2,
            AComplementar = 3,
            Concordancia = 4,
            ConcordanciaAutomatica = 5,
            Liberada = 9,
            LiberadaAutomatica = 10,
            EmLancamento = 12,
            LiquidadaFisicamente = 14,
            Liquidada = 15,
            Cancelada = 16,
            CanceladaOrigem = 17,
            Estornada = 18,
            Rejeitada = 19,
            Pendencia = 20,
            Expirada = 21,
            AConciliar = 22,
            ConciliadaAutomatica = 23,
            Inconsistencia = 24,
            ManualEmSer = 25,
            Conciliada = 26,
            BaixadaViaContingencia = 27,
            LiquidadaViaContingencia = 28,
            LiquidadaConvertida = 29,
            Registrada = 30,
            RegistradaAutomatica = 31,
            ConcordanciaBackoffice = 32,
            ConcordanciaAdmArea = 33,
            DiscordanciaBackoffice = 34,
            DiscordanciaAdmArea = 35,
            PagamentoLib = 36,
            LiquidadaFisicamenteAutomatica = 37,
            Confirmada = 40,
            PagamentoBackoffice = 41,
            ConcordanciaLib = 42,
            DiscordanciaLib = 43,
            LiberadaCliente1 = 44,
            LiquidadaCliente1 = 45,
            ConcordanciaBackofficePrevia = 46,
            LiberadaAntecipada = 47,
            RejeitadaPiloto = 48,
            AConciliarBMA0013 = 49,
            ConcordanciaBackofficeBMA0013 = 50,
            ConcordanciaBackofficeAutomatico = 200,
            PagamentoLiberadoAutomatico = 201,
            PagamentoBackofficeAutomatico = 202,
            LiberadaTQ = 203,//    'status transitório, antes de Liquidada, somente para o periodo de integraçao com o sistema TQ (liquidacao CETIP)
            ConcordanciaLibAuto = 204,
            RecebimentoLib = 205,
            LiquidacaoFutura = 206,
            DebitoMoedaNacionalLiquidado = 207,
            DebitoMoedaEstrangeiraLiquidado = 208,
            Inativa = 209,
            RejeitadaLiquidacao = 210,
            PendenteLiquidacao = 211,
            ConcordanciaBalcao = 212,
            ConcordanciaBalcaoAutomatica = 213,
            AConciliarRegistro = 214,
            AConciliarAceite = 215,
            CanceladaCamara = 216,
            CancelamentoSolicitado = 217,
            ReativacaoSolicitada = 218,
            ConcordanciaReativacao = 219,
            ConcordanciaReativacaoAutomatica = 220,
            LiberadaReativacao = 221,
            LiberadaReativacaoAutomatica = 222,
            ConcordanciaAceite = 223,
            ConcordanciaAceiteAutomatica = 224,

            // CCR
            Excluida = 304,
            Recolhida = 305,
            Reembolsada = 306,
            PendenteDeAceite = 310,
            PendenteDeRegistro = 311,

            // Sisbacen Interbancario/Arbitragem
            Reativada = 321,
            Respondida = 322
        }

        public enum EnumStatusMensagem
        {
            NA = 0,
            ManualEmSer = 51,
            Concordancia = 52,
            EnviadaBUS = 53,
            LidaBUS = 54,
            ErrosNaoTratados = 55,
            GEN0004 = 56,
            GEN0009 = 57,
            ErroNegocioSPB = 58,
            SistemasInternosNZ_PK = 59,
            R1 = 60,
            R2 = 61,
            Aviso = 62,
            Informação = 63,
            CanceladaOrigem = 64,
            Respondida = 65,
            Encerrada = 66,
            AConciliar = 67,
            Conciliada = 68,
            MensagemAgendada = 69,
            MensagemInconsistente = 70,
            MensagemLiquidada = 71,
            Liberada = 72,
            MensagemEmLancamento = 73,
            MensagemPendente = 74,
            MensagemExpirada = 75,
            MensagemRejeitada = 76,
            PagamentoLib = 77,
            ConcordanciaBackoffice = 78,
            ConcordanciaAdmArea = 79,
            DiscordanciaBackoffice = 80,
            DiscordanciaAdmArea = 81,
            Confirmada = 82,
            PagamentoBackoffice = 83,
            MensagemCancelada = 84,
            ConcordanciaLib = 85,
            DiscordanciaLib = 86,
            ConcordanciaBackofficePrevia = 87,
            LiquidadaFisicamente = 88,
            Discordada = 89,
            ConcordanciaBackofficeAutomatico = 90,
            PagamentoBackofficeAutomatico = 91,
            PagamentoLibAutomatico = 92,
            ConcordanciaAutomatica = 93,
            ConcordanciaLibAuto = 94,
            PendenteLibAlcadaAdmArea = 95,
            PendenteLibAlcadaAdmGeral = 96,
            RecebimentoLib = 97,
            Inativa = 98,
            ConciliadaAutomatica = 100,
            Registrada = 300,
            CanceladaCamara = 301,
            ConciliadaContingencia = 302,
            Devolvida = 303,

            // CCR
            Excluida = 307,
            Recolhida = 308,
            Reembolsada = 309,
            PendenteDeAceite = 312,
            PendenteDeRegistro = 313,

            // Sisbacen Interbancario/Arbitragem
            EnviadaLegado = 323,
            Reativada = 324
        }

        public enum EnumTipoFluxo
        {
            TipoFluxo1 = 1,  // Requisição de Serviço
            TipoFluxo2 = 2,  // Requisição de Transferência
            TipoFluxo3 = 3,  // Requisição de transferência com notificação
            TipoFluxo4 = 4,  // Consulta
            TipoFluxo5 = 5,  // Informação à IF
            TipoFluxo6 = 6,  // Informação ao provedor com resposta
            TipoFluxo7 = 7,  // Aviso à IF
            TipoFluxo8 = 8,  // Informação da IF para prestador de serviço
            TipoFluxo9 = 9,  // Requisição de Serviço à IF
            TipoFluxo10 = 10 // Requisição de transferência de arquivos
        }

        public enum EnumInidicador
        {
            Sim = 1,
            Nao = 2
        }

        public enum EnumTipoMensagem
        {
            //CAM 
            ContratacaoMercadoPrimario = 168,
            ContrataçãoMercadoPrimarioR2 = 170,
            EdicaoContratacaoMercadoPrimario = 171,
            EdicaoContratacaoMercadoPrimarioR2 = 173,
            ConfirmacaoEdicaoContratacaoMercadoPrimario = 174,
            ConfirmacaoEdicaoContratacaoMercadoPrimarioR2 = 176,
            AlteracaoContrato = 177,
            AlteracaoContratoR2 = 179,
            EdicaoAlteracaoContrato = 180,
            EdicaoAlteracaoContratoR2 = 182,
            ConfirmacaoEdicaoAlteracaoContrato = 183,
            ConfirmaçãoEdicaoAlteracaoContratoR2 = 185,
            LiquidacaoMercadoPrimario = 186,
            LiquidacaoMercadoPrimarioR2 = 188,
            BaixaValorLiquidar = 189,
            BaixaValorLiquidarR2 = 191,
            RestabelecimentoBaixa = 192,
            RestabelecimentoBaixaR2 = 194,
            CancelamentoValorLiquidar = 195,
            CancelamentoValorLiquidarR2 = 197,
            EdicaoCancelamentoValorLiquidar = 198,
            EdicaoCancelamentoValorLiquidarR2 = 200,
            ConfirmacaoEdicaoCancelamentoValorLiquidar = 201,
            ConfirmacaoEdicaoCancelamentoValorLiquidarR2 = 203,
            VinculacaoContratos = 204,
            VinculacaoContratosR2 = 206,
            AnulacaoEvento = 207,
            AnulacaoEventoR2 = 209,
            CorretoraRequisitaClausulasEspecificas = 210,
            CorretoraRequisitaClausulasEspecificasR2 = 212,
            IFInformaClausulasEspecificas = 213,
            IFInformaClausulasEspecificasR2 = 215,
            ManutencaoCadastroAgenciaCentralizadoraCambio = 216,
            CredenciamentoDescredenciamentoDispostoRMCCI = 218,
            IncorporacaoContratos = 220,
            IncorporacaoContratosR2 = 222,
            AceiteRejeicaoIncorporacaoContratos = 223,
            AceiteRejeicaoIncorporacaoContratosR2 = 225,
            AvisoAceiteRejeicaoIncorporacaoContratos = 226,
            ConsultaContratosEmSer = 227,
            ConsultaEventosUmDia = 229,
            ConsultaDetalhamentoContratoInterbancario = 231,
            ConsultaEventosContratoMercadoPrimario = 233,
            ConsultaEventosContratoIntermediadoMercadoPrimario = 235,
            ConsultaHistoricoIncorporacoes = 237,
            ConsultaContratosIncorporacao = 239,
            ConsultaCadeiaIncorporacoesContrato = 241,
            ConsultaPosicaoCambioMoeda = 243,
            AtualizaçãoInclusãoInstrucoesPagamento = 245,
            ConsultaInstrucoesPagamento = 247,
            RegistroOperacaoInterbancaria = 249,
            RegistroOperacaoArbitragem = 253,
            ComplementoInformacoesContratacaoInterbancarioViaLeilao = 257,
            IFInformaLiquidacaoInterbancaria = 258,
            IFCamaraConsultaContratosCambioMercadoInterbancario = 260
        }

        public enum EnumTipoBackOffice
        {
            FundosProprios = 1,
            Tesouraria = 2,
            FundosTerceiros = 3,
            Corretoras = 4,
            GBM = 5,
            GrandesEmpresas = 6,
            Empresas = 7,
            Comex = 8,
            Todos = 9   // Alterado para 9 para possibilitar a inclusao de novos tipos de backoffice em ordem
        }

        public enum EnumLocalLiquidacao
        {
            STR = 1,
            CIP = 2,
            COMPE_NOTURNO = 3,
            CONTA_CORRENTE = 4,
            CETIP = 5,
            CLBCAcoes = 6,
            CLBCTpPriv = 7,
            CLBCTPub = 8,
            SELIC = 13,
            BMC = 10,
            BMD = 11,
            COMPE_DIURNO = 15,
            BMA = 17,
            PAG = 19,
            Compuls = 20,
            CCR = 21,
            CAM = 22
        }

        public enum EnumTipoMovimentoPJ
        {
            Previsto = 100,
            Realizado = 200,
            EstornoPrevisto = 300,
            EstornoRealizado = 400
        }

        public enum EnumTipoCaixaPJ
        {
            Reserva = 100,
            Futuro = 200
        }

        public enum EnumTipoDebitoCredito
        {
            Debito = 1,
            Credito = 2
        }

        public enum EnumTipoDebitoCreditoPJME
        {
            Credito = 1,
            Debito = 2
        }

        public enum EnumTipoEntradaSaida
        {
            Entrada = 1,
            Saida = 2
        }

        public enum EnumTipoProcessamentoPJ
        {
            OnLine = 1,
            Batch = 2
        }

        public enum EnumTipoEnvioPJ
        {
            Total = 1,
            Parcial = 2
        }

        public enum EnumTipoMoedaPJ
        {
            MoedaNacional = 1,
            MoedaEstrangeira = 2
        }

        public enum EnumTipoMovimento
        {
            Previsto = 1,
            RealizadoSolicitado = 2,
            RealizadoConfirmado = 3,
            EstornoPrevisto = 4,
            EstornoRealizadoSolicitado = 5,
            EstornoRealizadoConfirmado = 6,
            PrevistoCompromIda = 7,
            EstornoPrevistoCompromIda = 8
        }

        public enum EnumIndicadorSimNao
        {
            Sim = 1,
            Nao = 2
        }

        public enum EnumTipoOperacao
        {
            RegistroOperacaoInterbancariaSemTelaCega = 230,
            ConfirmacaoOperacaoInterbancariaSemTelaCega = 231,
            RegistroOperacaoInterbancariaSemCamara = 232,
            ConfirmacaoOperacaoInterbancariaSemCamara = 233,
            RegistroOperacaoInterbancarioEletronico = 234,
            InformaContrArbitParceiroExteriorPaisPropriaIF = 235,
            InformaOperacaoArbitragemParceiroPais = 236,
            InformaConfirmacaoOperArbitragemParceiroPais = 237,
            CAMInformaContratacaoInterbancarioViaLeilao = 238,
            InformaLiquidacaoInterbancaria = 239,
            ConsultaContratosCambioMercadoInterbancario = 240
        }

        public enum EnumTipoAcao
        {
            AlteracaoHorarioAgendamento = 1,
            AlteracaoTipoCompromisso = 2,
            CancelamentoSolicitado = 3,
            CancelamentoEnviado = 4,
            EstornoSolicitado = 5,
            EstornoEnviado = 6,
            RejeicaoConcordancia = 7,            //<- Rejeicao de Concordancia do BackOffice
            RejeicaoDiscordancia = 8,            //<- Rejeicao de Discordancia do BackOffice
            EnviadaLDL1002 = 9,
            EnviadaSEL1023 = 10,
            EnviadaLDL1016 = 11,
            RejeicaoConcordanciaAdmArea = 12,    //<- Rejeicao de Discordancia do Adm de Area
            RejeicaoDiscordanciaAdmArea = 13,    //<- Rejeicao de Discordancia do Adm de Area
            EnviadaLDL0003Concordancia = 14,
            EnviadaLDL0003Discordancia = 15,
            EnviadoPagamento = 16,
            EnviadoPagamentoContingencia = 17,
            AjusteValor = 18,
            EnviadaLTR0002Concordancia = 19,
            EnviadaLTR0002Discordancia = 20,
            EnviadoPagamentoSTR = 21,
            EnviadoPagamentoBACEN = 22,
            EnviadaLDL1006 = 23,
            EnviadaLTR0008Concordancia = 24,
            EnviadaLTR0008Discordancia = 25,
            ConcordanciaEmSer = 26,
            ConcordanciaManualEmSer = 27,
            RejeicaoConcordanciaEmSer = 28,
            RejeicaoConcordanciaManualEmSer = 29,
            Liberacao = 30,
            LiberacaoAntecipada = 31,
            Liquidacao = 32,
            EnviadaSTR0004Pagamento = 33,
            EnviadaSTR0007 = 34,
            EnviadaLDL1003 = 35,
            Concordancia = 36,
            EnviadaSEL1007 = 37,
            CancelamentoPendente = 38,
            CancelamentoRejeitado = 39,
            CancelamentoSolicitadoComMensagem = 40,
            EnviadoRecebimento = 41,
            EnviadaBMC0102 = 42,
            EnviadaBMC0012 = 43,
            RegistroContingencia = 44,
            DiscordanciaAdmBO = 45,
            LTR0001ComISPBJaExistente = 46,
            LTR0007ComISPBJaExistente = 47,
            RejeicaoPorDuplicidade = 48,
            EnviadaLTR0004Pagamento = 49,
            EnviadaLTR0003Pagamento = 50,
            EnviadaBMC0001 = 51,
            EnviadaCAM0002 = 52,
            EnviadaBMC0002 = 53,
            EnviadaBMC0003 = 54,
            EnviadaConfirmacaoContingencia = 55,
            PreviaLiquidada = 56,
            EnviadaCAM0054 = 57,
            EnviadaCAM0006 = 58,
            EnviadaCAM0009 = 59
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // ATENCAO: quando criar um tipo acao novo, colocar tambem a descricao na funcao VB 'A8.basA8LQS.fgDescricaoTipoAcao'
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        }

        public enum enumJustificativa
        {
            SistemaEmContingencia = 1,
            RegraWorkflow = 2,
            GradeHorario = 3,
            CadastroGrupoUsuario = 4,
            OperacaoRetroativa = 5,
            EntradaManual = 6,
            CompromissadaIda = 7,
            VeicluloLegal = 8,
            CNPJContraparte = 9,
            IdentificadorTitulo = 10,
            DataVencimentoTitulo = 11,
            NumeroComando = 12,
            DataOperacao = 13,
            DebitoCredito = 14,
            DataLiquidacao = 15,
            QuantidadeTitulo = 16,
            PU = 17,
            ValorFinanceiro = 18,
            ContaCustodia = 19,
            TaxaNegociada = 20,
            TitulaCustodiante = 21,
            TipoNegociacao = 22,
            MensagemOperacaoNaoInformado = 23,
            SequenciaOperacao = 24,
            OperacaoJaAlterada = 25,
            MensagemNaoEncontrada = 26,
            MensagemJaAlterada = 27,
            OperacaoNaoEncontrada = 28,
            IndentificadorContraparteCamara = 29,
            CodigoOperacaoCETIP = 30,
            DescricaoAtivo = 31,
            ModalidadeLiquidacao = 32,
            TipoMensagem = 33,
            OperacaoNaoLiquidada = 34,
            DataRetorno = 35,
            PrazoDiasRetorno = 36,
            ValorRetorno = 37,
            IndentificadorParticipanteCamara = 38,
            CodigoPraca = 39,
            MoedaEstrangeira = 40,
            ValorMoedaEstrangeira = 41,
            CodigoContratacaoSISBACEN = 42,
            CanalSISBACENCorretora = 43,
            ComponenteMBSNaoEncontrado = 44,
            ErroAoAcessarFuncaoMBS = 45,
            ErroNaoIdentificadoMBS = 46,
            NumeroIdentNegociacaoBMC = 47,
            SemMensagemCCR0006 = 48,
            CodigoAssociacaoCambio = 49,
            SemMensagemBMC0015 = 50,
            VariasMensagensBMC0015 = 51
        }

        public enum enumTipoSolicitacao
        {
            Inclusao = 1,
            Complementacao = 2,
            Cancelamento = 3,
            CancelamentoComMensagem = 4,
            RetornoLegado = 5, //Utilizado somente internamente não irá vir do BUS
            Alteracao = 6,
            LivreMovimentacao = 7,
            CancelamentoPorLastro = 8,
            Reativacao = 9,
            Confirmacao = 10
        }
        #endregion

        #region >>> ConvertDtToDateTime >>>
        public static DateTime ConvertDtToDateTime(string dataYYYYMMDD)
        {

            try
            {
                System.Globalization.CultureInfo PtBR = new System.Globalization.CultureInfo("pt-BR");
                return DateTime.Parse(dataYYYYMMDD.Substring(6, 2) + "/" + dataYYYYMMDD.Substring(4, 2) + "/" + dataYYYYMMDD.Substring(0, 4), PtBR);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }
        #endregion

        #region >>> Tratamentos Base64 >>>
        public static string Base64Encode(string data)
        {
            try
            {
                byte[] encData_byte = new byte[data.Length];
                encData_byte = System.Text.Encoding.UTF8.GetBytes(data);
                string encodedData = Convert.ToBase64String(encData_byte);
                return encodedData;
            }
            catch (Exception ex)
            {
                throw new Exception("Base64Encode()" + ex.ToString());
            }
        }

        public static string Base64Decode(string data)
        {
            try
            {
                System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
                System.Text.Decoder utf8Decode = encoder.GetDecoder();

                byte[] todecode_byte = Convert.FromBase64String(data);
                int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
                char[] decoded_char = new char[charCount];
                utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
                string result = new String(decoded_char);
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception("Base64Decode()" + ex.ToString());
            }
        }
        #endregion

        #region >>> LerNode >>>
        public static string LerNode (XmlDocument xmlDocument, string tag)
        {
            try
            {
                return xmlDocument.SelectSingleNode("//" + tag).InnerText;
            }
            catch
            {
                return string.Empty;
            }
        }
        #endregion

        #region >>> AppendNode >>>
        //Adiona um node a um objeto XML.
        public static void AppendNode(ref XmlDocument xmlDocument,
                                      string nodeContext,
                                      string nodeName,
                                      string nodeValue)
        {

            XmlNode NodeAux;
            XmlNode NodeContextAux;

            try
            {
                if (nodeName == string.Empty) //Parâmetro pstrNodeName deve ser diferente de vbNullString
                {
                    throw new Exception("Comum.AppendNode() - Parâmetro nodeName deve ser diferente de string.Empty ");
                }

                NodeContextAux = xmlDocument.DocumentElement.SelectSingleNode("//" + nodeContext);
                NodeAux = xmlDocument.CreateElement(nodeName);
                NodeAux.InnerXml = nodeValue;
                NodeContextAux.AppendChild(NodeAux);
            }
            catch (Exception ex)
            {
                throw new Exception("Comum.AppendNode() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< Constantes >>>
        public static readonly string UsuarioSistema = "SISTEMA";
        public static readonly string CodigoMoeda = "0790";
        public static readonly string NomeMaquina = System.Environment.MachineName;
        #endregion
    }
}
