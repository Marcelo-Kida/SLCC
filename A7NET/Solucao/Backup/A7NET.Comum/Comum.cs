using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace A7NET.Comum
{
    public static class Comum
    {
        #region <<< Enumeradores >>>
        public enum EnumTipoMensagemEntrada
        {
            MensagemNZA8 = 1000, // Usado CAM
            MensagemA8PJRealizado = 1001,
            MensagemA8NZ = 1002, // Usado CAM
            MensagemErroNZA8 = 1003, // Usado CAM
            MensagemA8PJPrevisto = 1004,
            MensagemIDADV = 1005,
            MensagemRetornoDV = 1006,
            MensagemIDABG = 1007,
            MensagemRetornoBG = 1008,
            MensagemPZErro = 1009,
            MensagemPZR1 = 1010,
            MensagemPZR2 = 1011,
            MensagemSTR0010R2PZA8 = 1012,
            MensagemRetornoPZOk = 1013,
            MensagemA8PJMoedaEstrangeira = 1014,
            MensagemA8NZR1_HD = 1999
        }

        public enum EnumTipoEntradaMensagem
        {
            EntradaXML = 1,
            EntradaString = 2,
            EntradaCSV = 3
        }

        public enum EnumTipoSaidaMensagem
        {
            NaoSeAplica = 0,
            SaidaXML = 1,
            SaidaString = 2,
            SaidaCSV = 3,
            SaidaStringXML = 4,
            SaidaCSVXML = 5
        }

        public enum EnumTipoParteSaida
        {
            SemParte = 0,
            ParteId = 1,
            ParteXML = 2,
            ParteSTR = 3,
            ParteCSV = 4,
            ParteIdLocalLiquidacao = 5,
            ParteIdOperacao = 6
        }

        public enum EnumTipoTraducaoBookSPB
        {
            BookNovo_BookAntigo = 1,
            BookAntigo_BookNovo = 2
        }

        public enum EnumNaturezaMensagem
        {
            MensagemEnvio = 1,
            MensagemConsulta = 2,
            MensagemECO = 3
        }

        public enum EnumTipoMovimentoPJ
        {
            Previsto = 100,
            Realizado = 200,
            EstornoPrevisto = 300,
            EstornoRealizado = 400
        }

        public enum EnumFluxoMonitor
        {
            FLUXO_MONITOR_NORMAL = 1,
            FLUXO_MONITOR_LEGADINHO = 2,
            FLUXO_MONITOR_CONT_MAXIMA = 3,
            FLUXO_MONITOR_MSG_EXTERNA = 4,
            FLUXO_MONITOR_MSG_CARGA = 5
        }

        public enum EnumOcorrencia
        {
            RecebimentoBemSucedido = 1,
            PostagemBemSucedida = 2,
            ConfirmacaoEntregaRecebida = 3,
            ConfirmacaoRetiradaRecebida = 4,
            RejeicaoNaoAutenticidade = 5,
            CanceladaErroTraducao = 6,
            ErroPostagem = 7,
            FalhaNaoPrevista = 8,
            UsuarioInativo = 9
        }

        public enum EnumTipoSolicitacao
        {
            Inclusao = 1,
            Complementacao = 2,
            Cancelamento = 3,
            CancelamentoComMensagem = 4,
            RetornoLegado = 5, //Utilizado somente internamente não irá vir do BUS
            Alteracao = 6,
            LivreMovimentacao = 7,
            CancelamentoPorLastro = 8,
            Reativacao = 9
        }

        public enum EnumStatusMonitor
        {
            MONITOR_ENVIADA_NZ = 1,
            MONITOR_ENVIADA_PK = 2,
            MONITOR_ENVIADA_PILOTO = 3,
            MONITOR_ERRO_PKI = 4,
            MONITOR_ENVIADA_EXTERNO = 5,
            MONITOR_RECEB_COA = 6,
            MONITOR_RECEB_COD = 7,
            MONITOR_ENVIADO_R1_NZ = 11,
            MONITOR_ENVIADO_R1_LEGADO = 12,
            MONITOR_RECEB_R1 = 13,
            MONITOR_RECEB_MSG_REQ_CMAX = 14,
            MONITOR_ENV_ERRO_EXT_NZ = 21,
            MONITOR_ENV_ERRO_EXT_LEGADO = 22,
            MONITOR_RECEB_ERRO_EXTERNO = 23,
            MONITOR_ENV_CANCEL_PILOTO_NZ = 31,
            MONITOR_ENV_CANCEL_PILOTO_LEGADO = 32,
            MONITOR_RECEB_CANCEL_PILOTO = 33,
            MONITOR_ENV_ERRO_PKNZ = 41,
            MONITOR_ENV_ERRO_PKLEGADO = 42,
            MONITOR_RECEB_ERRO_PK = 43,
            MONITOR_ENV_ERRO_NZ = 51,
            MONITOR_RECEB_ERRO_NZ = 52,
            MONITOR_ENV_DUPLIC = 61,
            MONITOR_RECEB_DUPLIC = 62,
            MONITOR_ENV_MENS_EXTERNA_NZ = 71,
            MONITOR_ENV_MENS_EXTERNA_LEGADO = 72,
            MONITOR_RECEB_MENS_EXTERNA = 73,
            MONITOR_ENV_CARGA_NZ = 81,
            MONITOR_CARGA_EFETUADA = 82,
            MONITOR_REGUL_ERRO_PKI = 0,
            MONITOR_ERRO_PKI_MSG_EXT = 93,
        }

        public enum EnumTipoMensagemLQS
        {
            Definitiva = 1,
            Compromissada = 3,
            RetornoCompromissada = 5,
            TermoD0 = 7,
            TermoDataLiquidacao = 9,
            Leilao = 11,
            VinculoDesvinculoTransf = 13,
            EventosSelic = 15,
            DespesasSelic = 17,
            Redesconto = 19,
            PgtoRedesconto = 21,
            ConversaoRedesconto = 23,
            AlteracaoDadosContaCorrente = 27,
            TransferenciaLDL_BMA = 30,
            RegistroOperacaoBMA = 32,
            LiquidacaoOperacoesBMA = 34,
            LiquidacaoEventosBMA = 36,
            EspecificacaoOperacoesBMA = 38,
            IntermediacaoOperacoesInternasBMA = 40,
            LiquidacaoFisicaOperacaoBMA = 42,
            OperacoesComCorretorasCETIP = 50,
            MovimentacoesInstFinancCETIP = 52,
            MovimentacoesCustodiaCETIP = 54,
            ResgateFundoInvestimentoCETIP = 56,
            ExercicioDesistenciaCETIP = 58,
            ConversaoPermutaValorImobCETIP = 60,
            EspecificacaoQuantidadesCotasCETIP = 62,
            OperacaoDefinitivaCETIP = 64,
            OperacaoCompromissadaCETIP = 66,
            OperacaoRetornoAntecipacaoCETIP = 68,
            OperacaoRetencaoIRF_CETIP = 70,
            RegistroContratoSWAP = 72,
            RegistroOperacaoesCETIP = 74,
            RegistroContratoTermoCETIP = 76,
            ExercicioOpcaoContratoSwapCETIP = 78,
            AntecipacaoResgateContratoDerivativoCETIP = 80,
            LancamentoPU_CETIP = 82,
            MovimentacoesContratoDerivativo = 84,
            EventoJurosCETIP = 86,
            DespesasCETIP = 88,
            LivreMovimentacao = 90,
            RegistroContratoSWAPCetip21 = 94,
            
            //CBLC
            RegistroLiquidacaoMultilateralCBLC = 120,
            RegistroLiquidacaoBrutaCBLC = 122,
            RegistroLiquidacaoEventoCBLC = 124,

            //BMF
            RegistroLiquidacaoMultilateralBMF = 126,

            //BMC
            RegistroOperacoesBMC = 130,
            RegistroOperacoesBMCRetorno = 131,
            LiquidacaoMultilateralBMC = 132,
            TransferenciasBMC = 134,
            DespesasBMC = 136,
            RegistroOperacoesRodaDolar = 138,

            //TED
            EnvioTEDClientes = 150,
            
            //Automatizaco Pag Despesas
            EnvioPagDespesas = 154,
            
            //LancamentoContaCorrenteBG
            LancamentoContaCorrenteBG = 156,
            
            //BACEN
            MensagemNZ = 1000,
            MensagemErroNZ = 1003,
            MensagemCCDV = 1006,
            MensagemCCBG = 1008,
            MensagemErroPZ = 1009,
            MensagemR1PZ = 1010,
            MensagemR2PZ = 1011,
            MensagemSTR0010R2PZA8 = 1012,
            
            //CCR
            ConsultaOperacaoCCR = 158,
            EmissaoOperacaoCCR = 160,
            NegociacaoOperacaoCCR = 162,
            DevolucaoRecolhimentoEstornoReembolsoCCR = 164,
            ConsultaLimitesImportacaoCCR = 166
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
                if (nodeName == string.Empty) //Parametro NodeName deve ser diferente de vbNullString
                {
                    throw new Exception("Comum.AppendNode() - Parâmetro nodeName deve ser diferente de string.Empty.");
                }

                if (nodeContext == string.Empty)
                {
                    //Se a Tag for o root passar NodeContextAux = Nome Tag Principal
                    NodeContextAux = xmlDocument;
                }
                else
                {
                    NodeContextAux = xmlDocument.DocumentElement.SelectSingleNode("//" + nodeContext);
                }

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

        #region >>> AppendAttribute >>>
        //Adiona um attribute a um objeto XML.
        public static void AppendAttribute(ref XmlDocument xmlDocument,
                                           string nodeContext,
                                           string nomeAtributo,
                                           string valorAtributo)
        {

            XmlAttribute Attribute;
            XmlNode NodeContextAux;

            try
            {
                if (nomeAtributo == string.Empty) //Parametro nomeAtributo deve ser diferente de nulo
                {
                    throw new Exception("Comum.AppendNode() - Parâmetro nomeAtributo deve ser diferente de nulo.");
                }

                if (nodeContext == string.Empty) //Parametro nodeContext deve ser diferente de nulo
                {
                    throw new Exception("Comum.AppendNode() - Parâmetro nodeContext deve ser diferente de nulo.");
                }

                NodeContextAux = xmlDocument.DocumentElement.SelectSingleNode("//" + nodeContext);
                
                Attribute = xmlDocument.CreateAttribute(nomeAtributo);
                Attribute.InnerText = valorAtributo;
                NodeContextAux.Attributes.SetNamedItem(Attribute);

            }
            catch (Exception ex)
            {

                throw new Exception("Comum.AppendNode() - " + ex.ToString());
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

        #region <<< ValorToSTR >>>
        public static string ValorToSTR(string valor, string tipo, int tamanho, int decimais, string obrigatorio, string nomeTag)
        {
            string[] Numeros;
            string Numero;
            string NumeroDecimal;
            string ValorRetorno = "";

            try
            {
                switch (tipo.ToUpper())
                {
                    case "STRING":
                        ValorRetorno = valor.PadRight(tamanho, ' ');

                        if (obrigatorio == "1")
                        {
                            if (ValorRetorno.Trim() == "")
                            {
                                ValorRetorno = "";
                            }
                        }
                        break;

                    case "NUMBER":
                        if (valor.IndexOf(",") > 0)
                        {
                            Numeros = valor.Split(',');
                            Numero = Numeros[0];
                            NumeroDecimal = Numeros[1];

                            if (Math.Abs(decimais) > 0)
                            {
                                ValorRetorno = Numero.PadLeft(tamanho - decimais, '0') + NumeroDecimal.PadRight(decimais, '0');
                            }
                            else
                            {
                                ValorRetorno = Numero.PadLeft(tamanho, '0');
                            }
                        }
                        else
                        {
                            ValorRetorno = valor.PadLeft(tamanho - decimais, '0');
                            ValorRetorno = ValorRetorno.PadRight(tamanho, '0');
                        }

                        if (obrigatorio == "1")
                        {
                            if (ValorRetorno.Trim() == ""
                            || ValorRetorno.Replace('0', ' ').Replace(',', ' ').Trim() == "")
                            {
                                if (nomeTag.PadLeft(3) == "HO_")
                                {
                                    ValorRetorno = "0000";
                                }
                                else
                                {
                                    ValorRetorno = "";
                                }
                            }
                        }
                        break;

                    default:
                        break;
                }

                return ValorRetorno;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< ValorToXML >>>
        public static string ValorToXML(string valor, string tipo, int tamanho, int decimais, string obrigatorio)
        {
            string[] Numeros;
            string Numero;
            string NumeroDecimal;
            string Sinal;
            string ValorRetorno = "";
            decimal OutN;

            try
            {
                switch (tipo.ToUpper())
                {
                    case "STRING":
                        ValorRetorno = valor.Trim();
                        break;

                    case "NUMBER":
                        if (!decimal.TryParse(valor, out OutN))
                        {
                            return valor.Trim();
                        }

                        if (obrigatorio == "1")
                        {
                            if (valor.Replace('0', ' ').Trim() == "")
                            {
                                return "0";
                            }
                        }

                        if (OutN < 0)
                        {
                            Sinal = "-";
                        }
                        else
                        {
                            Sinal = "";
                        }

                        if (valor.IndexOf(",") > 0)
                        {
                            Numeros = valor.Split(',');
                            Numero = Numeros[0];
                            NumeroDecimal = Numeros[1];

                            if (Math.Abs(decimais) > 0)
                            {
                                ValorRetorno = Convert.ToString(int.Parse(Numero.PadLeft(tamanho - decimais)) + "," + NumeroDecimal.PadRight(decimais));
                            }
                            else
                            {
                                ValorRetorno = Convert.ToString(int.Parse(Numero.PadLeft(tamanho)));
                            }

                            ValorRetorno = Sinal + ValorRetorno;
                        }
                        else
                        {
                            ValorRetorno = valor.PadLeft(tamanho);
                            if (tamanho <= 15)
                            {
                                ValorRetorno = Convert.ToString(int.Parse(valor.PadLeft(tamanho)));
                                ValorRetorno = Sinal + ValorRetorno;
                            }
                        }
                        break;

                    default:
                        break;

                }

                return ValorRetorno;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< TipoEntradaToSTR >>>
        public static string TipoEntradaToSTR(int tipoEntrada)
        {
            switch (tipoEntrada)
            {
                case (int)A7NET.Comum.Comum.EnumTipoEntradaMensagem.EntradaString:
                    return "String";

                case (int)A7NET.Comum.Comum.EnumTipoEntradaMensagem.EntradaXML:
                    return "XML";

                default:
                    return "";

            }

        }
        #endregion

    }
}
