using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using A8NET.Comum;
using A8NET.Data;

namespace A8NET.Mensagem
{
    public class GestaoCaixa
    {
        #region <<< Variaveis >>>
        EstruturaRemessaMovimento DadosRemessaMovimento;
        EstruturaMaioresValores DadosMaioresValores;
        EstruturaMoedaEstrangeira DadosMoedaEstrangeira;
        #endregion

        #region <<< Estruturas >>>
        private struct EstruturaRemessaMovimento
        {
            public string TipoRemessa;               // String * 3
            public string CodigoRemessa;             // String * 23
            public string DataRemessa;               // String * 8
            public string HoraRemessa;               // String * 4
            public string CodigoEmpresa;             // String * 5
            public string SiglaSistema;              // String * 3
            public string CodigoMoeda;               // String * 4
            public string CodigoBanqueiro;           // String * 12
            public string TipoCaixa;                 // String * 3
            public string CodigoItemCaixa;           // String * 9
            public string TipoAtivoPassivo;          // String * 1
            public string CodigoProduto;             // String * 4
            public string TipoConta;                 // String * 3
            public string CodigoSegmento;            // String * 3
            public string EventoFinanceiro;          // String * 3
            public string CodigoIndexador;           // String * 3
            public string CodigoLocalLiquidacao;     // String * 4
            public string CodigoFaixaValor;          // String * 3
            public string TipoMovimento;             // String * 3
            public string DataMovimento;             // String * 8
            public string HoraMovimento;             // String * 4
            public string TipoEntradaSaida;          // String * 1
            public string ValorMovimento;            // String * 19
            public string ValorContabil;             // String * 19
            public string TipoProcessamento;         // String * 1
            public string TipoEnvio;                 // String * 1
            public string Filler;                    // String * 46

            public void Inicializa()
            {
                this.TipoRemessa = string.Empty;
                this.CodigoRemessa = string.Empty;
                this.DataRemessa = string.Empty;
                this.HoraRemessa = string.Empty;
                this.CodigoEmpresa = string.Empty;
                this.SiglaSistema = string.Empty;
                this.CodigoMoeda = string.Empty;
                this.CodigoBanqueiro = string.Empty;
                this.TipoCaixa = string.Empty;
                this.CodigoItemCaixa = string.Empty;
                this.TipoAtivoPassivo = string.Empty;
                this.CodigoProduto = string.Empty;
                this.TipoConta = string.Empty;
                this.CodigoSegmento = string.Empty;
                this.EventoFinanceiro = string.Empty;
                this.CodigoIndexador = string.Empty;
                this.CodigoLocalLiquidacao = string.Empty;
                this.CodigoFaixaValor = string.Empty;
                this.TipoMovimento = string.Empty;
                this.DataMovimento = string.Empty;
                this.HoraMovimento = string.Empty;
                this.TipoEntradaSaida = string.Empty;
                this.ValorMovimento = string.Empty;
                this.ValorContabil = string.Empty;
                this.TipoProcessamento = string.Empty;
                this.TipoEnvio = string.Empty;
                this.Filler = string.Empty;
            }

            public override string ToString()
            {
                StringBuilder Concatena = new StringBuilder();

                Concatena.Append(this.TipoRemessa.Trim().PadRight(3, ' '));
                Concatena.Append(this.CodigoRemessa.Trim().PadRight(23, ' '));
                Concatena.Append(this.DataRemessa.Trim().PadRight(8, ' '));
                Concatena.Append(this.HoraRemessa.Trim().PadRight(4, ' '));
                Concatena.Append(this.CodigoEmpresa.Trim().PadLeft(5, '0'));
                Concatena.Append(this.SiglaSistema.Trim().PadRight(3, ' '));
                Concatena.Append(this.CodigoMoeda.Trim().PadRight(4, ' '));
                Concatena.Append(this.CodigoBanqueiro.Trim().PadRight(12, ' '));
                Concatena.Append(this.TipoCaixa.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoItemCaixa.Trim().PadLeft(9, '0'));
                Concatena.Append(this.TipoAtivoPassivo.Trim().PadLeft(1, '0'));
                Concatena.Append(this.CodigoProduto.Trim().PadLeft(4, '0'));
                Concatena.Append(this.TipoConta.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoSegmento.Trim().PadLeft(3, '0'));
                Concatena.Append(this.EventoFinanceiro.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoIndexador.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoLocalLiquidacao.Trim().PadLeft(4, '0'));
                Concatena.Append(this.CodigoFaixaValor.Trim().PadLeft(3, '0'));
                Concatena.Append(this.TipoMovimento.Trim().PadLeft(3, '0'));
                Concatena.Append(this.DataMovimento.Trim().PadRight(8, ' '));
                Concatena.Append(this.HoraMovimento.Trim().PadRight(4, ' '));
                Concatena.Append(this.TipoEntradaSaida.Trim().PadRight(1, '0'));
                Concatena.Append(this.ValorMovimento.Trim().PadLeft(19, '0'));
                Concatena.Append(this.ValorContabil.Trim().PadLeft(19, '0'));
                Concatena.Append(this.TipoProcessamento.Trim().PadRight(1, '0'));
                Concatena.Append(this.TipoEnvio.Trim().PadRight(1, '0'));
                Concatena.Append(this.Filler.Trim().PadRight(46, ' '));

                return Concatena.ToString();
            }
        }
        private struct EstruturaMaioresValores
        {
            public string TipoRemessa;              //  String * 3
            public string CodigoRemessa;            //  String * 23
            public string DataRemessa;              //  String * 8
            public string HoraRemessa;              //  String * 4
            public string CodigoEmpresa;            //  String * 5
            public string SiglaSistema;             //  String * 3
            public string CodigoMoeda;              //  String * 4
            public string CodigoBanqueiro;          //  String * 12
            public string TipoCaixa;                //  String * 3
            public string CodigoItemCaixa;          //  String * 9
            public string CodigoProduto;            //  String * 4
            public string TipoConta;                //  String * 3
            public string CodigoSegmento;           //  String * 3
            public string CodigoEventoFinanceiro;   //  String * 3
            public string CodigoIndexador;          //  String * 3
            public string CodigoLocalLiquidacao;    //  String * 4
            public string TipoMovimento;            //  String * 3
            public string DataMovimento;            //  String * 8
            public string HoraMovimento;            //  String * 4
            public string TipoEntradaSaida;         //  String * 1
            public string ValorMovimento;           //  String * 17
            public string CodigoBanco;              //  String * 3
            public string CodigoAgencia;            //  String * 5
            public string NumeroContaCorrente;      //  String * 13
            public string TipoPessoa;               //  String * 1
            public string CodigoCNPJ_CPF;           //  String * 15
            public string NomeCliente;              //  String * 64
            public string TipoProcessamento;        //  String * 1
            public string TipoEnvio;                //  String * 1
            public string NumeroOperacaoSelic;      //  String * 6
            public string ContaCedenteCessionaria;  //  String * 9
            public string Filler;                   //  String * 5

            public void Inicializa()
            {
                this.TipoRemessa = string.Empty;
                this.CodigoRemessa = string.Empty;
                this.DataRemessa = string.Empty;
                this.HoraRemessa = string.Empty;
                this.CodigoEmpresa = string.Empty;
                this.SiglaSistema = string.Empty;
                this.CodigoMoeda = string.Empty;
                this.CodigoBanqueiro = string.Empty;
                this.TipoCaixa = string.Empty;
                this.CodigoItemCaixa = string.Empty;
                this.CodigoProduto = string.Empty;
                this.TipoConta = string.Empty;
                this.CodigoSegmento = string.Empty;
                this.CodigoEventoFinanceiro = string.Empty;
                this.CodigoIndexador = string.Empty;
                this.CodigoLocalLiquidacao = string.Empty;
                this.TipoMovimento = string.Empty;
                this.DataMovimento = string.Empty;
                this.HoraMovimento = string.Empty;
                this.TipoEntradaSaida = string.Empty;
                this.ValorMovimento = string.Empty;
                this.CodigoBanco = string.Empty;
                this.CodigoAgencia = string.Empty;
                this.NumeroContaCorrente = string.Empty;
                this.TipoPessoa = string.Empty;
                this.CodigoCNPJ_CPF = string.Empty;
                this.NomeCliente = string.Empty;
                this.TipoProcessamento = string.Empty;
                this.TipoEnvio = string.Empty;
                this.NumeroOperacaoSelic = string.Empty;
                this.ContaCedenteCessionaria = string.Empty;
                this.Filler = string.Empty;
            }

            public override string ToString()
            {
                StringBuilder Concatena = new StringBuilder();

                Concatena.Append(this.TipoRemessa.Trim().PadRight(3, ' '));
                Concatena.Append(this.CodigoRemessa.Trim().PadRight(23, ' '));
                Concatena.Append(this.DataRemessa.Trim().PadRight(8, ' '));
                Concatena.Append(this.HoraRemessa.Trim().PadRight(4, ' '));
                Concatena.Append(this.CodigoEmpresa.Trim().PadLeft(5, '0'));
                Concatena.Append(this.SiglaSistema.Trim().PadRight(3, ' '));
                Concatena.Append(this.CodigoMoeda.Trim().PadRight(4, ' '));
                Concatena.Append(this.CodigoBanqueiro.Trim().PadRight(12, ' '));
                Concatena.Append(this.TipoCaixa.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoItemCaixa.Trim().PadLeft(9, '0'));
                Concatena.Append(this.CodigoProduto.Trim().PadLeft(4, '0'));
                Concatena.Append(this.TipoConta.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoSegmento.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoEventoFinanceiro.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoIndexador.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoLocalLiquidacao.Trim().PadLeft(4, '0'));
                Concatena.Append(this.TipoMovimento.Trim().PadLeft(3, '0'));
                Concatena.Append(this.DataMovimento.Trim().PadRight(8, ' '));
                Concatena.Append(this.HoraMovimento.Trim().PadRight(4, ' '));
                Concatena.Append(this.TipoEntradaSaida.Trim().PadRight(1, '0'));
                Concatena.Append(this.ValorMovimento.Trim().PadLeft(17, '0'));
                Concatena.Append(this.CodigoBanco.Trim().PadLeft(3, '0'));
                Concatena.Append(this.CodigoAgencia.Trim().PadLeft(5, '0'));
                Concatena.Append(this.NumeroContaCorrente.Trim().PadLeft(13, '0'));
                Concatena.Append(this.TipoPessoa.Trim().PadRight(1, '0'));
                Concatena.Append(this.CodigoCNPJ_CPF.Trim().PadLeft(15, '0'));
                Concatena.Append(this.NomeCliente.Trim().PadRight(64, ' '));
                Concatena.Append(this.TipoProcessamento.Trim().PadRight(1, '0'));
                Concatena.Append(this.TipoEnvio.Trim().PadRight(1, '0'));
                Concatena.Append(this.NumeroOperacaoSelic.Trim().PadRight(6, '0'));
                Concatena.Append(this.ContaCedenteCessionaria.Trim().PadRight(9, '0'));
                Concatena.Append(this.Filler.Trim().PadRight(5, ' '));

                return Concatena.ToString();
            }
        }
        private struct EstruturaMoedaEstrangeira
        {
            public string TipoRemessa;            // String * 3
            public string CodigoEmpresa;          // String * 5
            public string SiglaSistema;           // String * 3
            public string IdentificadorMovimento; // String * 25
            public string CodigoMoeda;            // String * 4
            public string CodigoBanqueiroSwift;   // String * 30
            public string CodigoProduto;          // String * 4
            public string DataMovimento;          // String * 8
            public string CodigoReferenciaSwift;  // String * 16
            public string TipoEntradaSaida;       // String * 1
            public string ValorMovimento;         // String * 19
            public string NomeCliente;            // String * 50
            public string TipoMovimento;          // String * 3
            public string TipoProcessamento;      // String * 1
            public string ContaBanqueiro;         // String * 35
            public string Filler;                 // String * 93

            public void Inicializa()
            {
                this.TipoRemessa = string.Empty;
                this.CodigoEmpresa = string.Empty;
                this.SiglaSistema = string.Empty;
                this.IdentificadorMovimento = string.Empty;
                this.CodigoMoeda = string.Empty;
                this.CodigoBanqueiroSwift = string.Empty;
                this.CodigoProduto = string.Empty;
                this.DataMovimento = string.Empty;
                this.CodigoReferenciaSwift = string.Empty;
                this.TipoEntradaSaida = string.Empty;
                this.ValorMovimento = string.Empty;
                this.NomeCliente = string.Empty;
                this.TipoMovimento = string.Empty;
                this.TipoProcessamento = string.Empty;
                this.ContaBanqueiro = string.Empty;
                this.Filler = string.Empty;
            }

            public override string ToString()
            {
                StringBuilder Concatena = new StringBuilder();

                Concatena.Append(this.TipoRemessa.Trim().PadRight(3, ' '));
                Concatena.Append(this.CodigoEmpresa.Trim().PadLeft(5, '0'));
                Concatena.Append(this.SiglaSistema.Trim().PadRight(3, ' '));
                Concatena.Append(this.IdentificadorMovimento.Trim().PadRight(25, ' '));
                Concatena.Append(this.CodigoMoeda.Trim().PadRight(4, ' '));
                Concatena.Append(this.CodigoBanqueiroSwift.Trim().PadRight(30, ' '));
                Concatena.Append(this.CodigoProduto.Trim().PadLeft(4, '0'));
                Concatena.Append(this.DataMovimento.Trim().PadRight(8, ' '));
                Concatena.Append(this.CodigoReferenciaSwift.Trim().PadRight(16, ' '));
                Concatena.Append(this.TipoEntradaSaida.Trim().PadRight(1, '0'));
                Concatena.Append(this.ValorMovimento.Trim().PadLeft(19, '0'));
                Concatena.Append(this.NomeCliente.Trim().PadRight(50, ' '));
                Concatena.Append(this.TipoMovimento.Trim().PadLeft(3, '0'));
                Concatena.Append(this.TipoProcessamento.Trim().PadRight(1, '0'));
                Concatena.Append(this.ContaBanqueiro.Trim().PadRight(35, ' '));
                Concatena.Append(this.Filler.Trim().PadRight(93, ' '));

                return Concatena.ToString();
            }
        }
        #endregion

        #region <<< Construtores >>>
        public GestaoCaixa()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~GestaoCaixa()
        {
            this.Dispose();
        }
        #endregion

        #region <<< Enviar Previsao Item Caixa >>>
        public string EnviarPrevisaoItemCaixa(XmlDocument remessa, bool estorno, string dataLiquidacaoPJ)
        {
            A8NET.Data.DAO.GestaoCaixaDAO GestaoCaixaDAO;
            A8NET.Data.DAO.TextXmlDAO TextXmlDAO;
            StringBuilder Protocolo = new StringBuilder();
            A8NET.Comum.Comum.EnumTipoMovimentoPJ TipoMovimentoPJ;
            A8NET.Comum.Comum.EnumTipoMovimento TipoMovimento;
            string Previsao;
            int CodigoTextXML;

            try 
	        {
                // Complementar as informacoes necessarias para enviar a operacao para o PJ
                Protocolo.Append("1004     ");  // Tipo de Mensagem
                Protocolo.Append("A8 ");   // Sigla do Sistema de Origem
                Protocolo.Append("PJ ");   // Sigla do Sistema de Destino
                Protocolo.Append(remessa.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText.PadLeft(5, '0'));  // Codigo da Empresa

                // Verifica Tipo Movimento
                if (estorno)
                {
                    TipoMovimentoPJ = A8NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto;
                    TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.EstornoPrevisto;
                }
                else
                {
                    TipoMovimentoPJ = A8NET.Comum.Comum.EnumTipoMovimentoPJ.Previsto;
                    TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.Previsto;
                }

                Previsao = MontarPrevisaoItemCaixa(remessa, TipoMovimentoPJ, dataLiquidacaoPJ);

                // Grava Mensagem Base64
                TextXmlDAO = new A8NET.Data.DAO.TextXmlDAO();
                CodigoTextXML = TextXmlDAO.InserirBase64(string.Concat(Protocolo.ToString(), Previsao));

                // Grava Historico
                GestaoCaixaDAO = new A8NET.Data.DAO.GestaoCaixaDAO();
                GestaoCaixaDAO.Inserir(Convert.ToDecimal(remessa.DocumentElement.SelectSingleNode("//NU_SEQU_OPER_ATIV").InnerText),
                                       TipoMovimento, CodigoTextXML);

                return string.Concat(Protocolo.ToString(), Previsao);

	        }
	        catch
	        {
                return string.Empty;
	        }
        }
        #endregion

        #region <<< Enviar Maiores Valores >>>
        public string EnviarMaioresValores(XmlDocument remessa, A8NET.Comum.Comum.EnumTipoMovimentoPJ tipoMovimentoPJ, string dataLiquidacaoPJ)
        {
            A8NET.Data.DAO.GestaoCaixaDAO GestaoCaixaDAO;
            A8NET.Data.DAO.TextXmlDAO TextXmlDAO;
            StringBuilder Protocolo = new StringBuilder();
            A8NET.Comum.Comum.EnumTipoMovimento TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.RealizadoConfirmado;
            string MaioresValores;
            int CodigoTextXML;

            try 
	        {
                // Complementar as informacoes necessarias para enviar a operacao para o PJ
                Protocolo.Append("1001     ");  // Tipo de Mensagem
                Protocolo.Append("A8 ");   // Sigla do Sistema de Origem
                Protocolo.Append("PJ ");   // Sigla do Sistema de Destino
                Protocolo.Append(remessa.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText.PadLeft(5, '0'));  // Codigo da Empresa

                MaioresValores = MontarMaioresValores(remessa, tipoMovimentoPJ, dataLiquidacaoPJ);

                // Verifica Tipo Movimento
                switch (tipoMovimentoPJ)
                {
                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.Previsto:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.Previsto;
                        break;

                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.Realizado:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.RealizadoConfirmado;
                        break;

                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.EstornoPrevisto;
                        break;

                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoRealizado:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.EstornoRealizadoSolicitado;
                        break;

                    default:
                        break;
                }

                if (MaioresValores != null)
                {
                    // Grava Mensagem Base64
                    TextXmlDAO = new A8NET.Data.DAO.TextXmlDAO();
                    CodigoTextXML = TextXmlDAO.InserirBase64(string.Concat(Protocolo.ToString(), MaioresValores));

                    // Grava Historico
                    GestaoCaixaDAO = new A8NET.Data.DAO.GestaoCaixaDAO();
                    GestaoCaixaDAO.Inserir(Convert.ToDecimal(remessa.DocumentElement.SelectSingleNode("//NU_SEQU_OPER_ATIV").InnerText),
                                           TipoMovimento, CodigoTextXML);

                    return string.Concat(Protocolo.ToString(), MaioresValores);
                }
                else
                {
                    return string.Empty;
                }

	        }
	        catch
	        {
                return string.Empty;
	        }
        }
        #endregion

        #region <<< Enviar Moeda Estrangeira >>>
        public string EnviarMoedaEstrangeira(XmlDocument remessa, A8NET.Comum.Comum.EnumTipoMovimentoPJ tipoMovimentoPJ, string dataLiquidacaoPJ)
        {
            A8NET.Data.DAO.GestaoCaixaDAO GestaoCaixaDAO;
            A8NET.Data.DAO.TextXmlDAO TextXmlDAO;
            StringBuilder Protocolo = new StringBuilder();
            A8NET.Comum.Comum.EnumTipoMovimento TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.Previsto;
            string MoedaEstrangeira;
            int CodigoTextXML;

            try
            {
                // Complementar as informacoes necessarias para enviar a operacao para o PJ
                Protocolo.Append("1014     ");  // Tipo de Mensagem
                Protocolo.Append("A8 ");   // Sigla do Sistema de Origem
                Protocolo.Append("PJ ");   // Sigla do Sistema de Destino
                Protocolo.Append(remessa.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText.PadLeft(5, '0'));  // Codigo da Empresa

                MoedaEstrangeira = MontarMoedaEstrangeira(remessa, tipoMovimentoPJ, dataLiquidacaoPJ);

                // Grava Mensagem Base64
                TextXmlDAO = new A8NET.Data.DAO.TextXmlDAO();
                CodigoTextXML = TextXmlDAO.InserirBase64(string.Concat(Protocolo.ToString(), MoedaEstrangeira));

                // Verifica Tipo Movimento
                switch (tipoMovimentoPJ)
                {
                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.Previsto:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.Previsto;
                        break;

                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.Realizado:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.RealizadoConfirmado;
                        break;

                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.EstornoPrevisto;
                        break;

                    case A8NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoRealizado:
                        TipoMovimento = A8NET.Comum.Comum.EnumTipoMovimento.EstornoRealizadoSolicitado;
                        break;

                    default:
                        break;
                }

                // Grava Historico
                GestaoCaixaDAO = new A8NET.Data.DAO.GestaoCaixaDAO();
                GestaoCaixaDAO.Inserir(Convert.ToDecimal(remessa.DocumentElement.SelectSingleNode("//NU_SEQU_OPER_ATIV").InnerText),
                                       TipoMovimento, CodigoTextXML);

                return string.Concat(Protocolo.ToString(), MoedaEstrangeira);

            }
            catch
            {
                return string.Empty;
            }
        }
        #endregion

        #region <<< Montar Previsao Item Caixa >>>
        private string MontarPrevisaoItemCaixa(XmlDocument remessa, A8NET.Comum.Comum.EnumTipoMovimentoPJ tipoMovimentoPJ, string dataLiquidacaoPJ)
        {
            A8NET.Data.DAO.GestaoCaixaDAO GestaoCaixaDAO;
            string DataOperacao = string.Empty;
            string IndicadorDebitoCredito = string.Empty;
            string IdentificadorRemessaPJ = string.Empty;
            string[] Valor;

            try
            {
                GestaoCaixaDAO = new A8NET.Data.DAO.GestaoCaixaDAO();
                DadosRemessaMovimento = new EstruturaRemessaMovimento();
                DadosRemessaMovimento.Inicializa();

                // Preenche Data da Operacao
                if (dataLiquidacaoPJ == null)
                {
                    if (remessa.DocumentElement.SelectSingleNode("//DT_OPER_ATIV") == null)
                    {
                        DataOperacao = remessa.DocumentElement.SelectSingleNode("//DT_MESG").InnerText;
                    }
                    else
                    {
                        DataOperacao = remessa.DocumentElement.SelectSingleNode("//DT_OPER_ATIV").InnerText;
                    }
                }
                else
                {
                    DataOperacao = dataLiquidacaoPJ;
                }

                // Busca Identificador de Remessa do PJ
                IdentificadorRemessaPJ = GestaoCaixaDAO.ObterIdentificadorRemessaPJ(string.Format("{0:yyyyMMdd}", DateTime.Today));

                // Preenche Estrutura do Layout do PJ para Item de Caixa
                DadosRemessaMovimento.TipoRemessa = "100";
                DadosRemessaMovimento.CodigoRemessa = IdentificadorRemessaPJ;
                DadosRemessaMovimento.DataRemessa = string.Format("{0:yyyyMMdd}", DateTime.Today);
                DadosRemessaMovimento.HoraRemessa = string.Format("{0:HHmm}", DateTime.Now);
                DadosRemessaMovimento.CodigoEmpresa = remessa.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText;
                DadosRemessaMovimento.SiglaSistema = remessa.DocumentElement.SelectSingleNode("//SG_SIST_ORIG").InnerText;
                DadosRemessaMovimento.CodigoMoeda = A8NET.Comum.Comum.CodigoMoeda;
                DadosRemessaMovimento.TipoCaixa = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoCaixaPJ.Futuro);
                DadosRemessaMovimento.CodigoProduto = remessa.DocumentElement.SelectSingleNode("//CO_PROD").InnerText;
                DadosRemessaMovimento.CodigoLocalLiquidacao = remessa.DocumentElement.SelectSingleNode("//CO_LOCA_LIQU").InnerText;
                DadosRemessaMovimento.TipoMovimento = Convert.ToString((int)tipoMovimentoPJ);
                DadosRemessaMovimento.DataMovimento = DataOperacao;
                DadosRemessaMovimento.HoraMovimento = string.Format("{0:HHmm}", DateTime.Now);
                DadosRemessaMovimento.TipoProcessamento = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoProcessamentoPJ.OnLine);
                DadosRemessaMovimento.TipoEnvio = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoEnvioPJ.Parcial);

                // Tipo Conta
                if (remessa.DocumentElement.SelectSingleNode("//TP_CONT") != null)
                {
                    DadosRemessaMovimento.TipoConta = remessa.DocumentElement.SelectSingleNode("//TP_CONT").InnerText;
                }

                // Segmento
                if (remessa.DocumentElement.SelectSingleNode("//CO_SEGM") != null)
                {
                    DadosRemessaMovimento.CodigoSegmento = remessa.DocumentElement.SelectSingleNode("//CO_SEGM").InnerText;
                }

                // Evento Financeiro
                if (remessa.DocumentElement.SelectSingleNode("//CO_EVEN_FINC") != null)
                {
                    DadosRemessaMovimento.EventoFinanceiro = remessa.DocumentElement.SelectSingleNode("//CO_EVEN_FINC").InnerText;
                }

                // Indexador
                if (remessa.DocumentElement.SelectSingleNode("//CO_INDX") != null)
                {
                    DadosRemessaMovimento.CodigoIndexador = remessa.DocumentElement.SelectSingleNode("//CO_INDX").InnerText;
                }

                // Tipo Debito/Credito
                if (remessa.DocumentElement.SelectSingleNode("//TP_OPER_CAMB").InnerText == "V")
                {
                    IndicadorDebitoCredito = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoEntradaSaida.Entrada);
                }
                else
                {
                    IndicadorDebitoCredito = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoEntradaSaida.Saida);
                }
                DadosRemessaMovimento.TipoEntradaSaida = IndicadorDebitoCredito;

                // Valor
                Valor = remessa.DocumentElement.SelectSingleNode("//VA_MOED_NACIO").InnerText.Split(',');
                if (Valor.Length == 1)
                {
                    DadosRemessaMovimento.ValorMovimento = string.Concat(Valor[0].Trim().ToString(), "00");
                    DadosRemessaMovimento.ValorContabil = DadosRemessaMovimento.ValorMovimento;
                }
                else
                {
                    DadosRemessaMovimento.ValorMovimento = string.Concat(Valor[0].Trim().ToString(), Valor[1].Trim().PadRight(2, '0').ToString());
                    DadosRemessaMovimento.ValorContabil = DadosRemessaMovimento.ValorMovimento;
                }

                return DadosRemessaMovimento.ToString();

            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixa.MontarPrevisaoItemCaixa() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< Montar Maiores Valores >>>
        private string MontarMaioresValores(XmlDocument remessa, A8NET.Comum.Comum.EnumTipoMovimentoPJ tipoMovimentoPJ, string dataLiquidacaoPJ)
        {
            A8NET.Data.DAO.GestaoCaixaDAO GestaoCaixaDAO;
            string DataOperacao = string.Empty;
            string IndicadorDebitoCredito = string.Empty;
            string IdentificadorRemessaPJ = string.Empty;
            string[] Valor;

            try
            {
                GestaoCaixaDAO = new A8NET.Data.DAO.GestaoCaixaDAO();
                DadosMaioresValores = new EstruturaMaioresValores();
                DadosMaioresValores.Inicializa();

                // Preenche Data da Operacao
                if (dataLiquidacaoPJ == null)
                {
                    if (remessa.DocumentElement.SelectSingleNode("//DT_OPER_ATIV") == null)
                    {
                        DataOperacao = remessa.DocumentElement.SelectSingleNode("//DT_MESG").InnerText;
                    }
                    else
                    {
                        DataOperacao = remessa.DocumentElement.SelectSingleNode("//DT_OPER_ATIV").InnerText;
                    }
                }
                else
                {
                    DataOperacao = dataLiquidacaoPJ;
                }

                // Busca Identificador de Remessa do PJ
                IdentificadorRemessaPJ = GestaoCaixaDAO.ObterIdentificadorRemessaPJ(string.Format("{0:yyyyMMdd}", DateTime.Today));

                // Preenche Estrutura do Layout do PJ para Maiores Valores
                DadosMaioresValores.TipoRemessa = "200";
                DadosMaioresValores.CodigoRemessa = IdentificadorRemessaPJ;
                DadosMaioresValores.DataRemessa = string.Format("{0:yyyyMMdd}", DateTime.Today);
                DadosMaioresValores.HoraRemessa = string.Format("{0:HHmm}", DateTime.Now);
                DadosMaioresValores.CodigoEmpresa = remessa.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText;
                DadosMaioresValores.SiglaSistema = remessa.DocumentElement.SelectSingleNode("//SG_SIST_ORIG").InnerText;
                DadosMaioresValores.CodigoMoeda = A8NET.Comum.Comum.CodigoMoeda;
                DadosMaioresValores.TipoCaixa = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoCaixaPJ.Futuro);
                DadosMaioresValores.CodigoProduto = remessa.DocumentElement.SelectSingleNode("//CO_PROD").InnerText;
                DadosMaioresValores.CodigoLocalLiquidacao = remessa.DocumentElement.SelectSingleNode("//CO_LOCA_LIQU").InnerText;
                DadosMaioresValores.TipoMovimento = Convert.ToString((int)tipoMovimentoPJ);
                DadosMaioresValores.DataMovimento = DataOperacao;
                DadosMaioresValores.HoraMovimento = string.Format("{0:HHmm}", DateTime.Now);
                DadosMaioresValores.TipoPessoa = "2";
                DadosMaioresValores.TipoProcessamento = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoProcessamentoPJ.OnLine);
                DadosMaioresValores.TipoEnvio = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoEnvioPJ.Parcial);

                // Tipo Conta
                if (remessa.DocumentElement.SelectSingleNode("//TP_CONT") != null)
                {
                    DadosMaioresValores.TipoConta = remessa.DocumentElement.SelectSingleNode("//TP_CONT").InnerText;
                }

                // Segmento
                if (remessa.DocumentElement.SelectSingleNode("//CO_SEGM") != null)
                {
                    DadosMaioresValores.CodigoSegmento = remessa.DocumentElement.SelectSingleNode("//CO_SEGM").InnerText;
                }

                // Evento Financeiro
                if (remessa.DocumentElement.SelectSingleNode("//CO_EVEN_FINC") != null)
                {
                    DadosMaioresValores.CodigoEventoFinanceiro = remessa.DocumentElement.SelectSingleNode("//CO_EVEN_FINC").InnerText;
                }

                // Indexador
                if (remessa.DocumentElement.SelectSingleNode("//CO_INDX") != null)
                {
                    DadosMaioresValores.CodigoIndexador = remessa.DocumentElement.SelectSingleNode("//CO_INDX").InnerText;
                }

                // Tipo Debito/Credito
                if (remessa.DocumentElement.SelectSingleNode("//TP_OPER_CAMB").InnerText == "V")
                {
                    IndicadorDebitoCredito = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoEntradaSaida.Entrada);
                }
                else
                {
                    IndicadorDebitoCredito = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoEntradaSaida.Saida);
                }
                DadosMaioresValores.TipoEntradaSaida = IndicadorDebitoCredito;

                // Valor
                Valor = remessa.DocumentElement.SelectSingleNode("//VA_MOED_NACIO").InnerText.Split(',');
                if (Valor.Length == 1)
                {
                    DadosMaioresValores.ValorMovimento = string.Concat(Valor[0].Trim().ToString(), "00");
                }
                else
                {
                    DadosMaioresValores.ValorMovimento = string.Concat(Valor[0].Trim().ToString(), Valor[1].Trim().PadRight(2, '0').ToString());
                }

                // Nome Contraparte
                if (remessa.DocumentElement.SelectSingleNode("//NM_CLIE_MOED_ESTR") != null)
                {
                    DadosMaioresValores.NomeCliente = remessa.DocumentElement.SelectSingleNode("//NM_CLIE_MOED_ESTR").InnerText;
                }

                if (GestaoCaixaDAO.VerificarEnvioMaioresValores(int.Parse(remessa.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText),
                                                                int.Parse(remessa.DocumentElement.SelectSingleNode("//CO_PROD").InnerText),
                                                                decimal.Parse(remessa.DocumentElement.SelectSingleNode("//VA_MOED_NACIO").InnerText)))
                {
                    return DadosMaioresValores.ToString();
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixa.MontarMaioresValores() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< Montar Moeda Estrangeira >>>
        private string MontarMoedaEstrangeira(XmlDocument remessa, A8NET.Comum.Comum.EnumTipoMovimentoPJ tipoMovimentoPJ, string dataLiquidacaoPJ)
        {
            A8NET.Data.DAO.GestaoCaixaDAO GestaoCaixaDAO;
            XmlDocument xmlRemessaAux;
            string DataOperacao = string.Empty;
            string IndicadorDebitoCredito = string.Empty;
            string IdentificadorRemessaPJ = string.Empty;
            string[] Valor;
            int TipoMensagem;

            try
            {
                GestaoCaixaDAO = new A8NET.Data.DAO.GestaoCaixaDAO();
                DadosMoedaEstrangeira = new EstruturaMoedaEstrangeira();
                DadosMoedaEstrangeira.Inicializa();
                xmlRemessaAux = new XmlDocument();
                xmlRemessaAux.LoadXml(remessa.OuterXml);

                // Carrega variáveis
                TipoMensagem = int.Parse(xmlRemessaAux.SelectSingleNode("//TP_MESG").InnerText);

                // Layout de Registro de Arbitragem com Grupo de Contratação, deixa no XML apenas o GrupoContratacao da ponta de Dólar
                if (TipoMensagem == (int)Comum.Comum.EnumTipoMensagem.RegistroOperacaoArbitragem)
                {
                    TratarOperacoesComGruposContratacao(ref xmlRemessaAux);
                }

                // Preenche Data da Operacao
                if (dataLiquidacaoPJ == null)
                {
                    if (xmlRemessaAux.DocumentElement.SelectSingleNode("//DT_OPER_ATIV") == null)
                    {
                        DataOperacao = xmlRemessaAux.DocumentElement.SelectSingleNode("//DT_MESG").InnerText;
                    }
                    else
                    {
                        DataOperacao = xmlRemessaAux.DocumentElement.SelectSingleNode("//DT_OPER_ATIV").InnerText;
                    }
                }
                else
                {
                    DataOperacao = dataLiquidacaoPJ;
                }

                // Busca Identificador de Remessa do PJ
                IdentificadorRemessaPJ = GestaoCaixaDAO.ObterIdentificadorRemessaPJ(string.Format("{0:yyyyMMdd}", DateTime.Today));

                // Preenche Estrutura do Layout do PJ para Moeda Estrangeira
                DadosMoedaEstrangeira.TipoRemessa = "250";
                DadosMoedaEstrangeira.CodigoEmpresa = xmlRemessaAux.DocumentElement.SelectSingleNode("//CO_EMPR").InnerText;
                DadosMoedaEstrangeira.SiglaSistema = xmlRemessaAux.DocumentElement.SelectSingleNode("//SG_SIST_ORIG").InnerText;
                DadosMoedaEstrangeira.IdentificadorMovimento = IdentificadorRemessaPJ;
                
                // Obter Código Moeda
                if (xmlRemessaAux.DocumentElement.SelectSingleNode("//CO_MOED_ISO") != null)
                {
                    DadosMoedaEstrangeira.CodigoMoeda = xmlRemessaAux.DocumentElement.SelectSingleNode("//CO_MOED_ISO").InnerText;
                }
                else
                {
                    DadosMoedaEstrangeira.CodigoMoeda = " ".PadRight(4); 
                }
                                
                DadosMoedaEstrangeira.CodigoBanqueiroSwift = xmlRemessaAux.DocumentElement.SelectSingleNode("//CO_BANQ_SWIFT").InnerText;
                DadosMoedaEstrangeira.CodigoProduto = xmlRemessaAux.DocumentElement.SelectSingleNode("//CO_PROD_MOED_ESTR").InnerText;
                DadosMoedaEstrangeira.DataMovimento = DataOperacao;
                DadosMoedaEstrangeira.NomeCliente = xmlRemessaAux.DocumentElement.SelectSingleNode("//NM_CLIE_MOED_ESTR").InnerText;
                DadosMoedaEstrangeira.ContaBanqueiro = xmlRemessaAux.DocumentElement.SelectSingleNode("//NR_CNTA_BANQ").InnerText;
                DadosMoedaEstrangeira.TipoMovimento = Convert.ToString((int)tipoMovimentoPJ);
                DadosMoedaEstrangeira.TipoProcessamento = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoProcessamentoPJ.OnLine);

                // Tipo Debito/Credito
                if (xmlRemessaAux.DocumentElement.SelectSingleNode("//TP_OPER_CAMB").InnerText == "C")
                {
                    IndicadorDebitoCredito = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoDebitoCreditoPJME.Credito);
                }
                else
                {
                    IndicadorDebitoCredito = Convert.ToString((int)A8NET.Comum.Comum.EnumTipoDebitoCreditoPJME.Debito);
                }
                DadosMoedaEstrangeira.TipoEntradaSaida = IndicadorDebitoCredito;

                // Valor
                Valor = xmlRemessaAux.DocumentElement.SelectSingleNode("//VA_MOED_ESTRG").InnerText.Split(',');
                if (Valor.Length == 1)
                {
                    DadosMoedaEstrangeira.ValorMovimento = string.Concat(Valor[0].Trim().ToString(), "00");
                }
                else
                {
                    DadosMoedaEstrangeira.ValorMovimento = string.Concat(Valor[0].Trim().ToString(), Valor[1].Trim().PadRight(2, '0').ToString());
                }

                return DadosMoedaEstrangeira.ToString();

            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixa.MontarMoedaEstrangeira() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< Tratar Operacoes Com Grupos de Contratacao >>>
        private XmlDocument TratarOperacoesComGruposContratacao(ref XmlDocument remessa)
        {
            string CodigoMensagemSPB = string.Empty;
            string XMLAux;
            decimal ValorME = 0;

            try
            {
                if (remessa.SelectNodes("//GR_CONTR").Count > 0)
                {
                    foreach (XmlNode Node in remessa.SelectNodes("//GR_CONTR"))
                    {
                        if (Node.SelectSingleNode("CO_MOED_ISO") != null)
                        {
                            if (Node.SelectSingleNode("CO_MOED_ISO").InnerText != "USD")
                            {
                                Node.ParentNode.RemoveChild(Node);
                            }
                        }
                    }
                }
                else
                {
                    #region <<< Trata caso haja Mensagem SPB appendada no XML >>>
                    
                    CodigoMensagemSPB = remessa.SelectSingleNode("//CodMsg").InnerText;

                    if (CodigoMensagemSPB != string.Empty)
                    {

                        remessa.LoadXml(remessa.OuterXml.Replace("_" + CodigoMensagemSPB.Trim(), string.Empty));

                        if (remessa.SelectNodes("//Grupo_Contr").Count > 0)
                        {
                            XMLAux = remessa.OuterXml;
                            XMLAux = XMLAux.Replace("CodMoedaISO", "CO_MOED_ISO");
                            XMLAux = XMLAux.Replace("VlrME", "VA_MOED_ESTRG");
                            XMLAux = XMLAux.Replace("TpOpCAM", "TP_OPER_CAMB");
                            remessa.LoadXml(XMLAux);

                            foreach (XmlNode Node in remessa.SelectNodes("//Grupo_Contr"))
                            {
                                if (Node.SelectSingleNode("CO_MOED_ISO") != null)
                                {
                                    if (Node.SelectSingleNode("CO_MOED_ISO").InnerText != "USD")
                                    {
                                        Node.ParentNode.RemoveChild(Node);
                                    }
                                }
                            }
                        }
                    }
                    
                    #endregion
                }

                return remessa;

            }
            catch
            {
                return remessa;
            }
        }
        #endregion

    }
}
