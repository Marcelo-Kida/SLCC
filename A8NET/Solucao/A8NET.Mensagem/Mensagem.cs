using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using A8NET.Data.DAO;
using A8NET.Data;
using System.Data.OracleClient;
using System.Data;

namespace A8NET.Mensagem
{
    public abstract class Mensagem
    {
        #region <<< Variaveis >>>
        Data.DsParametrizacoes _DsCache;
        protected MensagemSpbDAO _MensagemSpbDATA;
        protected HistoricoSituacaoMensagemDAO _HistoricoMensagemDATA;
        protected OperacaoDAO _OperacaoDATA;
        protected OperacaoDAO.EstruturaOperacao _ParametroOPER;
        protected ConciliacaoDAO _ConciliacaoDATA;
        protected OperacaoMensagemInternaDAO _OperacaoMensagemInternaDATA = new OperacaoMensagemInternaDAO();
        protected HistoricoSituacaoOperacaoDAO _HistSituacaoOperacaoDATA = new HistoricoSituacaoOperacaoDAO();
        #endregion

        #region <<< Propriedades Protegidas >>>
        protected Data.DsParametrizacoes DataSetCache
        {
            get
            {
                return _DsCache;
            }
            set
            {
                _DsCache = value;
            }
        }
        #endregion

        #region <<< Metodos Abstract >>>
        public abstract void ProcessaMensagem(string nomeFila, string mensagemRecebida);
        #endregion

        #region <<< Construtor >>>
        public Mensagem()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion
        }
        #endregion

        #region <<< SelecionarTextoBase64 >>>
        public string SelecionarTextoBase64(int codigoTexto)
        {
            TextXmlDAO TextDAO = new TextXmlDAO();
            StringBuilder RetornoAppend = new StringBuilder();
            try
            {
                OracleDataReader DrDados = TextDAO.SelecionarTextoXML(codigoTexto);

                while (DrDados.Read())
                {
                    RetornoAppend.Append(DrDados["TX_XML"].ToString());
                }
                return Comum.Comum.Base64Decode(RetornoAppend.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Mensagem.SelecionarTextoBase64() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterOperacaoXML >>>
        public XmlDocument ObterOperacaoXML(int codigoOperacao)
        {
            DataTable DtOper = new DataTable();
            OperacaoDAO OperDAO = new OperacaoDAO();
            XmlDocument XmlOperacao = new XmlDocument();
            DataRow LinhaOPER;
            string ColunasNaoTratarLoop = @"|HO_ENVI_MESG_SPB|VA_OPER_ATIV|TP_CPRO_OPER_ATIV|TP_CPRO_RETN_OPER_ATIV|";
            try
            {
                DtOper = OperDAO.ObterOperacaoXML(codigoOperacao);

                if(DtOper.Rows.Count ==0) return new XmlDocument();
                LinhaOPER = DtOper.Rows[0];

                XmlOperacao.LoadXml(this.SelecionarTextoBase64(int.Parse(LinhaOPER["CO_TEXT_XML"].ToString())));

                // criar as colunas que não estão no xml
                foreach (DataColumn coluna in DtOper.Columns)
                {
                    if (ColunasNaoTratarLoop.IndexOf("|" + coluna + "|") != -1) continue;

                    if (XmlOperacao.DocumentElement.SelectSingleNode(coluna.ColumnName) == null)
                    {
                        Comum.Comum.AppendNode(ref XmlOperacao, "MESG", coluna.ColumnName.Trim(), "");
                    }
                    XmlOperacao.DocumentElement.SelectSingleNode(coluna.ColumnName).InnerXml = this.FormatarValor(coluna.ColumnName,
                        LinhaOPER[coluna.ColumnName].ToString());
                }

                if (!LinhaOPER.IsNull("HO_ENVI_MESG_SPB") && XmlOperacao.DocumentElement.SelectSingleNode("HO_AGND").InnerXml == "0")
                {
                    XmlOperacao.DocumentElement.SelectSingleNode("HO_AGND").InnerXml = DateTime.Parse( LinhaOPER["HO_ENVI_MESG_SPB"].ToString()).ToString("HHmm");
                }

                if (!LinhaOPER.IsNull("VA_OPER_ATIV") && XmlOperacao.DocumentElement.SelectSingleNode("HO_AGND") != null)
                {
                    XmlOperacao.DocumentElement.SelectSingleNode("VA_OPER_ATIV").InnerXml = this.ValorToXml(LinhaOPER["VA_OPER_ATIV"].ToString());
                }

                if (!LinhaOPER.IsNull("TP_CPRO_OPER_ATIV") && XmlOperacao.DocumentElement.SelectSingleNode("TP_CPRO_OPER_ATIV") != null)
                {
                    XmlOperacao.DocumentElement.SelectSingleNode("TP_CPRO_OPER_ATIV").InnerXml = LinhaOPER["TP_CPRO_OPER_ATIV"].ToString();
                }

                if (!LinhaOPER.IsNull("TP_CPRO_RETN_OPER_ATIV") && XmlOperacao.DocumentElement.SelectSingleNode("TP_CPRO_RETN_OPER_ATIV") != null)
                {
                    XmlOperacao.DocumentElement.SelectSingleNode("TP_CPRO_RETN_OPER_ATIV").InnerXml = LinhaOPER["TP_CPRO_RETN_OPER_ATIV"].ToString();
                }

                return XmlOperacao;
   
            }
            catch (Exception ex)
            {
                throw new Exception("Mensagem.ObterOperacaoXML() - " + ex.ToString());
            }
            finally
            {
                OperDAO = null;
            }
        }
        #endregion

        #region <<< AlterStatusOperacao >>>
        public void AlterarStatusOperacao(int seqOperacao, int statusOperacao, int tipoJustificativaSituacao, int tipoAcao)
        {
            OperacaoDAO OperacaDATA = new OperacaoDAO();
            HistoricoSituacaoOperacaoDAO HistoricoDAO = new HistoricoSituacaoOperacaoDAO();
            HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao ParametroHistorico = new HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao();
            OperacaoDAO.EstruturaOperacao ParametroOPER = new OperacaoDAO.EstruturaOperacao();
            //tipoJustificativaSituacao = 0;
            
            try
            {
                // preencher os dados da operação q devem ser alterados
                ParametroOPER.NU_SEQU_OPER_ATIV = seqOperacao;
                ParametroOPER.CO_ULTI_SITU_PROC = statusOperacao;
                ParametroOPER.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroOPER.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                // altera o status da operação
                OperacaDATA.AtualizarStatus(ParametroOPER);

                // preencher os dados do histórico q devem ser alterados
                ParametroHistorico.CO_ETCA_USUA_ATLZ = Comum.Comum.NomeMaquina;
                ParametroHistorico.CO_USUA_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroHistorico.CO_SITU_PROC = statusOperacao;
                ParametroHistorico.NU_SEQU_OPER_ATIV = seqOperacao;
                ParametroHistorico.DH_SITU_ACAO_OPER_ATIV = (HistoricoDAO.ObteveDataGravacao ? HistoricoDAO.DataGravacao : HistoricoDAO.ObterDataGravacao(seqOperacao, statusOperacao));
                if (tipoJustificativaSituacao != 0) ParametroHistorico.TP_JUST_SITU_PROC = tipoJustificativaSituacao;
                if (tipoAcao != 0) ParametroHistorico.TP_ACAO_OPER_ATIV = tipoAcao;
                
                // inserir o histórico na base de dados.
                HistoricoDAO.Inserir(ParametroHistorico);
            }
            catch (Exception ex)
            {
                return;
                //throw new Exception("Mensagem.AlterStatusOperacao() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< AlterarStatusMensagemSPB >>>
        /// <summary>
        /// Alterar o status da mensagem, na alteração além de alterar a tabela TB_MESG_RECB_ENVI_SPB inclui um novo registro de histórico
        /// </summary>
        public void AlterarStatusMensagemSPB(ref MensagemSpbDAO.EstruturaMensagemSPB parametro, Comum.Comum.EnumStatusMensagem status)
        {
            try
            {
                MensagemSpbDAO _MensagemSpbDATA = new MensagemSpbDAO();
                HistoricoSituacaoMensagemDAO.EstruturaHistoricoSituacaoMsg ParametroHistorico = new HistoricoSituacaoMensagemDAO.EstruturaHistoricoSituacaoMsg();
                HistoricoSituacaoMensagemDAO HistoricoMensagemDATA = new HistoricoSituacaoMensagemDAO();

                // Alterar status da Mensagem
                parametro.CO_ULTI_SITU_PROC = (int)status;
                parametro.DH_ULTI_ATLZ = DateTime.Now;

                // Salvar MensagemSPB
                _MensagemSpbDATA.Atualizar(parametro);

                // Preencher os dados de histórico
                ParametroHistorico.NU_CTRL_IF = parametro.NU_CTRL_IF;
                ParametroHistorico.NU_SEQU_CNTR_REPE = parametro.NU_SEQU_CNTR_REPE;
                ParametroHistorico.CO_SITU_PROC = parametro.CO_ULTI_SITU_PROC;
                ParametroHistorico.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                ParametroHistorico.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroHistorico.DH_REGT_MESG_SPB = parametro.DH_REGT_MESG_SPB;
                ParametroHistorico.DH_SITU_ACAO_MESG_SPB = (HistoricoMensagemDATA.ObteveDataGravacao ? HistoricoMensagemDATA.DataGravacao : HistoricoMensagemDATA.ObterDataGravacao(parametro.NU_CTRL_IF.ToString()));

                // Salvar histórico
                HistoricoMensagemDATA.Inserir(ParametroHistorico);
            }
            catch (Exception ex)
            {
                return;
            }
        }
        #endregion

        #region <<< AlterarMensagemSPB >>>
        /// <summary>
        /// Alterar qualquer campo da tabela TB_MESG_RECB_ENVI_SPB
        /// </summary>
        public void AlterarMensagemSPB(ref MensagemSpbDAO.EstruturaMensagemSPB parametro)
        {
            try
            {
                MensagemSpbDAO _MensagemSpbDATA = new MensagemSpbDAO();

                parametro.DH_ULTI_ATLZ = DateTime.Now;
                _MensagemSpbDATA.Atualizar(parametro);
            }
            catch (Exception ex)
            {
                return;
            }
        }
        #endregion

        #region >>> ObterTipoAcaoEnvioMensagemSPB >>>
        /// <summary>
        /// Obtêm o TipoAcao da Operacao, de acordo com o Tipo de Operacao.
        /// Retorna 0(zero) caso não encontre.
        /// </summary>
        /// <param name="udtOperacaoRecebida">udt da Operação que está sendo processada</param>
        /// <returns></returns>
        public int ObterTipoAcaoEnvioMensagemSPB(int tipoOperacao)
        {

            try
            {

                switch (tipoOperacao)
                {
                    case (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemTelaCega:
                        return (int)Comum.Comum.EnumTipoAcao.EnviadaCAM0006;
                    case (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemCamara:
                        return (int)Comum.Comum.EnumTipoAcao.EnviadaCAM0009;
                    case (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancarioEletronico:
                        return (int)Comum.Comum.EnumTipoAcao.EnviadaCAM0054;
                    default:
                        return 0;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.ObterTipoAcaoEnvioMensagemSPB() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< FormatarValor >>>
        public string FormatarValor(string nomeColuna, string valor)
        {
            string NuloPreencherBranco = @"CO_MESG_SPB_REGT_OPER|CO_MESG_SPB_REGT_OPER|NU_COMD_OPER|DT_OPER_ATIV|CO_PRAC|CO_MOED_ESTR|PE_TAXA_NEGO|
                                         CO_SISB_COTR|TP_CNAL_VEND|CD_SUB_PROD|CD_ASSO_CAMB|NU_ATIV_MERC";

            if (valor == null || valor == "")
            {
                if (NuloPreencherBranco.IndexOf("|" + nomeColuna + "|") != -1)
                    return "";
                else
                    return "0";
            }
            else
            {
                if (nomeColuna.Substring(0, 2) == "DT")
                {
                    return DateTime.Parse(valor).ToString("yyyyMMdd");
                }
                else if (nomeColuna.Substring(0, 2) == "DH")
                {
                    return DateTime.Parse(valor).ToString("yyyyMMddhhmmss");
                }
                else
                {
                    return valor;
                }
            }
        }
        #endregion

        #region <<< ValorToXml >>>
        protected string ValorToXml(string valor)
        {
            return valor.Replace(System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyGroupSeparator, "")
                .Replace(System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator, ",");
        }
        #endregion

        #region >>> ObterTipoBackOffice >>>
        protected static int ObterTipoBackOffice(string codigoMensagem)
        {
            switch (codigoMensagem.Trim())
            {
                case "CAM0005R2":
                case "CAM0015":
                case "CAM0055":
                    return (int)Comum.Comum.EnumTipoBackOffice.Todos;

                case "CAM0006R2":
                case "CAM0007R2":
                case "CAM0008R2":
                case "CAM0009R2":
                case "CAM0010R2":
                    return (int)Comum.Comum.EnumTipoBackOffice.Tesouraria;

                default:
                    if (codigoMensagem.Substring(0, 3).Equals("CAM"))
                    {
                        return (int)Comum.Comum.EnumTipoBackOffice.Comex;
                    }
                    else
                    {
                        return (int)Comum.Comum.EnumTipoBackOffice.Todos;
                    }
            }
        }
        #endregion

        #region <<< MontarHeaderMensagemNZ() >>>
        /// <summary></summary>
        /// <param name="parametroOPER"></param>
        /// <param name="entidadeMensagem"></param>
        /// <returns></returns>
        protected string MontarHeaderMensagemNZ(string codigoMensagem, int codigoEmpresa, string numeroControleIF)
        {
            string HeaderMensagemNZ = string.Empty;
            string DominioSistema = string.Empty;

            try
            {
                if (codigoMensagem.Substring(0, 3) == "CAM" || codigoMensagem.Substring(0, 3) == "CCR")
                {
                    DominioSistema = "2"; //MES01
                }
                else
                {
                    DominioSistema = "0"; //SPB01
                }

                HeaderMensagemNZ = string.Concat(
                                   "A8 ", //SiglaSistemaEnviouNZ
                                   codigoMensagem.ToString().PadRight(9), //CodigoMensagem
                                   numeroControleIF.ToString().PadRight(20), //ControleRemessaNZ
                                   DateTime.Today.ToString("yyyyMMdd"), //DataRemessa
                                   codigoEmpresa.ToString().PadLeft(5, '0'), //CodigoEmpresa
                                   "00790", //CodigoMoeda
                                   2, //FormatoMensagem
                                   "0".PadLeft(50, '0'), //AssinaturaInterna
                                   "A8 ", //SiglaSistemaLegadoOrigem
                                   "A8" + numeroControleIF.Substring(numeroControleIF.Length - 6, 6), //ReferenciaContabil
                                   "0".PadLeft(15, '0'), //BancoAgencia
                                   "000001", //QuantidadeMensagem
                                   "0".PadLeft(23, '0'), //NuOP
                                   "0".PadLeft(42, '0'), //FILLER
                                   DominioSistema, //DominioSistema (definido pelo Bacen, cada grupo de mensagens pertencem a um Domínio Sistema)
                                   "0".ToString().PadLeft(1, '0') //FILLER_1
                                   );

                return HeaderMensagemNZ;

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.MontarHeaderMensagemNZ() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> AtualizaStatusOperacao >>>
        /// <summary>
        /// Por default esta função retorna =True, há apenas alguns cenários de exceção que podem
        /// ocorrer no processamento de algumas mensagens (CAM0054R1/CAM0006R1/CAM0007R1/CAM0007R2) devido as mensagens SPB
        /// serem recebidas do Bacen na ordem invertida ao que está definido nos Fluxos. Por exemplo a BMC0005 deveria ser
        /// a última mensagem a ser processada em diversos fluxos, no entanto algumas vezes ocorre dela chegar chegar
        /// ao SLCC antes de outras Mensagens intermediárias do fluxo.
        /// </summary>
        /// <param name="codigoMensagemSPB">Código da Mensagem SPB</param>
        /// <param name="numeroSequenciaOperacao">Número de Sequência da Operação</param>
        /// <returns></returns>
        protected bool AtualizaStatusOperacao(string codigoMensagemSPB, long numeroSequenciaOperacao)
        {
            bool AtualizaStatusOperacao;
            ConciliacaoDAO ConciliacaoDATA = new ConciliacaoDAO();
            MensagemSpbDAO MensagemSPBData = new MensagemSpbDAO();
            OperacaoDAO OperacaoDATA = new OperacaoDAO();

            try
            {

                // Por Default esta função retorna TRUE
                AtualizaStatusOperacao = true;

                // ao processar CAM54R1/CAM6R1/CAM7R1/CAM7R2/CAM8R2, se já houver uma BMC0005 conciliada à Operação, então não será necessário 
                // atualizar o status da Operação, pois a BMC0005 já a deixou no status final
                if (codigoMensagemSPB == "CAM0054R1"
                 || codigoMensagemSPB == "CAM0006R1"
                 || codigoMensagemSPB == "CAM0007R1"
                 || codigoMensagemSPB == "CAM0007R2"
                 || codigoMensagemSPB == "CAM0008R2")
                {
                    ConciliacaoDATA.SelecionarOperacaoConciliada(long.Parse(numeroSequenciaOperacao.ToString()));
                    if (ConciliacaoDATA.Itens.Length > 0)
                    {
                        for (int row_conciliacao = 0; row_conciliacao < ConciliacaoDATA.Itens.Length; row_conciliacao++)
                        {
                            MensagemSPBData.SelecionarMensagensPorControleIF(ConciliacaoDATA.Itens[row_conciliacao].NU_CTRL_IF.ToString());
                            if (MensagemSPBData.Itens.Length > 0)
                            {
                                for (int row_mensagemspb = 0; row_mensagemspb < MensagemSPBData.Itens.Length; row_mensagemspb++)
                                {
                                    if (MensagemSPBData.Itens[row_mensagemspb].CO_MESG_SPB.ToString().Trim() == "BMC0005".ToString())
                                    {
                                        AtualizaStatusOperacao = false;
                                        break;
                                    }
                                }
                                if (AtualizaStatusOperacao == false) break;
                            }
                        }
                    }
                }

                return AtualizaStatusOperacao;

            }
            catch (Exception ex)
            {
                throw new Exception("Mensagem.AtualizaStatusOperacao() - " + ex.ToString());
            }
        }
        #endregion

    }
}
