using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using A8NET.Comum;
using System.Data.OracleClient;
using A8NET.Data;
using System.Xml;
using A8NET.Data.DAO;
using A8NET.Historico;
using A8NET.ConfiguracaoMQ;
using System.Configuration;
using System.IO;
using System.Collections;
using A8NET.Mensagem;

namespace A8NET.Mensagem.Operacao
{
    public class Operacao : Mensagem, IDisposable
    {

        #region <<< Variaveis >>>
        protected Data.DsParametrizacoes _DsCache;
        DsTB_OPER_ATIV _DsTB_OPER_ATIV = new DsTB_OPER_ATIV();
        protected Data.DAO.OperacaoDAO _OperacaoDATA = new Data.DAO.OperacaoDAO();
        protected OperacaoDAO.EstruturaOperacao _ParametroOperacao = new OperacaoDAO.EstruturaOperacao();
        private ConciliacaoDAO _ConciliacaoDATA = new ConciliacaoDAO();
        protected int _CodigoTextXML;
        protected HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao ParametroHistoricoSituacaoOperacao = new HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao();
        protected DsTB_MESG_RECB_ENVI_SPB _DsTB_MESG_RECB_ENVI_SPB = new DsTB_MESG_RECB_ENVI_SPB();
        protected int _RetornoConciliacao = 0;
        protected DataTable _DtErro;
        protected XmlDocument _XMLOperacao;
        private long _NumeroSequenciaOperacao = 0;
        udtOperacao _UdtOperacaoRecebida = new udtOperacao();
        Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = null;
        int _TipoSolicitacao;
        string _MensagemSPB = string.Empty;
        bool _ExisteIdOperacao;
        string _EventoProcessamento;
        private Conciliacao _ConciliacaoBO = new Conciliacao();
        private TextXmlDAO _TextXMLData = new TextXmlDAO();
        private OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna _ParametroOperacaoMensagemInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
        #endregion

        #region <<< Construtores >>>
        public Operacao(Data.DsParametrizacoes dataSetCache)
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion

            #region >>> Cria DataTable de Erro >>>
            _DtErro = new DataTable();
            _DtErro.Columns.Add("CD_ERRO");
            _DtErro.Columns.Add("DS_ERRO");
            _DtErro.Columns.Add("CM_ERRO");
            #endregion

            base.DataSetCache = dataSetCache;
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~Operacao()
        {
            this.Dispose();
        }
        #endregion

        #region <<< Fields >>>
        public object NU_SEQU_OPER_ATIV = 0;
        public object TP_OPER = DBNull.Value;
        public object CO_LOCA_LIQU = DBNull.Value;
        public object TP_LIQU_OPER_ATIV = DBNull.Value;
        public object CO_EMPR = DBNull.Value;
        public object CO_USUA_CADR_OPER = DBNull.Value;
        public object HO_ENVI_MESG_SPB = DBNull.Value;
        public object CO_OPER_ATIV = DBNull.Value;
        public object NU_COMD_OPER = DBNull.Value;
        public object NU_COMD_OPER_RETN = DBNull.Value;
        public object DT_OPER_ATIV = DBNull.Value;
        public object DT_OPER_ATIV_RETN = DBNull.Value;
        public object CO_VEIC_LEGA = DBNull.Value;
        public object SG_SIST = DBNull.Value;
        public object CO_CNTA_CUTD_SELIC_VEIC_LEGA = DBNull.Value;
        public object CO_CNPJ_CNPT = DBNull.Value;
        public object CO_CNTA_CUTD_SELIC_CNPT = DBNull.Value;
        public object NO_CNPT = DBNull.Value;
        public object IN_OPER_DEBT_CRED = DBNull.Value;
        public object NU_ATIV_MERC = DBNull.Value;
        public object DE_ATIV_MERC = DBNull.Value;
        public object PU_ATIV_MERC = DBNull.Value;
        public object QT_ATIV_MERC = DBNull.Value;
        public object IN_ENTR_SAID_RECU_FINC = DBNull.Value;
        public object DT_VENC_ATIV = DBNull.Value;
        public object VA_OPER_ATIV = DBNull.Value;
        public object VA_OPER_ATIV_REAJ = DBNull.Value;
        public object DT_LIQU_OPER_ATIV = DBNull.Value;
        public object TP_CPRO_OPER_ATIV = DBNull.Value;
        public object TP_CPRO_RETN_OPER_ATIV = DBNull.Value;
        public object IN_DISP_CONS = DBNull.Value;
        public object IN_ENVI_PREV_SIST_PJ = DBNull.Value;
        public object IN_ENVI_RELZ_SIST_PJ = DBNull.Value;
        public object IN_ENVI_PREV_SIST_A6 = DBNull.Value;
        public object IN_ENVI_RELZ_SIST_A6 = DBNull.Value;
        public object CO_ULTI_SITU_PROC = DBNull.Value;
        public object TP_ACAO_OPER_ATIV_EXEC = DBNull.Value;
        public object NU_COMD_ACAO_EXEC = DBNull.Value;
        public object IN_ENTR_MANU = DBNull.Value;
        public object NU_PRTC_OPER_LG = DBNull.Value;
        public object NU_SEQU_CNCL_OPER_ATIV_MESG = DBNull.Value;
        public object NU_CTRL_MESG_SPB_ORIG = DBNull.Value;
        public object PE_TAXA_NEGO = DBNull.Value;
        public object CO_TITL_CUTD = DBNull.Value;
        public object CO_OPER_CETIP = DBNull.Value;
        public object CO_ISPB_BANC_LIQU_CNPT = DBNull.Value;
        public object DH_ULTI_ATLZ = DBNull.Value;
        public object CO_ETCA_TRAB_ULTI_ATLZ = DBNull.Value;
        public object CO_USUA_ULTI_ATLZ = DBNull.Value;
        public object TP_IF_CRED_DEBT = DBNull.Value;
        public object CO_AGEN_COTR = DBNull.Value;
        public object NU_CC_COTR = DBNull.Value;
        public object PZ_DIAS_RETN_OPER_ATIV = DBNull.Value;
        public object VA_OPER_ATIV_RETN = DBNull.Value;
        public object TP_CNPT = DBNull.Value;
        public object CO_CNPT_CAMR = DBNull.Value;
        public object CO_IDEF_LAST = DBNull.Value;
        public object CO_PARP_CAMR = DBNull.Value;
        public object TP_PGTO_LDL = DBNull.Value;
        public object CO_GRUP_LANC_FINC = DBNull.Value;
        public object CO_MOED_ESTR = DBNull.Value;
        public object CO_CNTR_SISB = DBNull.Value;
        public object CO_ISPB_IF_CNPT = DBNull.Value;
        public object CO_PRAC = DBNull.Value;
        public object VA_MOED_ESTR = DBNull.Value;
        public object DT_LIQU_OPER_ATIV_MOED_ESTR = DBNull.Value;
        public object CO_SISB_COTR = DBNull.Value;
        public object CO_CNAL_OPER_INTE = DBNull.Value;
        public object CO_SITU_PROC_MESG_SPB_RECB = DBNull.Value;
        public object TP_CNAL_VEND = DBNull.Value;
        public object CD_SUB_PROD = DBNull.Value;
        public object NR_IDEF_NEGO_BMC = DBNull.Value;
        public object TP_NEGO = DBNull.Value;
        public object CD_ASSO_CAMB = DBNull.Value;
        public object CD_OPER_ETRT = DBNull.Value;
        public object NR_CNPJ_CPF = DBNull.Value;
        #endregion

        #region <<< Enumeradores >>>
        public enum enumTipoNegociacaoInterbancaria
        {
            SemCamara = 1,
            InterbancarioEletronico = 2,
            SemTelaCega = 3
        }

        public enum enumTipoNegociacaoArbitragem
        {
            ParceiroPais = 1,
            ParceiroExteriorPaisIF = 2
        }

        public enum enumTipoNegociacaoCambial
        {
            Interbancaria = 1,
            Arbitragem = 2
        }
        #endregion
        
        #region <<< ProcessaMensagem >>>
        public override void ProcessaMensagem(string nomeFila, string mensagemRecebida)
        {
            try
            {
                _UdtOperacaoRecebida.Parse(mensagemRecebida);
                GerenciaProcessamento(_UdtOperacaoRecebida, mensagemRecebida);
            }
            catch
            {
                throw;
            }

        }
        #endregion

        #region <<< GerenciaProcessamento >>>
        /// <summary>
        /// Método trata a mensagem conforme o tipo da mensagem
        /// </summary>
        /// <param name="udtMsg">udt da mensagem (header, linha com os dados, xml original, fila)</param>
        public void GerenciaProcessamento(udtOperacao udtOperacaoRecebida, string mensagemRecebida)
        {
            bool MensagemOK = false;
            string MensagemRejeitada = "";

            try
            {
                _XMLOperacao = udtOperacaoRecebida.XmlOperacao;

                // Validar Remessa
                MensagemOK = ValidaRemessa(udtOperacaoRecebida);

                // Obter Mensagem SPB Registro Operacao e Tipo Solicitacao, informações necessárias tanto quando MensagemOK e tambem quando MensagemNOK
                _TipoSolicitacao = int.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString());
                if (udtOperacaoRecebida.TipoOperacao != 0)
                {
                    DataSetCache.TB_TIPO_OPER.DefaultView.RowFilter = string.Format("TP_OPER ='{0}'", udtOperacaoRecebida.TipoOperacao);
                    _MensagemSPB = DataSetCache.TB_TIPO_OPER.DefaultView[0]["CO_MESG_SPB_REGT_OPER"].ToString().Trim();
                }

                #region >>> Mensagem Validada >>>
                if (MensagemOK)
                {
                    //appenda tags no xml que serão necessárias no processamento adiante
                    Comum.Comum.AppendNode(ref _XMLOperacao, "MESG", "CO_MESG_SPB", _MensagemSPB);
                    Comum.Comum.AppendNode(ref _XMLOperacao, "MESG", "TP_OPER", udtOperacaoRecebida.TipoOperacao.ToString());
                    Comum.Comum.AppendNode(ref _XMLOperacao, "MESG", "TP_BKOF", udtOperacaoRecebida.TipoBackoffice.ToString());
                    udtOperacaoRecebida.Parse(udtOperacaoRecebida.XmlOperacao.InnerXml); // Executa Parse para atualizar o DataRow (RowOperacao) do udtOperacaoRecebida
                    
                    #region >>> Processa Operacao de acordo com o Tipo de Solicitação >>>

                    switch (_TipoSolicitacao)
                    {
                        case (int)Comum.Comum.enumTipoSolicitacao.Complementacao:
                            ProcessarComplementacao(udtOperacaoRecebida);
                            break;
                        case (int)Comum.Comum.enumTipoSolicitacao.Confirmacao:
                            ProcessarConfirmacao(udtOperacaoRecebida);
                            break;
                        case (int)Comum.Comum.enumTipoSolicitacao.Reativacao:
                            ProcessarReativacaoFluxo(udtOperacaoRecebida);
                            break;
                        case (int)Comum.Comum.enumTipoSolicitacao.Cancelamento:
                            ProcessarCancelamento(udtOperacaoRecebida);
                            break;
                    }

                    // verifica se ocorreu algum erro de negócio durante o processamento
                    if (_DtErro != null) if (_DtErro.Rows.Count > 0) MensagemOK = false;

                    #endregion
                }
                #endregion

                #region >>> Mensagem Rejeitada >>>
                if (!MensagemOK)
                {
                    // Grava Mensagem Rejeitada
                    SalvarMensagemRejeitada(mensagemRecebida);

                    // Monta Mensagem Rejeitada para Envio ao Sistema Legado
                    MensagemRejeitada = MontaMensagemRejeitada(mensagemRecebida, _MensagemSPB);

                    // Envia Sistema Legado
                    EnviarMensagemRejeicaoLegado(MensagemRejeitada);

                }
                #endregion
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        #endregion

        #region >>> ValidaRemessa() >>>
        protected bool ValidaRemessa(udtOperacao udtOperacaoRecebida)
        {
            bool ValidaDominio;
            int TipoMensagem;
            int TipoSolicitacao;
            int TipoNegociacaoInterbancaria = 0;
            DateTime DataOperacao;

            try
            {
                // Carrega Tipo de Mensagem
                TipoMensagem = int.Parse(udtOperacaoRecebida.RowOperacao["TP_MESG"].ToString());

                // Carrega Tipo de Solicitacao
                TipoSolicitacao = int.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString());

                // Carrega Tipo Negociacao Interbancaria
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_INTB") != null)
                {
                    TipoNegociacaoInterbancaria = int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_INTB").InnerText);
                }

                // Verifica se Identificador da Operacao ja existe (CO_OPER_ATIV e SG_SIST_ORIG)
                _ExisteIdOperacao = _OperacaoMensagemInternaDATA.VerificaIdentificadorOperacao(udtOperacaoRecebida.RowOperacao["CO_OPER_ATIV"].ToString(),
                                                                                              udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString());

                // De acordo com o Tipo de Mensagem e Tipo de Solicitacao envia uma Mensagem de Erro
                if (_ExisteIdOperacao)
                {
                    switch (TipoSolicitacao)
                    {
                        case (int)Comum.Comum.enumTipoSolicitacao.Confirmacao:
                            // Se Tipo Mensagem = 249 (Registro Operação Interbancária), Tipo Negociação Interbancária = 3 (Sem Tela Cega - PCAM383),
                            //  e Tipo Solicitação = 10 (Confirmação) então Identificador da Operação (CO_OPER_ATIV) pode ser repetido
                            if (TipoMensagem != (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria
                             || TipoNegociacaoInterbancaria != (int)enumTipoNegociacaoInterbancaria.SemTelaCega)
                            {
                                _DtErro.Rows.Add("4463", "Identificador da Operacao Sisbacen ja existe.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }
                            break;

                        case (int)Comum.Comum.enumTipoSolicitacao.Reativacao:
                        case (int)Comum.Comum.enumTipoSolicitacao.Cancelamento:
                            break;


                        default:
                            // Se Tipo Mensagem = 257 (Complemento de informações de contratação interbancário via leilão)
                            // então Identificador da Operação (CO_OPER_ATIV) pode ser repetido
                            if (TipoMensagem != (int)A8NET.Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao)
                            {
                                _DtErro.Rows.Add("4463", "Identificador da Operacao Sisbacen ja existe.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }
                            break;
                    }
                }
                else
                {
                    if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao
                    || TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Reativacao
                    || TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Cancelamento)
                    {
                        _DtErro.Rows.Add("4466", "Identificador da Operacao Sisbacen nao existe.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }

                // Valida Tipo Operacao
                udtOperacaoRecebida.TipoOperacao = ObterTipoOperacao(udtOperacaoRecebida);
                if (udtOperacaoRecebida.TipoOperacao == 0) //retorna erro caso não tenha conseguido obter TipoOperacao
                {
                    _DtErro.Rows.Add("3003", "Tipo de Operação para mensagem recebida do sistema BUS Inexistente.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                    return false;
                }
                
                // Valida Tipo Backoffice
                udtOperacaoRecebida.TipoBackoffice = ObterTipoBackoffice(udtOperacaoRecebida);
                if (udtOperacaoRecebida.TipoBackoffice == 0) //retorna erro caso não tenha conseguido obter TipoBackoffice
                {
                    _DtErro.Rows.Add("4468", "Codigo Veiculo Legal invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                    return false;
                }

                #region Valida Layouts (249, 253, 257, 258 e 260)

                // Verifica o Tipo de Solicitacao para Mensagens 249 e 253
                if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoArbitragem)
                {
                    if ((Comum.Comum.enumTipoSolicitacao)short.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString()) != Comum.Comum.enumTipoSolicitacao.Complementacao
                     && (Comum.Comum.enumTipoSolicitacao)short.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString()) != Comum.Comum.enumTipoSolicitacao.Cancelamento
                     && (Comum.Comum.enumTipoSolicitacao)short.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString()) != Comum.Comum.enumTipoSolicitacao.Reativacao
                     && (Comum.Comum.enumTipoSolicitacao)short.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString()) != Comum.Comum.enumTipoSolicitacao.Confirmacao)
                    {
                        _DtErro.Rows.Add("4464", "Tipo Solicitacao Sisbacen invalida.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }

                // Verifica o Tipo de Solicitacao para Mensagens 257, 258 e 260
                if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.IFInformaLiquidacaoInterbancaria
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.IFCamaraConsultaContratosCambioMercadoInterbancario)
                {
                    if ((Comum.Comum.enumTipoSolicitacao)short.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString()) != Comum.Comum.enumTipoSolicitacao.Complementacao)
                    {
                        _DtErro.Rows.Add("4464", "Tipo Solicitacao Sisbacen invalida.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }

                // Verifica o Local de Liquidação para Mensagens 249, 257 e 258
                if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria
                || TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao
                || TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.IFInformaLiquidacaoInterbancaria)
                {
                    if ((A8NET.Comum.Comum.EnumLocalLiquidacao)short.Parse(udtOperacaoRecebida.RowOperacao["CO_LOCA_LIQU"].ToString()) != A8NET.Comum.Comum.EnumLocalLiquidacao.STR
                    && (A8NET.Comum.Comum.EnumLocalLiquidacao)short.Parse(udtOperacaoRecebida.RowOperacao["CO_LOCA_LIQU"].ToString()) != A8NET.Comum.Comum.EnumLocalLiquidacao.BMC
                    && (A8NET.Comum.Comum.EnumLocalLiquidacao)short.Parse(udtOperacaoRecebida.RowOperacao["CO_LOCA_LIQU"].ToString()) != A8NET.Comum.Comum.EnumLocalLiquidacao.PAG)
                    {
                        _DtErro.Rows.Add("4469", "Codigo Local Liquidacao invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }

                // Verifica o Local de Liquidação para Mensagens 253, 260
                if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoArbitragem
                || TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.IFCamaraConsultaContratosCambioMercadoInterbancario)
                {
                    if ((A8NET.Comum.Comum.EnumLocalLiquidacao)short.Parse(udtOperacaoRecebida.RowOperacao["CO_LOCA_LIQU"].ToString()) != A8NET.Comum.Comum.EnumLocalLiquidacao.BMC)
                    {
                        _DtErro.Rows.Add("4469", "Codigo Local Liquidacao invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }

                if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoArbitragem
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.IFInformaLiquidacaoInterbancaria
                ||  TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.IFCamaraConsultaContratosCambioMercadoInterbancario)
                {
                    // Valida Data da Operacao
                    DataOperacao = A8NET.Comum.Comum.ConvertDtToDateTime(udtOperacaoRecebida.RowOperacao["DT_OPER_ATIV"].ToString());
                    if (DataOperacao != DateTime.Today)
                    {
                        _DtErro.Rows.Add("4467", "Data Operacao Sisbacen invalida.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Validacao de Todos os Dominios
                    for (int i = 0; i < udtOperacaoRecebida.RowOperacao.Table.Columns.Count; i++)
                    {
                        if (base.DataSetCache.TB_CTRL_DOMI.Select("NO_ATRB='" + udtOperacaoRecebida.RowOperacao.Table.Columns[i].ToString().ToUpper() + "'").Length > 0)
                        {
                            ValidaDominio = false;
                            foreach (DataRow Row in base.DataSetCache.TB_CTRL_DOMI.Select("NO_ATRB='" + udtOperacaoRecebida.RowOperacao.Table.Columns[i].ToString().ToUpper() + "'"))
                            {
                                if (udtOperacaoRecebida.RowOperacao[i].ToString().ToUpper() == string.Empty)
                                {
                                    ValidaDominio = true;
                                    break;
                                }
                                else
                                {
                                    if (Row["CO_DOMI"].ToString().Trim().ToUpper() == udtOperacaoRecebida.RowOperacao[i].ToString().ToUpper())
                                    {
                                        ValidaDominio = true;
                                        break;
                                    }
                                }
                            }
                            if (!ValidaDominio)
                            {
                                _DtErro.Rows.Add("4465", "Tag (" + udtOperacaoRecebida.RowOperacao.Table.Columns[i].ToString().ToUpper() + ") - Dominio Sisbacen invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                if (_DtErro.Rows.Count == 3)
                                {
                                    return false;
                                }
                            }
                        }
                    }
                    if (_DtErro != null) if (_DtErro.Rows.Count > 0) return false;
                    
                    // Valida Codigo Produto
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_CAML") != null)
                    {
                        if (int.Parse(udtOperacaoRecebida.RowOperacao["TP_NEGO_CAML"].ToString()) == (int)enumTipoNegociacaoCambial.Interbancaria)
                        {
                            if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_PROD") == null
                            ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_PROD").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4470", "Codigo Produto obrigatorio para Tipo Negociacao Cambial Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }
                            else
                            {
                                if (DataSetCache.TB_PRODUTO.Select("CO_PROD=" + udtOperacaoRecebida.RowOperacao["CO_PROD"].ToString() +
                                                                   " AND CO_EMPR_FUSI=2 AND DT_INIC_VIGE<='" + DataOperacao +
                                                                   "' AND (DT_FIM_VIGE IS NULL OR DT_FIM_VIGE>='" + DataOperacao + "')").Length == 0)
                                {
                                    _DtErro.Rows.Add("4471", "Codigo Produto invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                    return false;
                                }
                            }
                        }
                    }

                    // Valida Layout 249 - RegistroOperacaoInterbancaria
                    if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria)
                    {
                        if (!ValidaRemessaInterbancaria(udtOperacaoRecebida, DataOperacao))
                        {
                            return false;
                        }
                    }

                    // Valida Layout 253 - RegistroOperacaoArbitragem
                    if (TipoMensagem == (int)A8NET.Comum.Comum.EnumTipoMensagem.RegistroOperacaoArbitragem)
                    {
                        if (!ValidaRemessaArbitragem(udtOperacaoRecebida, DataOperacao))
                        {
                            return false;
                        }
                    }
                }
                #endregion

                return true;

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ValidaRemessa() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> ValidaRemessaInterbancaria() >>>
        protected bool ValidaRemessaInterbancaria(udtOperacao udtOperacaoRecebida, DateTime dataOperacao)
        {
            int TipoSolicitacao;
            int TipoNegociacaoInterbancaria;

            try
            {
                // Carrega Tipo de Solicitacao
                TipoSolicitacao = int.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString());

                // Carrega Tipo de Negociacao Interbancaria
                TipoNegociacaoInterbancaria = int.Parse(udtOperacaoRecebida.RowOperacao["TP_NEGO_INTB"].ToString());

                // Valida Codigo de Produto
                if (DataSetCache.TB_PRODUTO.Select("CO_PROD=" + udtOperacaoRecebida.RowOperacao["CO_PROD"].ToString() +
                                                   " AND CO_EMPR_FUSI=2 AND DT_INIC_VIGE<='" + dataOperacao +
                                                   "' AND (DT_FIM_VIGE IS NULL OR DT_FIM_VIGE>='" + dataOperacao + "')").Length == 0)
                {
                    _DtErro.Rows.Add("4471", "Codigo Produto invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                    return false;
                }

                // Valida Codigo de Produto Moeda Estrangeira
                if (DataSetCache.TB_PRODUTO.Select("CO_PROD=" + udtOperacaoRecebida.RowOperacao["CO_PROD_MOED_ESTR"].ToString() +
                                                   " AND CO_EMPR_FUSI=2 AND DT_INIC_VIGE<='" + dataOperacao +
                                                   "' AND (DT_FIM_VIGE IS NULL OR DT_FIM_VIGE>='" + dataOperacao + "')").Length == 0)
                {
                    _DtErro.Rows.Add("4506", "Codigo Produto Moeda Estrangeira invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                    return false;
                }

                #region Valida Tipo Solicitacao = 2 ou 3 (Complementacao ou Cancelamento)
                if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Complementacao
                 || TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Cancelamento)
                {

                    if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Complementacao)
                    {
                        // Valida Tipo Operacao Cambio
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_OPER_CAMB") == null
                        || udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_OPER_CAMB").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4523", "Tipo Operacao Cambio obrigatorio para este Tipo de Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }

                    if (TipoNegociacaoInterbancaria == (int)enumTipoNegociacaoInterbancaria.SemCamara)
                    {
                        // Valida CNPJ IF Compradora
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_COMPR") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_COMPR").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4472", "CNPJ IF Compradora obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida CNPJ IF Vendedora
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_VENC") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_VENC").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4473", "CNPJ IF Vendedora obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                        
                        // Valida Indicador Giro
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_GIRO") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_GIRO").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4475", "Indicador Giro obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Indicador Linha
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_LINHA") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_LINHA").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4476", "Indicador Linha obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Codigo Fato Natureza
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_FATO_NATU") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_FATO_NATU").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4477", "Codigo Fato Natureza obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Codigo Cliente Natureza
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CLIE_NATU") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CLIE_NATU").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4478", "Codigo Cliente Natureza obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Indicador Aval Natureza
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_AVAL_NATU") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_AVAL_NATU").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4479", "Indicador Aval Natureza obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Codigo Pagador ou Recebedor Exterior Natureza
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_PAGA_RECB_EXTE") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_PAGA_RECB_EXTE").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4480", "Codigo Pagador ou Recebedor Exterior Natureza obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Codigo Grupo Natureza
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_GRUP_NATU") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_GRUP_NATU").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4481", "Codigo Grupo Natureza obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Codigo Forma Entrega Moeda 
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_FORM_ENTR_MOED") == null
                        || udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_FORM_ENTR_MOED").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4505", "Codigo Forma Entrega Moeda obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }
                    else if (TipoNegociacaoInterbancaria == (int)enumTipoNegociacaoInterbancaria.InterbancarioEletronico)
                    {
                        // Valida CNPJ Base Camara
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_BASE_CAMR") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_BASE_CAMR").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4483", "CNPJ Base Camara obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Chave Associacao Cambio
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//ChACAM") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//ChACAM").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4484", "Chave Associacao Cambio obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida CNPJ IF
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4485", "CNPJ IF obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }
                    else if (TipoNegociacaoInterbancaria == (int)enumTipoNegociacaoInterbancaria.SemTelaCega)
                    {
                        // Valida CNPJ IF Compradora
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_COMPR") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_COMPR").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4472", "CNPJ IF Compradora obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida CNPJ IF Vendedora
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_VENC") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_VENC").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4473", "CNPJ IF Vendedora obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida CNPJ Camara
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_CAM") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_CAM").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4474", "CNPJ Camara obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Indicador Giro
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_GIRO") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_GIRO").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4475", "Indicador Giro obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        // Valida Indicador Linha
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_LINHA") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_LINHA").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4476", "Indicador Linha obrigatorio para este Tipo Negociacao Interbancaria.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }
                }
                #endregion

                #region Valida Tipo Solicitacao = 9 ou 10 (Reativacao de Fluxo ou Confirmacao)
                if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Reativacao
                 || TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Confirmacao)
                {

                    if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Confirmacao)
                    {
                        // Valida Tipo Operacao Cambio
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_OPER_CAMB") == null
                        || udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_OPER_CAMB").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4523", "Tipo Operacao Cambio obrigatorio para este Tipo de Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }
                    
                    // Valida Registro Operacao Cambial
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4487", "Registro Operacao Cambial obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Se TipoSolicitacao = Confirmacao, então valida se já chegou Mensagem R2 correspondente
                    if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Confirmacao)
                    {
                        if (_ConciliacaoDATA.ConciliarComMensagemSPB(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB").InnerText.Trim(),
                                                                    0,
                                                                    Comum.Comum.ConvertDtToDateTime(udtOperacaoRecebida.RowOperacao["DT_OPER_ATIV"].ToString()),
                                                                    Comum.Comum.EnumStatusMensagem.EnviadaLegado,
                                                                    _ConciliacaoBO.ObterCodigoMensagemSPBAConciliar((Comum.Comum.EnumTipoOperacao)udtOperacaoRecebida.TipoOperacao),
                                                                ref _DsTB_MESG_RECB_ENVI_SPB,
                                                                ref _RetornoConciliacao) == false)
                        {
                            _DtErro.Rows.Add("4507", "Não foi possível localizar a Mensagem R2 correspondente.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }
                }
                #endregion

                return true;

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ValidaRemessaInterbancaria() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> ValidaRemessaArbitragem() >>>
        protected bool ValidaRemessaArbitragem(udtOperacao udtOperacaoRecebida, DateTime dataOperacao)
        {
            int TipoSolicitacao;
            int TipoNegociacaoArbitragem;

            try
            {
                // Carrega Tipo de Solicitacao
                TipoSolicitacao = int.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString());

                // Carrega Tipo de Negociacao Arbitragem
                TipoNegociacaoArbitragem = int.Parse(udtOperacaoRecebida.RowOperacao["TP_NEGO_ARBT"].ToString());

                // Valida Codigo de Produto Moeda Estrangeira
                if (DataSetCache.TB_PRODUTO.Select("CO_PROD=" + udtOperacaoRecebida.RowOperacao["CO_PROD_MOED_ESTR"].ToString() +
                                                   " AND CO_EMPR_FUSI=2 AND DT_INIC_VIGE<='" + dataOperacao +
                                                   "' AND (DT_FIM_VIGE IS NULL OR DT_FIM_VIGE>='" + dataOperacao + "')").Length == 0)
                {
                    _DtErro.Rows.Add("4506", "Codigo Produto Moeda Estrangeira invalido.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                    return false;
                }

                #region Valida Tipo Solicitacao = 2 ou 3 (Complementacao ou Cancelamento)
                if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Complementacao
                || TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Cancelamento)
                {
                    if (TipoNegociacaoArbitragem == (int)enumTipoNegociacaoArbitragem.ParceiroPais)
                    {
                        // Valida CNPJ IF Parceira
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_PARC") == null
                        ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CNPJ_IF_PARC").InnerText.Trim() == string.Empty)
                        {
                            _DtErro.Rows.Add("4488", "CNPJ IF Parceira obrigatorio para este Tipo Negociacao Arbitragem.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }

                        //// Valida Numero Sequencia Instrucao Pagamento
                        //if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_INST_PAGTO") == null
                        //||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_INST_PAGTO").InnerText.Trim() == string.Empty)
                        //{
                        //    _DtErro.Rows.Add("4489", "Numero Sequencia Instrucao Pagamento obrigatorio para este Tipo Negociacao Arbitragem.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        //    return false;
                        //}
                    }

                    // Valida Grupo Contratacao
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//GR_CONTR") == null)
                    {
                        _DtErro.Rows.Add("4499", "Grupo Contratacao obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                    else
                    {
                        foreach (XmlNode Node in udtOperacaoRecebida.XmlOperacao.SelectNodes("//GR_CONTR"))
                        {
                            // Valida Tipo Operacao Cambio
                            if (Node.SelectSingleNode("//TP_OPER_CAMB") == null
                            ||  Node.SelectSingleNode("//TP_OPER_CAMB").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4500", "Tipo Operacao Cambio obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }

                            // Valida Codigo Moeda ISO
                            if (Node.SelectSingleNode("//CO_MOED_ISO") == null
                            ||  Node.SelectSingleNode("//CO_MOED_ISO").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4501", "Codigo Moeda ISO obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }

                            // Valida Valor Moeda Estrangeira
                            if (Node.SelectSingleNode("//VA_MOED_ESTRG") == null
                            ||  Node.SelectSingleNode("//VA_MOED_ESTRG").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4502", "Valor Moeda Estrangeira obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }

                            // Valida Taxa Cambio
                            if (Node.SelectSingleNode("//VA_TAXA_CAMB") == null
                            ||  Node.SelectSingleNode("//VA_TAXA_CAMB").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4503", "Taxa Cambio obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }

                            // Valida Data Entrega Moeda Estrangeira
                            if (Node.SelectSingleNode("//DT_ENTR_MOED_ESTR") == null
                            ||  Node.SelectSingleNode("//DT_ENTR_MOED_ESTR").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4504", "Data Entrega Moeda Estrangeira obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }

                            // Valida Codigo Forma Entrega Moeda
                            if (Node.SelectSingleNode("//CO_FORM_ENTR_MOED") == null
                            ||  Node.SelectSingleNode("//CO_FORM_ENTR_MOED").InnerText.Trim() == string.Empty)
                            {
                                _DtErro.Rows.Add("4505", "Codigo Forma Entrega Moeda obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                                return false;
                            }
                        }
                    }

                    // Valida Valor Moeda Nacional
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//VA_MOED_NACIO") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//VA_MOED_NACIO").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4490", "Valor Moeda Nacional obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Data Liquidacao
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_LIQU_OPER") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_LIQU_OPER").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4491", "Data Liquidacao obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Codigo Fato Natureza
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_FATO_NATU") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_FATO_NATU").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4492", "Codigo Fato Natureza obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Codigo Cliente Natureza
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CLIE_NATU") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_CLIE_NATU").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4493", "Codigo Cliente Natureza obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Indicador Aval Natureza
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_AVAL_NATU") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//IN_AVAL_NATU").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4494", "Indicador Aval Natureza obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Codigo Pagador ou Recebedor Exterior Natureza
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_PAGA_RECB_EXTE") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_PAGA_RECB_EXTE").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4495", "Codigo Pagador ou Recebedor Exterior Natureza obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Codigo Grupo Natureza
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_GRUP_NATU") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_GRUP_NATU").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4496", "Codigo Grupo Natureza obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }
                #endregion

                #region Valida Tipo Solicitacao = 9 (Reativacao de Fluxo)
                if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Reativacao)
                {
                    // Valida Registro Operacao Cambial
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4487", "Registro Operacao Cambial obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }
                }
                #endregion

                #region Valida Tipo Solicitacao = 10 (Confirmacao)
                if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Confirmacao)
                {
                    // Valida Registro Operacao Cambial
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4487", "Registro Operacao Cambial obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    // Valida Registro Operacao Cambial2
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB2") == null
                    ||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB2").InnerText.Trim() == string.Empty)
                    {
                        _DtErro.Rows.Add("4497", "Registro Operacao Cambial2 obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return false;
                    }

                    //// Valida Numero Sequencia Instrucao Pagamento
                    //if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_INST_PAGTO") == null
                    //||  udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_INST_PAGTO").InnerText.Trim() == string.Empty)
                    //{
                    //    _DtErro.Rows.Add("4498", "Numero Sequencia Instrucao Pagamento obrigatorio para este Tipo Solicitacao.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                    //    return false;
                    //}

                    // Se TipoSolicitacao = Confirmacao, então valida se já chegou Mensagem R2 correspondente
                    if (TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Confirmacao)
                    {
                        if (_ConciliacaoDATA.ConciliarComMensagemSPB(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB").InnerText.Trim(),
                                                                    long.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_REG_OPER_CAMB2").InnerText.Trim()),
                                                                    Comum.Comum.ConvertDtToDateTime(udtOperacaoRecebida.RowOperacao["DT_OPER_ATIV"].ToString()),
                                                                    Comum.Comum.EnumStatusMensagem.EnviadaLegado, 
                                                                    _ConciliacaoBO.ObterCodigoMensagemSPBAConciliar((Comum.Comum.EnumTipoOperacao)udtOperacaoRecebida.TipoOperacao),
                                                                ref _DsTB_MESG_RECB_ENVI_SPB,
                                                                ref _RetornoConciliacao) == false)          
                        {
                            _DtErro.Rows.Add("4507", "Não foi possível localizar a Mensagem R2 correspondente.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                            return false;
                        }
                    }

                }
                #endregion

                return true;

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ValidaRemessaArbitragem() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< ProcessarInclusao >>>
        public void ProcessarInclusao(ref udtOperacao udtOperacaoRecebida, ref long numeroSequenciaOperacao)
        {
            HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao ParametroHistoricoSituacaoOperacao = new HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao();
            XmlDocument XmlOperacaoAux = udtOperacaoRecebida.XmlOperacao;

            try
            {
                
                // Os layouts de Sisbacen Mercado Primário (168 à 248) são os únicos que não tem o campo CO_LOCA_LIQU.
                // Isto ocorreu por uma falha percebida muito tarde na definição dos layouts, portanto 
                // é necessário incluir CO_LOCA_LIQU = 22(CAM) nestes casos em que o XML está sem a tag CO_LOCA_LIQU.
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_LOCA_LIQU") == null)
                {
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "CO_LOCA_LIQU", Convert.ToString((int)Comum.Comum.EnumLocalLiquidacao.CAM));
                }

                #region <<< insere TB_OPER_ATIV >>>
                InserirTB_OPER_ATIV(udtOperacaoRecebida, ref _DsTB_OPER_ATIV);
                #endregion

                numeroSequenciaOperacao = long.Parse(_DsTB_OPER_ATIV.TB_OPER_ATIV[0].NU_SEQU_OPER_ATIV.ToString());

                #region <<< insere TB_TEXT_XML e TB_OPER_ATIV_MESG_INTE >>>
                // Armazena a mensagem original na tabela TB_TEXT_XML e TB_OPER_ATIV_MESG_INTE
                _CodigoTextXML = _TextXMLData.InserirBase64(udtOperacaoRecebida.XmlOperacao.InnerXml);
                _ParametroOperacaoMensagemInterna.NU_SEQU_OPER_ATIV = numeroSequenciaOperacao;
                _ParametroOperacaoMensagemInterna.DH_MESG_INTE = DateTime.Now;
                _ParametroOperacaoMensagemInterna.TP_MESG_INTE = int.Parse(udtOperacaoRecebida.RowOperacao["TP_MESG"].ToString()); 
                _ParametroOperacaoMensagemInterna.TP_SOLI_MESG_INTE = udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString(); 
                _ParametroOperacaoMensagemInterna.CO_TEXT_XML = _CodigoTextXML;
                _ParametroOperacaoMensagemInterna.TP_FORM_MESG_SAID = 0; //0 para R0's, mas analisa A7.TB_REGR_SIST_DEST para R1's e R2's;
                _OperacaoMensagemInternaDATA.Inserir(_ParametroOperacaoMensagemInterna);
                #endregion

                #region <<< insere TB_HIST_SITU_ACAO_OPER_ATIV >>>
                // insere o Historico de Situacao da Operacao
                ParametroHistoricoSituacaoOperacao.NU_SEQU_OPER_ATIV = numeroSequenciaOperacao;
                ParametroHistoricoSituacaoOperacao.DH_SITU_ACAO_OPER_ATIV = DateTime.Now; //DH_SITU_ACAO_OPER_ATIV = clsHistSituacaoOperacao.flObterDataGravacao
                ParametroHistoricoSituacaoOperacao.CO_SITU_PROC = _DsTB_OPER_ATIV.TB_OPER_ATIV[0].CO_ULTI_SITU_PROC;
                ParametroHistoricoSituacaoOperacao.CO_USUA_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroHistoricoSituacaoOperacao.CO_ETCA_USUA_ATLZ = Comum.Comum.NomeMaquina;
                _HistSituacaoOperacaoDATA.Inserir(ParametroHistoricoSituacaoOperacao);
                #endregion

                #region <<< appenda mais algumas tags que serão necessárias no procesamento adiante >>>
                //append NU_SEQU_OPER_ATIV
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_OPER_ATIV") == null)
                {
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "NU_SEQU_OPER_ATIV", _DsTB_OPER_ATIV.TB_OPER_ATIV[0].NU_SEQU_OPER_ATIV.ToString());
                }
                #endregion

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ProcessarInclusao() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ProcessarAlteracao >>>
        /// <summary>
        /// Método retorna a Operação a ser processada, e registra o envio dela pelo legado na tabela A8.TB_OPER_ATIV_MESG_INTE
        /// </summary>
        /// <param name="udtOperacaoRecebida">udt da Operação recebida</param>
        public void ProcessarAlteracao(ref udtOperacao udtOperacaoRecebida, ref long numeroSequenciaOperacao)
        {
            TextXmlDAO TextXMLData = new TextXmlDAO();
            DsTB_OPER_ATIV_MESG_INTE DsTB_OPER_ATIV_MESG_INTE = new DsTB_OPER_ATIV_MESG_INTE();
            HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao ParametroHistoricoSituacaoOperacao = new HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao();
            OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna ParametroOperacaoMensagemInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
            XmlDocument XmlOperacaoAux = udtOperacaoRecebida.XmlOperacao;
            XmlDocument XmlMensagemSPB = new XmlDocument();

            try
            {
                // Procura Operação pelo CO_OPER_ATIV
                _OperacaoDATA.ObterOperacao(udtOperacaoRecebida.RowOperacao["CO_OPER_ATIV"].ToString(), udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString());

                // Rejeita caso não encontre a Operação
                if (_OperacaoDATA == null)
                {
                    _DtErro.Rows.Add("4519", "Não foi possível efetuar a alteração. Operação não encontrada.", "A8NET.Mensagem.Operacao.Operacao.ProcessarAlteracao()");
                    return;
                }

                // Verifica se a Operação encontrada tem algum dos status esperados
                if (ValidaStatusOperacaoEncontrada(udtOperacaoRecebida, _OperacaoDATA.TB_OPER_ATIV) == false)
                {
                    _DtErro.Rows.Add("4520", "Não foi possível efetuar a alteração. Operação encontrada não está no status esperado.", "A8NET.Mensagem.Operacao.Operacao.ProcessarAlteracao()");
                    return;
                }

                numeroSequenciaOperacao = long.Parse(_OperacaoDATA.TB_OPER_ATIV.NU_SEQU_OPER_ATIV.ToString());

                #region <<< insere TB_TEXT_XML e TB_OPER_ATIV_MESG_INTE >>>
                // Armazena a mensagem original na tabela TB_TEXT_XML e TB_OPER_ATIV_MESG_INTE
                _CodigoTextXML = TextXMLData.InserirBase64(udtOperacaoRecebida.XmlOperacao.InnerXml);
                ParametroOperacaoMensagemInterna.NU_SEQU_OPER_ATIV = numeroSequenciaOperacao;
                ParametroOperacaoMensagemInterna.DH_MESG_INTE = DateTime.Now;
                ParametroOperacaoMensagemInterna.TP_MESG_INTE = int.Parse(udtOperacaoRecebida.RowOperacao["TP_MESG"].ToString());
                ParametroOperacaoMensagemInterna.TP_SOLI_MESG_INTE = udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString();
                ParametroOperacaoMensagemInterna.CO_TEXT_XML = _CodigoTextXML;
                ParametroOperacaoMensagemInterna.TP_FORM_MESG_SAID = 0;
                _OperacaoMensagemInternaDATA.Inserir(ParametroOperacaoMensagemInterna);
                #endregion

                #region >>> Appenda informações que serão necessárias nas remessas para o PJ >>>
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_OPER_ATIV") == null)
                {
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "NU_SEQU_OPER_ATIV", numeroSequenciaOperacao.ToString());
                }
                
                if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.CAMInformaContratacaoInterbancarioViaLeilao)
                {
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "VA_MOED_NACIO", _OperacaoDATA.TB_OPER_ATIV.VA_OPER_ATIV.ToString());
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "VA_MOED_ESTRG", _OperacaoDATA.TB_OPER_ATIV.VA_MOED_ESTR.ToString());

                    if (_OperacaoDATA.TB_OPER_ATIV.IN_OPER_DEBT_CRED.ToString() == "1")
                    {
                        Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "TP_OPER_CAMB", "V");
                    }
                    else // = "2"
                    {
                        Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "TP_OPER_CAMB", "C");
                    }

                    #region >>> Busca mensagem CAM0015 que originou esta Operação, para obter a CodMoedaISO que será necessário na integração PJ Moeda Estrangeira >>>
                    _MensagemSpbDATA = new MensagemSpbDAO();
                    _MensagemSpbDATA.SelecionarMensagensPorNumeroSequenciaOperacao(long.Parse(numeroSequenciaOperacao.ToString()), 
                                                                               ref _DsTB_MESG_RECB_ENVI_SPB);

                    XmlMensagemSPB.LoadXml(base.SelecionarTextoBase64(int.Parse(_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB.DefaultView[0]["CO_TEXT_XML"].ToString())));
                    
                    if (XmlMensagemSPB.DocumentElement.SelectSingleNode("//CodMoedaISO") != null)
                    {
                        Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "CO_MOED_ISO", XmlMensagemSPB.DocumentElement.SelectSingleNode("//CodMoedaISO").InnerText);
                    }
                    #endregion

                }
                #endregion

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.ProcessarAlteracao() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ProcessarComplementacao >>>
        public void ProcessarComplementacao(udtOperacao udtOperacaoRecebida)
        {
            try
            {

                _EventoProcessamento = "RecebimentoOperacao";

                // Verifica se Incluiu ou Altera Operação
                if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.CAMInformaContratacaoInterbancarioViaLeilao)
                {
                    ProcessarAlteracao(ref udtOperacaoRecebida, ref _NumeroSequenciaOperacao);
                }
                else
                {
                    ProcessarInclusao(ref udtOperacaoRecebida, ref _NumeroSequenciaOperacao);
                }

                #region >>> Se TipoOperacao = 239 (InformaLiquidacaoInterbancaria) então define EventoProcessamento específico >>>
                if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.InformaLiquidacaoInterbancaria)
                {
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_CAML") != null)
                    {
                        if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_CAML").InnerText.Trim() == ((int)enumTipoNegociacaoCambial.Interbancaria).ToString())
                        {
                            _EventoProcessamento = "RecebimentoOperacaoInterbanc";
                        }
                        else if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_CAML").InnerText.Trim() == ((int)enumTipoNegociacaoCambial.Arbitragem).ToString())
                        {
                            _EventoProcessamento = "RecebimentoOperacaoArbitragem";
                        }
                    }
                }
                #endregion

                // Gerencia Chamadas
                _OperacaoDATA.ObterOperacao((int)_NumeroSequenciaOperacao);
                GerenciarChamadas(udtOperacaoRecebida, _EventoProcessamento, _OperacaoDATA.TB_OPER_ATIV, null);

                // Disponibiliza operacao para consulta
                OperacaoDisponivelConsulta(_OperacaoDATA.TB_OPER_ATIV, Comum.Comum.EnumInidicador.Sim);

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ProcessarComplementacao() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ProcessarConfirmacao >>>
        public void ProcessarConfirmacao(udtOperacao udtOperacaoRecebida)
        {
            try
            {

                _EventoProcessamento = "RecebimentoConfirmacao";

                if (_ExisteIdOperacao == true)
                {
                    // Procura Operação pelo CO_OPER_ATIV
                    _OperacaoDATA.ObterOperacao(udtOperacaoRecebida.RowOperacao["CO_OPER_ATIV"].ToString(), udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString());

                    // Rejeita caso não encontre a Operação
                    if (_OperacaoDATA != null)
                    {
                        _NumeroSequenciaOperacao = long.Parse(_OperacaoDATA.TB_OPER_ATIV.NU_SEQU_OPER_ATIV.ToString());
                        _EventoProcessamento = "ReenvioOperacao";
                    }
                    else
                    {
                        _DtErro.Rows.Add("4466", "Identificador da Operacao Sisbacen nao existe.", "A8NET.Mensagem.Operacao.Operacao.ValidaRemessa()");
                        return;
                    }
                }
                else
                {
                    ProcessarInclusao(ref udtOperacaoRecebida, ref _NumeroSequenciaOperacao);
                }

                // Gerencia Chamadas
                _OperacaoDATA.ObterOperacao((int)_NumeroSequenciaOperacao);
                GerenciarChamadas(udtOperacaoRecebida, _EventoProcessamento, _OperacaoDATA.TB_OPER_ATIV, null);

                // Disponibiliza operacao para consulta
                OperacaoDisponivelConsulta(_OperacaoDATA.TB_OPER_ATIV, Comum.Comum.EnumInidicador.Sim);

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ProcessarConfirmacao() - " + ex.ToString());
            }

        }
        #endregion

        #region >>> InserirTB_OPER_ATIV() >>>
        /// <summary>
        /// Método para inserir na Table TB_OPER_ATIV as informações vindas das operações dos legados
        /// </summary>
        /// <param name="enumCodigosErro">codigo do erro</param>
        /// <param name="dataRefe">data de referencia da conta</param>
        /// <param name="descricaoConta">identificador do erro</param>
        /// <param name="descricaoErro">mensagem de erro</param>
        public void InserirTB_OPER_ATIV(udtOperacao udtOperacaoRecebida, ref DsTB_OPER_ATIV dsTB_OPER_ATIV)
        {
            DsTB_OPER_ATIV.TB_OPER_ATIVRow RowOperacao;

            try
            {
                RowOperacao = dsTB_OPER_ATIV.TB_OPER_ATIV.NewTB_OPER_ATIVRow();

                #region >>> Campos Default >>>
                RowOperacao.DH_ULTI_ATLZ = DateTime.Now;
                RowOperacao.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                RowOperacao.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                RowOperacao.NU_SEQU_OPER_ATIV = 0;
                RowOperacao.TP_OPER = decimal.Parse(udtOperacaoRecebida.RowOperacao["TP_OPER"].ToString());
                RowOperacao.CO_LOCA_LIQU = int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_LOCA_LIQU").InnerText);
                RowOperacao.CO_EMPR = int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_EMPR").InnerText);
                RowOperacao.CO_USUA_CADR_OPER = udtOperacaoRecebida.RowOperacao["CO_USUA_CADR_OPER"].ToString();
                RowOperacao.CO_OPER_ATIV = udtOperacaoRecebida.RowOperacao["CO_OPER_ATIV"].ToString();
                RowOperacao.DT_OPER_ATIV = Comum.Comum.ConvertDtToDateTime(udtOperacaoRecebida.RowOperacao["DT_OPER_ATIV"].ToString());
                RowOperacao.CO_VEIC_LEGA = udtOperacaoRecebida.RowOperacao["CO_VEIC_LEGA"].ToString();
                RowOperacao.SG_SIST = udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString();
                RowOperacao.IN_ENTR_SAID_RECU_FINC = 0; //0=vazio, ou seja, não é nem 1=Entrada e nem 2=Saída
                RowOperacao.IN_OPER_DEBT_CRED = 0;
                RowOperacao.VA_OPER_ATIV = 0;
                RowOperacao.IN_DISP_CONS = (int)Comum.Comum.EnumInidicador.Nao;
                RowOperacao.CO_ULTI_SITU_PROC = (decimal)Comum.Comum.EnumStatusOperacao.EmSer; //passar situacao como parametro...olhar como é no VB6...
                RowOperacao.IN_ENTR_MANU = (int)Comum.Comum.EnumInidicador.Nao;
                #endregion

                #region >>> Campos Dinâmicos, obtidos através do XML de De/Para dos layouts >>>
                try
                {
                    XmlDocument XmlLayoutOperacao = new XmlDocument();
                    string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\DePara_Layout_" + int.Parse(udtOperacaoRecebida.RowOperacao["TP_MESG"].ToString()) + ".xml";
                    object ConteudoTag = string.Empty;
                    object ConteudoCampoTabela = string.Empty;
                    ArrayList arrList = new ArrayList();

                    if (File.Exists(XmlAux) == true) //verifica se o layout tem XML de De/Para
                    {
                        XmlLayoutOperacao.Load(XmlAux);
                        XmlNodeList XmlNodeListCampos = XmlLayoutOperacao["CAMPOS"].SelectNodes("*");
                        foreach (XmlNode XmlAtributoCampo in XmlNodeListCampos)
                        {
                            //verifica se o campo_tabela já foi obtido
                            if (!arrList.Contains(XmlAtributoCampo.Attributes["campo_tabela"].Value.Trim().ToUpper()))
                            {

                                ConteudoTag = string.Empty;
                                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//" + XmlAtributoCampo.Attributes["campo_layout"].Value) != null)
                                {
                                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//" + XmlAtributoCampo.Attributes["campo_layout"].Value).InnerText != String.Empty)
                                    {
                                        ConteudoTag = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//" + XmlAtributoCampo.Attributes["campo_layout"].Value).InnerText;
                                    }
                                }
                                // verifica se a tag procurada tem conteudo
                                if ((string)ConteudoTag != string.Empty)
                                {
                                    //inclui o campo_tabela na lista de campos já obtidos
                                    arrList.Add(XmlAtributoCampo.Attributes["campo_tabela"].Value.Trim().ToUpper());

                                    //por default campo da tabela = conteudo original da tag
                                    ConteudoCampoTabela = ConteudoTag;

                                    //verifica se precisa fazer de/para do conteudo da tag
                                    if (XmlAtributoCampo.HasChildNodes == true)
                                    {
                                        XmlNodeList XmlNodeListDominios = XmlAtributoCampo["DOMINIOS"].SelectNodes("*");
                                        foreach (XmlNode XmlAtributoDominio in XmlNodeListDominios)
                                        {
                                            if (XmlAtributoDominio.Attributes["dominio_layout"].Value == (string)ConteudoTag)
                                            {
                                                ConteudoCampoTabela = XmlAtributoDominio.Attributes["dominio_tabela"].Value;
                                                break;
                                            }
                                        }
                                    }

                                    // Trata conteudo caso campo_tabela seja do tipo Data
                                    if (XmlAtributoCampo.Attributes["campo_tabela"].Value.Substring(0, 3) == "DT_")
                                    {
                                        ConteudoCampoTabela = Comum.Comum.ConvertDtToDateTime(ConteudoCampoTabela.ToString());
                                    }
                                
                                    RowOperacao[XmlAtributoCampo.Attributes["campo_tabela"].Value] = ConteudoCampoTabela;
 
                                }
                            }
                        }
                    }
                }
                catch { } //try-catch inserido para garantir que a Operação será gravada mesmo que ocorra erro na obtenção dos campos dinâmicos
                #endregion

                #region >>> persiste a operação >>>
                dsTB_OPER_ATIV.TB_OPER_ATIV.AddTB_OPER_ATIVRow(RowOperacao);
                _OperacaoDATA.PersisteTB_OPER_ATIV(ref dsTB_OPER_ATIV);
                #endregion
            }
            catch (Exception ex)
            {
                throw new Exception("InserirTB_OPER_ATIV()" + ex.ToString());
            }
            finally
            {
                RowOperacao = null;
            }
        }
        #endregion

        #region <<< GerenciarChamadas >>>
        /// <summary>
        /// O metodo basicamente altera o status da operação do R0 e envia a mensagem para o legado
        /// </summary>
        /// <param name="udtOperacaoRecebida">dado da operação R0</param>
        /// <param name="funcionalidade">mensagem recebida</param>
        /// <param name="entidadeOperacao">status da operação</param>
        /// <param name="statusRetornoLegado">indica o estorno</param>
        public void GerenciarChamadas(udtOperacao udtOperacaoRecebida, 
                                     string eventoProcessamento, 
                                     OperacaoDAO.EstruturaOperacao entidadeOperacao,
                                     Comum.Comum.EnumStatusOperacao? statusRetornoLegado)
        {
            OperacaoDAO OperacaDATA = new OperacaoDAO();
            RegraWorkflow RegraWorkflowBO = new RegraWorkflow(base.DataSetCache);
            GestaoCaixa GestaoCaixaBO = new GestaoCaixa();
            A8NET.Mensagem.Conciliacao ConciliacaoBO = new Conciliacao();
            MQConnector _MqConnector = null;
            XmlDocument XmlOperacaoAux = udtOperacaoRecebida.XmlOperacao;
            int _CodigoRetornoVerificacao = 0;
            string Mensagem = string.Empty;
            bool EnviarMensagem = false;
            string DataLiquidacaoPJ_MN = null;
            string DataLiquidacaoPJ_ME = null;
            RegraWorkflow.enumFuncaoSistema EnumFuncaoSistemaConfirmacao = new RegraWorkflow.enumFuncaoSistema();
            RegraWorkflow.enumFuncaoSistema EnumFuncaoSistemaConciliacao = new RegraWorkflow.enumFuncaoSistema();
            Comum.Comum.EnumStatusOperacao? StatusOperacaoSeConfirmacaoAutomatica = new A8NET.Comum.Comum.EnumStatusOperacao();
            Comum.Comum.EnumStatusOperacao? StatusOperacaoSeConciliacaoAutomatica = new A8NET.Comum.Comum.EnumStatusOperacao();
            Comum.Comum.EnumStatusOperacao StatusOperacao = new A8NET.Comum.Comum.EnumStatusOperacao();
            bool Estorno = false;
            Comum.Comum.EnumTipoMovimentoPJ TipoMovimentoPJPrevisto;
            Comum.Comum.EnumTipoMovimentoPJ TipoMovimentoPJRealizado;
            int CodigoTextXML_MensagemSPB = 0;

            try
            {
                
                #region >>> Obtem parametrização de processamento do TipoOperacao x Evento >>>
                DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.RowFilter = string.Format(@"NO_PROC_OPER_ATIV ='{0}' 
                                                                                        AND TP_OPER={1}", 
                                                                                        eventoProcessamento, 
                                                                                        udtOperacaoRecebida.TipoOperacao
                                                                                        );

                // aborta o processamento caso não haja parametrizaçao para o TipoOperacao x Evento
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.Count == 0) return;
                #endregion

                #region >>> Verifica se é Estorno e determina TiposMovimentoPJ >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ESTO_PJ_A6"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    Estorno = true;
                    TipoMovimentoPJPrevisto = Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto;
                    TipoMovimentoPJRealizado = Comum.Comum.EnumTipoMovimentoPJ.EstornoRealizado;
                }
                else
                {
                    Estorno = false;
                    TipoMovimentoPJPrevisto = Comum.Comum.EnumTipoMovimentoPJ.Previsto;
                    TipoMovimentoPJRealizado = Comum.Comum.EnumTipoMovimentoPJ.Realizado;
                }
                #endregion

                #region >>> Verifica regras especificas de alteração de status >>>
                if (VerificarSeRegraEspecificaAlteracaoStatus(udtOperacaoRecebida, ref StatusOperacao) == true)
                {
                    entidadeOperacao.CO_ULTI_SITU_PROC = (int)StatusOperacao;
                    base.AlterarStatusOperacao(int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString()), (int)StatusOperacao, 0, 0);
                }
                #endregion

                #region >>> Caso a Operação tenha alguma Mensagem R2/Informação associada, então appenda o XML dela na Operação >>>
                if (_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB != null) // este dataset _DsTB_MESG_RECB_ENVI_SPB é populado no começo do processamento, nas funções ValidaRemessa, ValidaRemessaArbitragem ou ValidaRemessaInterbancario
                {
                    if (_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB.Count == 1)
                    {
                        Comum.Comum.AppendNode (ref XmlOperacaoAux, "MESG", "MENSAGEM_SPB", base.SelecionarTextoBase64(int.Parse(_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB.DefaultView[0]["CO_TEXT_XML"].ToString())));
                    }
                }
                #endregion

                #region >>> Previsão PJ Moeda Nacional e Estrangeira >>>
                // verifica se deve enviar Previsão PJ Moeda Nacional
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_PREV_PJ"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    //obtem Data Liquidacao PJ Moeda Nacional
                    DataLiquidacaoPJ_MN = ObterDataLiquidacaoPJ(udtOperacaoRecebida, Comum.Comum.EnumTipoMoedaPJ.MoedaNacional);

                    using (_MqConnector = new MQConnector())
                    {
                        _MqConnector.MQConnect();
                        _MqConnector.MQQueueOpen("A7Q.E.ENTRADA", MQConnector.enumMQOpenOptions.PUT);
                    
                        // Envia Maiores Valores
                        Mensagem = GestaoCaixaBO.EnviarMaioresValores(udtOperacaoRecebida.XmlOperacao, TipoMovimentoPJPrevisto, DataLiquidacaoPJ_MN);
                        if (Mensagem != string.Empty)
                        {
                                _MqConnector.Message = Mensagem;
                                _MqConnector.MQPutMessage();
                        }

                        // Envia Item Caixa
                        Mensagem = GestaoCaixaBO.EnviarPrevisaoItemCaixa(udtOperacaoRecebida.XmlOperacao, Estorno, DataLiquidacaoPJ_MN);
                        if (Mensagem != string.Empty)
                        {
                            _MqConnector.Message = Mensagem;
                            _MqConnector.MQPutMessage();
                        }

                        _MqConnector.MQQueueClose();
                        _MqConnector.MQEnd();
                    }

                }

                // verifica se deve enviar Previsão PJ Moeda Estrangeira
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_PREV_PJ_ME"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    //obtem Data Liquidacao PJ Moeda Estrangeira
                    DataLiquidacaoPJ_ME = ObterDataLiquidacaoPJ(udtOperacaoRecebida, Comum.Comum.EnumTipoMoedaPJ.MoedaEstrangeira);
                    
                    // Envia Moeda Estrangeira
                    Mensagem = GestaoCaixaBO.EnviarMoedaEstrangeira(udtOperacaoRecebida.XmlOperacao, TipoMovimentoPJPrevisto, DataLiquidacaoPJ_ME);
                    if (Mensagem != string.Empty)
                    {
                        using (_MqConnector = new MQConnector())
                        {
                            _MqConnector.MQConnect();
                            _MqConnector.MQQueueOpen("A7Q.E.ENTRADA", MQConnector.enumMQOpenOptions.PUT);
                            _MqConnector.Message = Mensagem;
                            _MqConnector.MQPutMessage();
                            _MqConnector.MQQueueClose();
                            _MqConnector.MQEnd();
                        }
                    }
                }
                #endregion

                #region >>> Verifica regra de Confirmação Automática >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_VERI_REGR_CONF"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {

                    RegraWorkflowBO.ObterFuncaoSistemaStatus(udtOperacaoRecebida, RegraWorkflow.enumMacroFuncao.Confirmacao, ref EnumFuncaoSistemaConfirmacao, ref StatusOperacaoSeConfirmacaoAutomatica);

                    if (RegraWorkflowBO.VerificarRegraAutomatica(udtOperacaoRecebida, EnumFuncaoSistemaConfirmacao, ref _CodigoRetornoVerificacao) == true)
                    {
                        // altera o status da operação para ConcordanciaAutomatica
                        entidadeOperacao.CO_ULTI_SITU_PROC = (int)StatusOperacaoSeConfirmacaoAutomatica;
                        base.AlterarStatusOperacao(int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString()), (int)StatusOperacaoSeConfirmacaoAutomatica, 0, 0);
                    }
                    else
                    {
                        // atualiza justificativa TB_HIST_SITU_ACAO_OPER_ATIV
                        ParametroHistoricoSituacaoOperacao.NU_SEQU_OPER_ATIV = int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString());
                        ParametroHistoricoSituacaoOperacao.CO_SITU_PROC = int.Parse(entidadeOperacao.CO_ULTI_SITU_PROC.ToString());
                        ParametroHistoricoSituacaoOperacao.TP_JUST_SITU_PROC = _CodigoRetornoVerificacao;
                        _HistSituacaoOperacaoDATA.AtualizarJustificativa(ParametroHistoricoSituacaoOperacao);

                        return;
                    }

                }
                #endregion

                #region >>> Verifica regra de Conciliação Automática >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_VERI_REGR_CNCL"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {

                    RegraWorkflowBO.ObterFuncaoSistemaStatus(udtOperacaoRecebida, RegraWorkflow.enumMacroFuncao.Conciliacao, ref EnumFuncaoSistemaConciliacao, ref StatusOperacaoSeConciliacaoAutomatica);

                    if (RegraWorkflowBO.VerificarRegraAutomatica(udtOperacaoRecebida, EnumFuncaoSistemaConciliacao, ref _CodigoRetornoVerificacao) == false)
                    {
                        return; // encerra processamento devido parametrização Automática = Não
                    }

                    // verifica conciliação da Operação com Mensagem SPB. Caso a conciliação esteja OK, o status da Operação e da MensagemSPB já serão atualizados dentro de ConciliacaoBO.VerificaConciliacao()
                    if (ConciliacaoBO.VerificaConciliacao(entidadeOperacao,
                                                      ref XmlOperacaoAux,
                                                      ref _CodigoRetornoVerificacao) == false)
                    {
                        // atualiza TB_HIST_SITU_ACAO_OPER_ATIV caso haja Justificativa
                        if (_CodigoRetornoVerificacao != 0)
                        {
                            ParametroHistoricoSituacaoOperacao.NU_SEQU_OPER_ATIV = int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString());
                            ParametroHistoricoSituacaoOperacao.CO_SITU_PROC = int.Parse(entidadeOperacao.CO_ULTI_SITU_PROC.ToString());
                            ParametroHistoricoSituacaoOperacao.TP_JUST_SITU_PROC = _CodigoRetornoVerificacao;
                            _HistSituacaoOperacaoDATA.AtualizarJustificativa(ParametroHistoricoSituacaoOperacao);
                        }
                        return; // encerra processamento devido operação não conciliada
                    }
                }
                #endregion

                #region >>> Verifica regra de Liberação Automática >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_VERI_REGR_LIBE"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    if (LiberarAutomatico(udtOperacaoRecebida, entidadeOperacao) == false)
                    {
                        return;
                    }
                    
                }
                #endregion

                #region >>> Verifica regra de Envio Mensagem SPB >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_MESG_SPB"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    Mensagem = MontarMensagem(udtOperacaoRecebida, entidadeOperacao, ref EnviarMensagem);
                    if (EnviarMensagem == true)
                    {
                        using (_MqConnector = new MQConnector())
                        {
                            _MqConnector.MQConnect();
                            _MqConnector.MQQueueOpen("A7Q.E.ENTRADA_NET", MQConnector.enumMQOpenOptions.PUT);
                            _MqConnector.Message = Mensagem;
                            _MqConnector.MQPutMessage();
                            _MqConnector.MQQueueClose();
                            _MqConnector.MQEnd();
                        }

                    }
                }
                #endregion

                #region >>> Realizado PJ Moeda Nacional e Estrangeira >>>
                // verifica se deve enviar Realizado PJ Moeda Nacional
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_RELZ_PJ"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    // Envia Maiores Valores
                    Mensagem = GestaoCaixaBO.EnviarMaioresValores(udtOperacaoRecebida.XmlOperacao, TipoMovimentoPJRealizado, null);
                    if (Mensagem != string.Empty)
                    {
                        using (_MqConnector = new MQConnector())
                        {
                            _MqConnector.MQConnect();
                            _MqConnector.MQQueueOpen("A7Q.E.ENTRADA", MQConnector.enumMQOpenOptions.PUT);
                            _MqConnector.Message = Mensagem;
                            _MqConnector.MQPutMessage();
                            _MqConnector.MQQueueClose();
                            _MqConnector.MQEnd();
                        }
                    }
                }

                // verifica se deve enviar Realizado PJ Moeda Estrangeira
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_RELZ_PJ_ME"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    // Envia Moeda Estrangeira
                    Mensagem = GestaoCaixaBO.EnviarMoedaEstrangeira(udtOperacaoRecebida.XmlOperacao, TipoMovimentoPJRealizado, null);
                    if (Mensagem != string.Empty)
                    {
                        using (_MqConnector = new MQConnector())
                        {
                            _MqConnector.MQConnect();
                            _MqConnector.MQQueueOpen("A7Q.E.ENTRADA", MQConnector.enumMQOpenOptions.PUT);
                            _MqConnector.Message = Mensagem;
                            _MqConnector.MQPutMessage();
                            _MqConnector.MQQueueClose();
                            _MqConnector.MQEnd();
                        }
                    }
                }
                #endregion

                #region >>> Verifica se Envia Retorno para o Legado >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_MESG_RETN"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    // Appenda tag de situação da operação no XML de Retorno para o legado
                    Comum.Comum.AppendNode(ref _XMLOperacao, "MESG", "CO_ULTI_SITU_PROC", ((int)statusRetornoLegado).ToString());

                    // Envia retorno legado
                    Mensagem = TratarRetorno(_XMLOperacao);
                    using (_MqConnector = new MQConnector())
                    {
                        _MqConnector.MQConnect();
                        _MqConnector.MQQueueOpen("A7Q.E.ENTRADA_NET", MQConnector.enumMQOpenOptions.PUT);
                        _MqConnector.Message = Mensagem;
                        _MqConnector.MQPutMessage();
                        _MqConnector.MQQueueClose();
                        _MqConnector.MQEnd();
                    }

                }
                #endregion

                #region >>> Gravar NumeroSequenciaOperacao na Mensagem R2 associada, e gerar OperacaoMensagemInterna de RetornoLegado ref à ela >>>
                if (_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB != null) // este dataset _DsTB_MESG_RECB_ENVI_SPB é populado no começo do processamento, nas funções ValidaRemessaArbitragem ou ValidaRemessaInterbancario
                {
                    if (_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB.Count == 1)
                    {
                        MensagemSpbDAO.EstruturaMensagemSPB EstruturaMensagemSPB = new MensagemSpbDAO.EstruturaMensagemSPB();
                        _MensagemSpbDATA = new MensagemSpbDAO();
                        _MensagemSpbDATA.SelecionarMensagensPorControleIF(_DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPB.DefaultView[0]["NU_CTRL_IF"].ToString());
                        if (_MensagemSpbDATA.Itens.Length == 1)
                        {
                            EstruturaMensagemSPB = _MensagemSpbDATA.ObterMensagemLida();
                            EstruturaMensagemSPB.NU_SEQU_OPER_ATIV = int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString());
                            EstruturaMensagemSPB.CO_LOCA_LIQU = int.Parse(entidadeOperacao.CO_LOCA_LIQU.ToString());
                            base.AlterarMensagemSPB(ref EstruturaMensagemSPB);
                        }

                        // Gerar OperacaoMensagemInterna de RetornoLegado ref à Mensagem R2 associada
                        #region <<< insere TB_TEXT_XML e TB_OPER_ATIV_MESG_INTE >>>
                        CodigoTextXML_MensagemSPB = int.Parse(EstruturaMensagemSPB.CO_TEXT_XML.ToString());
                        _CodigoTextXML = _TextXMLData.InserirBase64(base.SelecionarTextoBase64(CodigoTextXML_MensagemSPB));
                        _ParametroOperacaoMensagemInterna.NU_SEQU_OPER_ATIV = int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString());
                        _ParametroOperacaoMensagemInterna.DH_MESG_INTE = EstruturaMensagemSPB.DH_REGT_MESG_SPB.ToString();
                        _ParametroOperacaoMensagemInterna.TP_MESG_INTE = DataSetCache.TB_TIPO_OPER.DefaultView[0]["TP_MESG_RETN_INTE"].ToString().Trim();
                        _ParametroOperacaoMensagemInterna.TP_SOLI_MESG_INTE = (int)Comum.Comum.enumTipoSolicitacao.RetornoLegado;
                        _ParametroOperacaoMensagemInterna.CO_TEXT_XML = _CodigoTextXML;
                        _ParametroOperacaoMensagemInterna.TP_FORM_MESG_SAID = 0;
                        _OperacaoMensagemInternaDATA.Inserir(_ParametroOperacaoMensagemInterna);
                        #endregion

                    }
                }
                #endregion

            }
            catch
            {
                throw;
            }
        }
        #endregion

        #region <<< OperacaoDisponivelConsulta >>>
        public void OperacaoDisponivelConsulta(OperacaoDAO.EstruturaOperacao parametroOPER, Comum.Comum.EnumInidicador enumDisponibilizaConsulta)
        {
            OperacaoDAO OperacaDATA = new OperacaoDAO();

            // preencher os dados da operação q devem ser alterados
            parametroOPER.IN_DISP_CONS = enumDisponibilizaConsulta;
            parametroOPER.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
            parametroOPER.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;

            // atualiza a operação
            OperacaDATA.Atualizar(parametroOPER);
        }
        #endregion

        #region <<< LiberarAutomatico >>>
        /// <summary>
        /// Verifica se a funcionalidade de Liberação está automático ou não no workflow, controlando o status que a operação deve ficar
        /// </summary>
        /// <param name="parametroOPER">daod da operação R0</param>
        /// <param name="entidadeMensagem">mensagem recebida</param>
        /// <param name="statusOperacao">status da operação</param>
        /// <param name="enumEstorno">indica o estorno</param>
        public bool LiberarAutomatico(udtOperacao udtOperacaoRecebida, OperacaoDAO.EstruturaOperacao entidadeOperacao)
        {
            //OperacaoDAO OperacaDATA = new OperacaoDAO();
            HistoricoOperacao HistoricoBO = new HistoricoOperacao();
            RegraWorkflow RegraWorkflowBO = new RegraWorkflow(base.DataSetCache);
            RegraWorkflow.enumFuncaoSistema EnumFuncaoSistemaLiberacao = new RegraWorkflow.enumFuncaoSistema();
            Comum.Comum.EnumStatusOperacao? StatusOperacaoSeLiberacaoAutomatica = new A8NET.Comum.Comum.EnumStatusOperacao();

            try
            {
                //Verificar a Grade de Horário
                //

                RegraWorkflowBO.ObterFuncaoSistemaStatus(udtOperacaoRecebida, RegraWorkflow.enumMacroFuncao.Liberacao, ref EnumFuncaoSistemaLiberacao, ref StatusOperacaoSeLiberacaoAutomatica);

                //Verifica regra para liberação automática
                int _CodigoRetornoVerificacao = 0;
                if (RegraWorkflowBO.VerificarRegraAutomatica(udtOperacaoRecebida, EnumFuncaoSistemaLiberacao, ref _CodigoRetornoVerificacao) == true)
                {
                    // alterar o status da operação
                    entidadeOperacao.CO_ULTI_SITU_PROC = (int)StatusOperacaoSeLiberacaoAutomatica;
                    base.AlterarStatusOperacao(int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString()), (int)StatusOperacaoSeLiberacaoAutomatica, 0, base.ObterTipoAcaoEnvioMensagemSPB(udtOperacaoRecebida.TipoOperacao));
                    return true;
                }
                else
                {
                    #region <<< atualiza justificativa TB_HIST_SITU_ACAO_OPER_ATIV >>>
                    ParametroHistoricoSituacaoOperacao.NU_SEQU_OPER_ATIV = int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString());
                    ParametroHistoricoSituacaoOperacao.CO_SITU_PROC = int.Parse(entidadeOperacao.CO_ULTI_SITU_PROC.ToString());
                    ParametroHistoricoSituacaoOperacao.TP_JUST_SITU_PROC = _CodigoRetornoVerificacao;
                    _HistSituacaoOperacaoDATA.AtualizarJustificativa(ParametroHistoricoSituacaoOperacao);
                    #endregion

                    return false;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.LiberarAutomatico() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< MontarMensagem() >>>
        /// <summary>
        /// Enviar mensagem de retorno para o Legado
        ///   - Montar protocolo de integração A7
        ///   - Montar Remessa
        ///   - Incluir a mensagem de retorno na tabela de Mensagem Interna
        /// </summary>
        /// <param name="parametroOPER"></param>
        /// <param name="entidadeMensagem"></param>
        /// <returns></returns>
        private string MontarMensagem(udtOperacao udtOperacaoRecebida, OperacaoDAO.EstruturaOperacao entidadeOperacao, ref bool enviarMensagem)
        {
            string CodigoMensagem;
            enviarMensagem = false;
            int TipoBackOffice;
            string NumeroControleIF;
            string Mensagem;
            string NumeroComandoOperacao = null;
            string RegistroOperacaoCambial2 = null;
            MensagemSpbDAO MensagemSPBDATA = new MensagemSpbDAO();
            TextXmlDAO TextXMLData = new TextXmlDAO();
            MensagemSpbDAO.EstruturaMensagemSPB ParametroMensagemSPB = new MensagemSpbDAO.EstruturaMensagemSPB();
            DateTime? DataOperacaoCambioSisbacen = new DateTime();

            try
            {
                TipoBackOffice = udtOperacaoRecebida.TipoBackoffice;
                CodigoMensagem = _MensagemSPB.Substring(0, 3);

                //obter CO_MESG do udtOperacaoRecebida
                if (CodigoMensagem == "STR" || CodigoMensagem == "PAG")
                {
                    return string.Empty;
                }
                else
                {
                    NumeroControleIF = ObterNumeroControleIF();
                    Mensagem = MontarMensagemNZ(ref udtOperacaoRecebida, NumeroControleIF, TipoBackOffice);
                }

                //=========================================================
                //obtem conteúdo de tags específicas
                //=========================================================
                //DT_OPER_CAMB_SISBACEN    
                DataOperacaoCambioSisbacen = null;
                if (_MensagemSPB == "CAM0043")
                {
                    if (udtOperacaoRecebida.RowOperacao.Table.Columns.Contains("DT_EVEN_CAMB") == true)
                        if (udtOperacaoRecebida.RowOperacao["DT_EVEN_CAMB"].ToString() != string.Empty) 
                            DataOperacaoCambioSisbacen = Comum.Comum.ConvertDtToDateTime(udtOperacaoRecebida.RowOperacao["DT_EVEN_CAMB"].ToString());
                }

                //CO_REG_OPER_CAMB
                if (udtOperacaoRecebida.RowOperacao.Table.Columns.Contains("CO_REG_OPER_CAMB") == true)
                    if (udtOperacaoRecebida.RowOperacao["CO_REG_OPER_CAMB"].ToString() != string.Empty)
                        NumeroComandoOperacao = udtOperacaoRecebida.RowOperacao["CO_REG_OPER_CAMB"].ToString();

                //CO_REG_OPER_CAMB2
                if (udtOperacaoRecebida.RowOperacao.Table.Columns.Contains("CO_REG_OPER_CAMB2") == true)
                    if (udtOperacaoRecebida.RowOperacao["CO_REG_OPER_CAMB2"].ToString() != string.Empty)
                        RegistroOperacaoCambial2 = udtOperacaoRecebida.RowOperacao["CO_REG_OPER_CAMB2"].ToString();
                //=========================================================

                _CodigoTextXML = TextXMLData.InserirBase64(udtOperacaoRecebida.XmlOperacao.InnerXml);
                ParametroMensagemSPB.NU_CTRL_IF = NumeroControleIF;
                ParametroMensagemSPB.DH_REGT_MESG_SPB = DateTime.Now;  //tem q usar funcao para obter data.....
                ParametroMensagemSPB.DH_RECB_ENVI_MESG_SPB = DateTime.Now;
                ParametroMensagemSPB.NU_SEQU_CNTR_REPE = 1;
                ParametroMensagemSPB.CO_ULTI_SITU_PROC = (int)Comum.Comum.EnumStatusMensagem.EnviadaBUS;
                ParametroMensagemSPB.CO_EMPR = udtOperacaoRecebida.RowOperacao["CO_EMPR"].ToString();
                ParametroMensagemSPB.CO_TEXT_XML = _CodigoTextXML;
                ParametroMensagemSPB.IN_ENTR_MANU = (int)Comum.Comum.EnumInidicador.Nao;
                ParametroMensagemSPB.CO_MESG_SPB = _MensagemSPB;
                ParametroMensagemSPB.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroMensagemSPB.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                ParametroMensagemSPB.DH_ULTI_ATLZ = DateTime.Now;
                ParametroMensagemSPB.TP_BKOF = TipoBackOffice;
                ParametroMensagemSPB.NU_SEQU_OPER_ATIV = entidadeOperacao.NU_SEQU_OPER_ATIV.ToString();
                ParametroMensagemSPB.CO_LOCA_LIQU = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_LOCA_LIQU").InnerText;
                ParametroMensagemSPB.SG_SIST = udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString();
                ParametroMensagemSPB.CO_VEIC_LEGA = udtOperacaoRecebida.RowOperacao["CO_VEIC_LEGA"].ToString();
                ParametroMensagemSPB.DT_OPER_CAMB_SISBACEN = DataOperacaoCambioSisbacen;
                ParametroMensagemSPB.NU_COMD_OPER = NumeroComandoOperacao;
                ParametroMensagemSPB.NR_OPER_CAMB_2 = RegistroOperacaoCambial2;
                MensagemSPBDATA.Inserir(ParametroMensagemSPB);
                   
                //gerar historico da situacao Mensagem SPB
                this.AlterarStatusMensagemSPB(
                 ref ParametroMensagemSPB,
                     Comum.Comum.EnumStatusMensagem.EnviadaBUS);
            
                Mensagem = string.Concat(Mensagem, udtOperacaoRecebida.XmlOperacao.InnerXml);
                enviarMensagem = true;

                return Mensagem;
            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.MontarMensagem() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< MontarMensagemNZ() >>>
        /// <summary>
        /// Enviar mensagem de retorno para o Legado
        ///   - Montar protocolo de integração A7
        ///   - Montar Remessa
        ///   - Incluir a mensagem de retorno na tabela de Mensagem Interna
        /// </summary>
        /// <param name="parametroOPER"></param>
        /// <param name="entidadeMensagem"></param>
        /// <returns></returns>
        private string MontarMensagemNZ(ref udtOperacao udtOperacaoRecebida, string numeroControleIF, int tipoBackOffice)
        {
            string TipoMensagemOriginal = string.Empty;
            string Protocolo = string.Empty;
            string HeaderNZ = string.Empty;
            XmlDocument XmlOperacaoAux = udtOperacaoRecebida.XmlOperacao;

            try
            {
                // Monta protocolo de Mensagem para o sistema NZ
                Protocolo = string.Concat(
                        _MensagemSPB.PadRight(9, ' '), //TipoMensagem    
                        "A8 ", //SiglaSistemaOrigem
                        "NZ ", //SiglaSistemaDestino
                        udtOperacaoRecebida.RowOperacao["CO_EMPR"].ToString().PadLeft(5, '0') 
                        );

                // Obtem Header Mensagem NZ
                HeaderNZ = base.MontarHeaderMensagemNZ(_MensagemSPB,
                                                       int.Parse(udtOperacaoRecebida.RowOperacao["CO_EMPR"].ToString()),
                                                       numeroControleIF);

                // Incluir tag CO_MESG
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_MESG") == null) Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "CO_MESG", _MensagemSPB);
                else udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//CO_MESG").InnerText = _MensagemSPB;

                // Appenda tag DT_OPER_ATIV caso ainda não exista
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_OPER_ATIV") == null)
                {
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "DT_OPER_ATIV", udtOperacaoRecebida.RowOperacao["DT_MESG"].ToString());
                }

                // Incluir Header NZ na mensagem para enviar BUS
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TX_HEAD_NZ") == null) Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "TX_HEAD_NZ", HeaderNZ);
                else udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TX_HEAD_NZ").InnerText = HeaderNZ;

                // Appenda tag NU_CTRL_IF caso ainda não exista
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_CTRL_IF") == null) Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "NU_CTRL_IF", numeroControleIF);
                else udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_CTRL_IF").InnerText = numeroControleIF;

                return Protocolo;
            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.MontarMensagemNZ() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> ObterNumeroControleIF() >>>
        private string ObterNumeroControleIF()
        {
            //OracleConnection OracleConn = null;
            OracleCommand OraCommand = new OracleCommand();
            //A8NETOracleParameter OracleParameter = new A8NETOracleParameter();
            //OracleParameter ParametroSeqMsg = OracleParameter.SQ_MESG_FILA(null, ParameterDirection.Output);
            OracleParameter ParametroOUT = A8NETOracleParameter.SEQUENCIA(null, ParameterDirection.Output);
            string NumeroControleIF = string.Empty;
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(new BaseDAO().GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    OraCommand.Parameters.Clear();
                    OraCommand.Connection = OracleConn;
                    OraCommand.CommandType = CommandType.StoredProcedure;
                    OraCommand.CommandText = "A8PROC.A8P_SEQUENCIA_NZ";
            
                    OraCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.SISTEMA("A8", ParameterDirection.Input),
                        ParametroOUT}
                     );
                    OraCommand.ExecuteNonQuery();
                }

                if (ParametroOUT.Value == DBNull.Value) throw new Exception("Não foi possível obter Número de Controle IF");
                else NumeroControleIF = ParametroOUT.Value.ToString();

                return NumeroControleIF;

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.ObterNumeroControleIF() - " + ex.ToString());
            }
        
        }
        #endregion

        #region <<< SalvarMensagemRejeitada >>>
        protected void SalvarMensagemRejeitada(string mensagemErro)
        {
            A8NET.Data.DAO.RemessaRejeitadaDAO.EstruturaRemessaRejeitada ParametroRemessaRejeitada;
            A8NET.Data.DAO.RemessaRejeitadaDAO RemessaRejeitadaDATA;
            A8NET.Data.DAO.TextXmlDAO TextXMLData = new A8NET.Data.DAO.TextXmlDAO();
            XmlDocument XmlMensagemErro;
            int CodigoTextXML = 0;
            string XmlErro = "";

            try
            {
                ParametroRemessaRejeitada = new RemessaRejeitadaDAO.EstruturaRemessaRejeitada();
                RemessaRejeitadaDATA = new Data.DAO.RemessaRejeitadaDAO();
                XmlMensagemErro = new XmlDocument();

                XmlMensagemErro.LoadXml(mensagemErro);
                CodigoTextXML = TextXMLData.InserirBase64(mensagemErro);

                ParametroRemessaRejeitada.SG_SIST_ORIG_INFO = XmlMensagemErro.SelectSingleNode("//SG_SIST_ORIG").InnerText;
                ParametroRemessaRejeitada.TP_MESG_INTE = int.Parse(XmlMensagemErro.SelectSingleNode("//TP_MESG").InnerText);
                ParametroRemessaRejeitada.CO_EMPR = int.Parse(XmlMensagemErro.SelectSingleNode("//CO_EMPR").InnerText);
                ParametroRemessaRejeitada.CO_TEXT_XML_REJE = CodigoTextXML;
                ParametroRemessaRejeitada.CO_TEXT_XML_RETN_SIST_ORIG = CodigoTextXML;

                // Monta Texto do Xml de Erro
                XmlErro = "<Erro>";
                for (int i = 0; i < _DtErro.Rows.Count; i++)
                {
                    XmlErro = string.Concat(XmlErro, "<Grupo_ErrorInfo>",
                                                        "<Number>", _DtErro.Rows[i]["CD_ERRO"].ToString(), "</Number>",
                                                        "<Description>", _DtErro.Rows[i]["DS_ERRO"].ToString(), "</Description>",
                                                        "<ComputerName>", A8NET.Comum.Comum.NomeMaquina, "</ComputerName>",
                                                        "<Source>", _DtErro.Rows[i]["CM_ERRO"].ToString(), "</Source>",
                                                        "<ErrorType>1</ErrorType>",
                                                     "</Grupo_ErrorInfo>");
                }
                XmlErro = string.Concat(XmlErro, "</Erro>");

                ParametroRemessaRejeitada.TX_XML_ERRO = A8NET.Comum.Comum.Base64Encode(XmlErro);
                ParametroRemessaRejeitada.DH_REME_REJE = DateTime.Now;
                RemessaRejeitadaDATA.Inserir(ParametroRemessaRejeitada);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< MontaMensagemRejeitada >>>
        private string MontaMensagemRejeitada(string mensagemErro, string mensagemSPB)
        {
            XmlDocument XmlMensagemRejeitada;
            string DataMovimento = string.Empty;

            try
            {
                XmlMensagemRejeitada = new XmlDocument();
                XmlMensagemRejeitada.LoadXml(mensagemErro);

                if (Comum.Comum.LerNode(XmlMensagemRejeitada, "DT_MOVI") != string.Empty) DataMovimento = XmlMensagemRejeitada.DocumentElement.SelectSingleNode("//DT_MOVI").InnerText;
                else DataMovimento = DateTime.Today.ToString("yyyyMMdd");

                A8NET.Comum.Comum.AppendNode(ref XmlMensagemRejeitada, "MESG", "TP_RETN", "2");
                A8NET.Comum.Comum.AppendNode(ref XmlMensagemRejeitada, "MESG", "DtMovto", DataMovimento);

                // se o TipoSolicitacao=9 então MensagemSPB é fixo = "CAM0005"
                if (_TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Reativacao) mensagemSPB = "CAM0005";
                
                // Appenda tag CO_MESG_SPB
                A8NET.Comum.Comum.AppendNode(ref XmlMensagemRejeitada, "MESG", "CO_MESG_SPB", mensagemSPB != string.Empty ? mensagemSPB : "NAO_IDENT");

                for (int i = 0; i < _DtErro.Rows.Count; i++)
                {
                    A8NET.Comum.Comum.AppendNode(ref XmlMensagemRejeitada, "MESG", "CO_ERRO" + (i + 1), _DtErro.Rows[i]["CD_ERRO"].ToString());
                    A8NET.Comum.Comum.AppendNode(ref XmlMensagemRejeitada, "MESG", "DE_ERRO" + (i + 1), _DtErro.Rows[i]["DS_ERRO"].ToString());
                }

                return XmlMensagemRejeitada.OuterXml;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< EnviarMensagemRejeicaoLegado >>>
        protected void EnviarMensagemRejeicaoLegado(string mensagemRejeitada)
        {
            MQConnector MqConnector = null;
            XmlDocument XmlMensagemSaida;
            string MensagemSaida;
            string SiglaSistemaDestino;
            string TipoMensagem;
            string TipoMensagemRetorno;
            string CodigoEmpresaRetorno;
            string ProtocoloRetornoLegado;
            int OutN;

            try
            {
                XmlMensagemSaida = new XmlDocument();
                XmlMensagemSaida.LoadXml(mensagemRejeitada);

                SiglaSistemaDestino = XmlMensagemSaida.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerText.Trim().ToUpper();
                TipoMensagem = XmlMensagemSaida.DocumentElement.SelectSingleNode("TP_MESG").InnerText;
                TipoMensagemRetorno = base.DataSetCache.TB_TIPO_OPER.Select("TP_MESG_RECB_INTE='" + (int.TryParse(TipoMensagem, out OutN) ? Convert.ToString(OutN) : TipoMensagem.Trim()) + "'")[0]["TP_MESG_RETN_INTE"].ToString();
                TipoMensagemRetorno = TipoMensagemRetorno.PadLeft(9, '0');
                CodigoEmpresaRetorno = XmlMensagemSaida.DocumentElement.SelectSingleNode("CO_EMPR").InnerText.PadLeft(5, '0');

                // Protocolo Retorno Legado
                // TipoMensagem        //String * 9 = TipoMensagemRetorno
                // SiglaSistemaOrigem  //String * 3 = "A8 "
                // SiglaSistemaDestino //String * 3 = SiglaSistemaDestino
                // CodigoEmpresa       //String * 5 = CodigoEmpresaRetorno
                ProtocoloRetornoLegado = string.Concat(TipoMensagemRetorno, "A8 ", SiglaSistemaDestino.PadRight(3, ' '), CodigoEmpresaRetorno);

                MensagemSaida = string.Concat(ProtocoloRetornoLegado, mensagemRejeitada);

                //Put na fila A7Q.E.ENTRADA
                using (MqConnector = new A8NET.ConfiguracaoMQ.MQConnector())
                {
                    MqConnector.MQConnect();
                    MqConnector.MQQueueOpen("A7Q.E.ENTRADA", MQConnector.enumMQOpenOptions.PUT);
                    MqConnector.Message = MensagemSaida;
                    MqConnector.MQPutMessage();
                    MqConnector.MQQueueClose();
                    MqConnector.MQEnd();
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region >>> ObterTipoOperacao() >>>
        /// <summary>
        ///  Obtêm o TipoOperacao da Operacao, de acordo com regras cadastradas. Retorna 0(zero) caso não encontre.
        /// </summary>
        /// <param name="udtOperacaoRecebida">udt da Operação que está sendo processada</param>
        /// <returns></returns>
        private int ObterTipoOperacao(udtOperacao udtOperacaoRecebida)
        {

            short QtdeCondicaoVerdadeira;
            int TipoMensagem;
            int TipoOperacao;
            string Atributo;
            string ConteudoTagCondicao;
            string ConteudoTagXML;
            DataTable dtDistinctCols;
            
            try
            {
        
                //carrega variaveis 
                TipoMensagem = int.Parse(udtOperacaoRecebida.RowOperacao["TP_MESG"].ToString());

                DataSetCache.CaseSensitive = false;
                DataSetCache.TB_TIPO_OPER.DefaultView.RowFilter = string.Format("tp_mesg_recb_inte ='{0}'", TipoMensagem);
                DataSetCache.TB_TIPO_OPER.DefaultView.Sort = "TP_OPER";

                if (DataSetCache.TB_TIPO_OPER.DefaultView.Count == 1)
                {
                    return int.Parse(DataSetCache.TB_TIPO_OPER.DefaultView[0]["tp_oper"].ToString());
                }
                
                else if (DataSetCache.TB_TIPO_OPER.DefaultView.Count > 1)
                {

                    foreach (DataRowView RowTipoOper in DataSetCache.TB_TIPO_OPER.DefaultView)
                    {
                        TipoOperacao = int.Parse(RowTipoOper["TP_OPER"].ToString());

                        DataSetCache.TB_TIPO_OPER_CNTD_ATRB.DefaultView.RowFilter = string.Format("tp_oper = {0}", TipoOperacao);
                        string[] distinctCols = { "NO_ATRB_MESG" };
                        dtDistinctCols = DataSetCache.TB_TIPO_OPER_CNTD_ATRB.DefaultView.ToTable(true, distinctCols);

                        QtdeCondicaoVerdadeira = 0;
                        foreach (DataRow RowAtributo in dtDistinctCols.Rows)
                        {
                            Atributo = RowAtributo["NO_ATRB_MESG"].ToString();
                            DataSetCache.TB_TIPO_OPER_CNTD_ATRB.DefaultView.RowFilter = string.Format(@"no_atrb_mesg ='{0}'
                                                                                                    AND tp_oper      = {1}",                                                                                                        
                                                                                                    Atributo,
                                                                                                    TipoOperacao);

                            foreach (DataRowView RowCondicao in DataSetCache.TB_TIPO_OPER_CNTD_ATRB.DefaultView)
                            {
                                ConteudoTagCondicao = RowCondicao["DE_CNTD_ATRB"].ToString();

                                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//" + Atributo) != null)
                                {
                                    ConteudoTagXML = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//" + Atributo).InnerText;

                                    if (ConteudoTagXML.Trim() == ConteudoTagCondicao.Trim())
                                    {
                                        QtdeCondicaoVerdadeira++;
                                        break;
                                    }
                                }
                            }
                        }

                        //verifica se a condição de todas as tags foi atendida
                        if (QtdeCondicaoVerdadeira == dtDistinctCols.DefaultView.Count)
                        {
                            return TipoOperacao;
                        }
                    }
                }

                return 0;

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.ObterTipoOperacao() - " + ex.ToString());
            }

        }
        #endregion

        #region >>> ObterTipoBackoffice() >>>
        /// <summary>
        /// Obtêm o TipoBackoffice da Operacao, de acordo com o Veículo Legal e Sigla Sistema Origem recebidos.
        /// Retorna 0(zero) caso não encontre.
        /// </summary>
        /// <param name="udtOperacaoRecebida">udt da Operação que está sendo processada</param>
        /// <returns></returns>
        private int ObterTipoBackoffice(udtOperacao udtOperacaoRecebida)
        {

            string CodigoVeiculoLegal;
            string SiglaSistemaOrigem;
            DateTime DataOperacao;
            string SelectVeiculoLegal;

            try
            {

                //carrega variaveis 
                SiglaSistemaOrigem = udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString();
                CodigoVeiculoLegal = udtOperacaoRecebida.RowOperacao["CO_VEIC_LEGA"].ToString();
                DataOperacao = A8NET.Comum.Comum.ConvertDtToDateTime(udtOperacaoRecebida.RowOperacao["DT_OPER_ATIV"].ToString());

                DataSetCache.CaseSensitive = false;
                SelectVeiculoLegal = @"CO_VEIC_LEGA='" + CodigoVeiculoLegal.Trim() +
                                                     "' AND SG_SIST='" + SiglaSistemaOrigem.Trim() +
                                                     "' AND DT_INIC_VIGE<='" + DataOperacao +
                                                     "' AND (DT_FIM_VIGE IS NULL OR DT_FIM_VIGE>='" + DataOperacao + "')";
                try
                {
                    return int.Parse(DataSetCache.TB_VEIC_LEGA.Select(SelectVeiculoLegal)[0]["TP_BKOF"].ToString());
                }
                catch //nao encontrou o VeiculoLegal para conseguir obter o TipoBackoffice
                {
                    return 0;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.ObterTipoBackoffice() - " + ex.ToString());
            }

        }
        #endregion

        #region >>> ObterDataLiquidacaoPJ() >>>
        /// <summary>
        /// Obtêm a Data de Liquidação a ser enviada para o PJ, de acordo com o Tipo de Remessa.
        /// Retorna string.Empty caso não encontre.
        /// </summary>
        /// <param name="TipoRemessaPJ">Tipo de Remessa PJ</param>
        /// <returns></returns>
        private string ObterDataLiquidacaoPJ(udtOperacao udtOperacaoRecebida, Comum.Comum.EnumTipoMoedaPJ TipoMoedaPJ)
        {

            string DataLiquidacaoPJ = null;
            
            try
            {

                switch (TipoMoedaPJ)
                {
                    case Comum.Comum.EnumTipoMoedaPJ.MoedaNacional:

                        if (int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_MESG").InnerText) == (int)Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria)
                        {
                            if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_LIQU_OPER_MOED_NACI") != null)
                            {
                                DataLiquidacaoPJ = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_LIQU_OPER_MOED_NACI").InnerText;
                            }
                        }
                        else if (int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_MESG").InnerText) == (int)Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao)
                        {
                            if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DtLiquid") != null)
                            {
                                DataLiquidacaoPJ = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DtLiquid").InnerText;
                            }
                        }
                        break;

                    case Comum.Comum.EnumTipoMoedaPJ.MoedaEstrangeira:

                        if (int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_MESG").InnerText) == (int)Comum.Comum.EnumTipoMensagem.RegistroOperacaoInterbancaria
                         || int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_MESG").InnerText) == (int)Comum.Comum.EnumTipoMensagem.RegistroOperacaoArbitragem
                         || int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_MESG").InnerText) == (int)Comum.Comum.EnumTipoMensagem.IFInformaLiquidacaoInterbancaria)
                        {
                            if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_LIQU_OPER") != null)
                            {
                                DataLiquidacaoPJ = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DT_LIQU_OPER").InnerText;
                            }
                        }
                        else if (int.Parse(udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_MESG").InnerText) == (int)Comum.Comum.EnumTipoMensagem.ComplementoInformacoesContratacaoInterbancarioViaLeilao)
                        {
                            if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DtLiquid") != null)
                            {
                                DataLiquidacaoPJ = udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//DtLiquid").InnerText;
                            }
                        }
                        break;
               
                    default:
                        break;
                }

                return DataLiquidacaoPJ;
            
            }

            catch (Exception ex)
            {
                return null;
                //throw new Exception("Operacao.ObterDataLiquidacaoPJ() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ProcessarReativacaoFluxo >>>
        private void ProcessarReativacaoFluxo(udtOperacao udtOperacaoRecebida)
        {
            bool ErroConciliacao = false;

            try
            {

                _EventoProcessamento = "RecebimentoReativacao";

                // Procura Operação pelo CO_OPER_ATIV
                _OperacaoDATA.ObterOperacao(udtOperacaoRecebida.RowOperacao["CO_OPER_ATIV"].ToString(), udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString());
                
                // Rejeita caso não encontre a Operação
                if (_OperacaoDATA == null)
                {
                    _DtErro.Rows.Add("4514", "Não foi possível efetuar a reativação. Operação não encontrada.", "A8NET.Mensagem.Operacao.Operacao.ProcessarReativacaoFluxo()");
                    return;
                }

                // Verifica se a Operação encontrada tem algum dos Status esperados
                if (ValidaStatusOperacaoEncontrada(udtOperacaoRecebida, _OperacaoDATA.TB_OPER_ATIV) == false)
                {
                    _DtErro.Rows.Add("4515", "Não foi possível efetuar a reativação. Operação encontrada não está no status CANCELADA CÂMARA.", "A8NET.Mensagem.Operacao.Operacao.ProcessarReativacaoFluxo()");
                    return;
                }

                // Verifica se NumeroOperacaoCambial e NumeroOperacaoCambial2 enviados no layout de Reativacao estão iguais aos da Operação encontrada (pelo CO_OPER_ATIV)
                if (_OperacaoDATA.TB_OPER_ATIV.NU_COMD_OPER.ToString() != string.Empty)
                {
                    if (udtOperacaoRecebida.RowOperacao["CO_REG_OPER_CAMB"].ToString().Trim() != _OperacaoDATA.TB_OPER_ATIV.NU_COMD_OPER.ToString().Trim())
                    {
                        _DtErro.Rows.Add("4521", "Não foi possível efetuar a reativação. Registro Operação Cambial informado não está igual ao da operação encontrada.", "A8NET.Mensagem.Operacao.Operacao.ProcessarReativacaoFluxo()");
                        ErroConciliacao = true;
                    }
                }
                if (_OperacaoDATA.TB_OPER_ATIV.NR_OPER_CAMB_2.ToString() != string.Empty)
                {
                    if (udtOperacaoRecebida.RowOperacao["CO_REG_OPER_CAMB2"].ToString() != _OperacaoDATA.TB_OPER_ATIV.NR_OPER_CAMB_2.ToString())
                    {
                        _DtErro.Rows.Add("4522", "Não foi possível efetuar a reativação. Registro Operação Cambial 2 informado não está igual ao da operação encontrada.", "A8NET.Mensagem.Operacao.Operacao.ProcessarReativacaoFluxo()");
                        ErroConciliacao = true;
                    }
                }
                if (ErroConciliacao == true) return;

                // Atualiza a mensagem a ser enviada para CAM0005, devido ser um fluxo de Reativação
                _MensagemSPB = "CAM0005";

                // Se TipoNegociacaoInterbancaria = 1 (SEM CÂMARA – PCAM380), então NÃO envia Mensagem SPB
                if (Comum.Comum.LerNode(_XMLOperacao, "TP_NEGO_INTB") == ((int)enumTipoNegociacaoInterbancaria.SemCamara).ToString())
                {
                    DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_MESG_SPB"] = (int)Comum.Comum.EnumInidicador.Nao;
                }

                // Gerencia Chamadas
                GerenciarChamadas(udtOperacaoRecebida, _EventoProcessamento, _OperacaoDATA.TB_OPER_ATIV, null);

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ProcessarReativacaoFluxo() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ProcessarCancelamento >>>
        private void ProcessarCancelamento(udtOperacao udtOperacaoRecebida)
        {
            XmlDocument XmlOperacaoAux = udtOperacaoRecebida.XmlOperacao;

            try
            {

                _EventoProcessamento = "RecebimentoCancelamento";

                // Procura Operação pelo CO_OPER_ATIV
                _OperacaoDATA.ObterOperacao(udtOperacaoRecebida.RowOperacao["CO_OPER_ATIV"].ToString(), udtOperacaoRecebida.RowOperacao["SG_SIST_ORIG"].ToString());

                // Rejeita caso não encontre a Operação
                if (_OperacaoDATA == null)
                {
                    _DtErro.Rows.Add("4517", "Não foi possível efetuar o cancelamento. Operação não encontrada.", "A8NET.Mensagem.Operacao.Operacao.ProcessarCancelamento()");
                    return;
                }

                // Verifica se a Operação encontrada tem algum dos Status esperados
                if (ValidaStatusOperacaoEncontrada(udtOperacaoRecebida, _OperacaoDATA.TB_OPER_ATIV) == false)
                {
                    _DtErro.Rows.Add("4518", "Não foi possível efetuar o cancelamento. Operação encontrada não está em nenhum dos status esperados.", "A8NET.Mensagem.Operacao.Operacao.ProcessarCancelamento()");
                    return;
                }

                _NumeroSequenciaOperacao = long.Parse(_OperacaoDATA.TB_OPER_ATIV.NU_SEQU_OPER_ATIV.ToString());

                // Appenda informações que serão necessárias nas remessas para o PJ >>>
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//NU_SEQU_OPER_ATIV") == null)
                {
                    Comum.Comum.AppendNode(ref XmlOperacaoAux, "MESG", "NU_SEQU_OPER_ATIV", _NumeroSequenciaOperacao.ToString());
                }

                // Gerencia Chamadas
                GerenciarChamadas(udtOperacaoRecebida, _EventoProcessamento, _OperacaoDATA.TB_OPER_ATIV, Comum.Comum.EnumStatusOperacao.CanceladaOrigem);

            }
            catch (Exception ex)
            {

                throw new Exception("Operacao.ProcessarCancelamento() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< TratarRetorno >>>
        /// <summary>
        /// Enviar mensagem de retorno para o Legado
        ///   - Montar protocolo de integração A7
        ///   - Montar Remessa
        ///   - Incluir a mensagem de retorno na tabela de Mensagem Interna
        /// </summary>
        /// <param name="parametroOPER"></param>
        /// <param name="entidadeMensagem"></param>
        /// <returns></returns>
        private string TratarRetorno(XmlDocument xmlOperacao)
        {
            OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna ParametroMsgInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
            OperacaoMensagemInternaDAO OperacaoInternaDAO = new OperacaoMensagemInternaDAO();
            TextXmlDAO TextXmlData = new TextXmlDAO();
            DsParametrizacoes.TB_TIPO_OPERRow RowTipoOPER = null;
            string TipoMensagemOriginal = xmlOperacao.DocumentElement.SelectSingleNode("//TP_OPER").InnerXml;
            string CodigoOPER = _NumeroSequenciaOperacao.ToString();
            string Protocolo = string.Empty;
            int FormatoSaidaMsg = 0;
            string TipoMensagem = "0";
            int CodigoTextXML = 0;

            try
            {
 
                if (xmlOperacao.DocumentElement.SelectSingleNode("TP_MESG_RETN_INTE") == null)
                {
                    RowTipoOPER = DataSetCache.TB_TIPO_OPER.FindByTP_OPER(decimal.Parse(xmlOperacao.DocumentElement.SelectSingleNode("TP_OPER").InnerXml));

                    if (RowTipoOPER == null || RowTipoOPER.TP_MESG_RETN_INTE == "0") return "";

                    TipoMensagem = RowTipoOPER.TP_MESG_RETN_INTE.ToString();
                }
                else TipoMensagem = xmlOperacao.DocumentElement.SelectSingleNode("TP_MESG_RETN_INTE").InnerXml;

                //corrige as tags SG_SIST_ORIG e SG_SIST_DEST
                xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml = xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml;
                xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml = "A8";

                Protocolo = string.Concat(TipoMensagem.PadLeft(9, '0'),
                        "A8 ",
                        xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml.Trim().PadRight(3, ' '),
                        xmlOperacao.DocumentElement.SelectSingleNode("CO_EMPR").InnerXml.PadLeft(5, '0'));

                xmlOperacao.DocumentElement.SelectSingleNode("TP_MESG").InnerXml = TipoMensagem.ToString().PadLeft(9, '0');
                Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "DT_MESG", DateTime.Today.ToString("yyyyMMdd"));
                Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "HO_MESG", DateTime.Now.ToString("HHmm"));

                Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "TP_RETN", "1");

                Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "DtMovto", xmlOperacao.SelectSingleNode("//DT_MOVI").InnerText);

                DataRow[] RowREGR = DataSetCache.TB_REGR_SIST_DEST.Select("TP_MESG='" + TipoMensagem + "' AND SG_SIST_DEST='" + xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml + "'", "DH_INIC_VIGE_REGR_TRAP DESC");
                if (RowREGR.Length > 0) int.TryParse(RowREGR[0]["TP_FORM_MESG_SAID"].ToString(), out FormatoSaidaMsg);

                // Armazena a mensagem original na tabela TB_TEXT_XML 
                CodigoTextXML = TextXmlData.InserirBase64(xmlOperacao.OuterXml);

                // preencher os dados da operação mensagem interna
                ParametroMsgInterna.NU_SEQU_OPER_ATIV = CodigoOPER;
                ParametroMsgInterna.TP_MESG_INTE = TipoMensagem; //TipoMensagemOriginal;
                ParametroMsgInterna.TP_FORM_MESG_SAID = FormatoSaidaMsg;
                ParametroMsgInterna.TP_SOLI_MESG_INTE = (int)Comum.Comum.enumTipoSolicitacao.RetornoLegado;
                ParametroMsgInterna.CO_TEXT_XML = CodigoTextXML;
                ParametroMsgInterna.DH_MESG_INTE = OperacaoInternaDAO.ObterDataGravacao(CodigoOPER).AddSeconds(1);
                // insere os dados no banco
                OperacaoInternaDAO.Inserir(ParametroMsgInterna);

                return Protocolo + xmlOperacao.OuterXml; //parametroOPER.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPB1.TratarRetorno() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> ValidaStatusOperacaoEncontrada >>>
        private bool ValidaStatusOperacaoEncontrada(udtOperacao udtOperacaoRecebida, OperacaoDAO.EstruturaOperacao entidadeOperacao)
        {
            Comum.Comum.EnumStatusOperacao[] ListaStatusOperacao = null;

            try
            {
                switch (_TipoSolicitacao)
                {
                    case (int)Comum.Comum.enumTipoSolicitacao.Cancelamento:
                        #region >>> Status para Cancelamento >>>
                        if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemTelaCega
                         || udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.ConfirmacaoOperacaoInterbancariaSemTelaCega)
                        {
                            Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = { Comum.Comum.EnumStatusOperacao.EmSer, Comum.Comum.EnumStatusOperacao.Concordancia, Comum.Comum.EnumStatusOperacao.ConcordanciaAutomatica };
                            ListaStatusOperacao = _ListaStatusOperacao;
                        }
                        if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemCamara
                         || udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.ConfirmacaoOperacaoInterbancariaSemCamara)
                        {
                            Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = { Comum.Comum.EnumStatusOperacao.EmSer, Comum.Comum.EnumStatusOperacao.Concordancia, Comum.Comum.EnumStatusOperacao.ConcordanciaAutomatica };
                            ListaStatusOperacao = _ListaStatusOperacao;
                        }
                        if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancarioEletronico)
                        {
                            Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = { Comum.Comum.EnumStatusOperacao.EmSer, Comum.Comum.EnumStatusOperacao.AConciliarRegistro };
                            ListaStatusOperacao = _ListaStatusOperacao;
                        }
                        if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.InformaOperacaoArbitragemParceiroPais
                         || udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.InformaConfirmacaoOperArbitragemParceiroPais)
                        {
                            Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = { Comum.Comum.EnumStatusOperacao.EmSer, Comum.Comum.EnumStatusOperacao.Concordancia, Comum.Comum.EnumStatusOperacao.ConcordanciaAutomatica };
                            ListaStatusOperacao = _ListaStatusOperacao;
                        }
                        break;
                        #endregion

                    case (int)Comum.Comum.enumTipoSolicitacao.Complementacao:
                        #region >>> Status para Complementacao >>>
                        if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.CAMInformaContratacaoInterbancarioViaLeilao)
                        {
                            Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = { Comum.Comum.EnumStatusOperacao.Registrada };
                            ListaStatusOperacao = _ListaStatusOperacao;
                        }
                        break;
                        #endregion

                    case (int)Comum.Comum.enumTipoSolicitacao.Reativacao:
                        #region >>> Status para Reativacao >>>
                        if (true)
                        {
                            Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = { Comum.Comum.EnumStatusOperacao.CanceladaCamara };
                            ListaStatusOperacao = _ListaStatusOperacao;
                        }
                        break;
                        #endregion

                    default:
                        return false;
                }

                // Valida status
                foreach (Comum.Comum.EnumStatusOperacao statusOperacao in ListaStatusOperacao)
                {
                    if ((int)statusOperacao == int.Parse(entidadeOperacao.CO_ULTI_SITU_PROC.ToString()))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.ValidaStatusOperacaoEncontrada() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> VerificarSeRegraEspecificaAlteracaoStatus >>>
        private bool VerificarSeRegraEspecificaAlteracaoStatus(udtOperacao udtOperacaoRecebida, ref Comum.Comum.EnumStatusOperacao statusOperacao)
        {
            try
            {
                // se TipoNegociacaoInterbancaria = InterbancarioEletronico então retorna status "A CONCILIAR REGISTRO"
                if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_INTB") != null)
                {
                    if (udtOperacaoRecebida.XmlOperacao.SelectSingleNode("//TP_NEGO_INTB").InnerText.Trim() == ((int)enumTipoNegociacaoInterbancaria.InterbancarioEletronico).ToString())
                    {
                        statusOperacao = Comum.Comum.EnumStatusOperacao.AConciliarRegistro;
                        return true;
                    }
                }

                // se TipoOperacao = CAMInformaContratacaoInterbancarioViaLeilao, então retorna status CONFIRMADA
                if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.CAMInformaContratacaoInterbancarioViaLeilao)
                {
                    statusOperacao = Comum.Comum.EnumStatusOperacao.Confirmada;
                    return true;
                }

                // se TipoSolicitacao = Reativacao
                if (_TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Reativacao)
                {
                    statusOperacao = Comum.Comum.EnumStatusOperacao.ReativacaoSolicitada;
                    return true;
                }

                // se TipoSolicitacao = Cancelamento
                if (_TipoSolicitacao == (int)Comum.Comum.enumTipoSolicitacao.Cancelamento)
                {
                    statusOperacao = Comum.Comum.EnumStatusOperacao.CanceladaOrigem;
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception("Operacao.VerificarSeRegraEspecificaAlteracaoStatus() - " + ex.ToString());
            }
        }
        #endregion

    }
}


