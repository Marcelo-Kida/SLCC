using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Data;
using A8NET.Data;
using A8NET.Data.DAO;
using A8NET.Historico;
using A8NET.ConfiguracaoMQ;
using A8NET.Mensagem.SPB;
using A8NET.Mensagem.SPB.udt;
using System.Xml;

namespace A8NET.Mensagem.SPB
{
    public class MensagemSPB : Mensagem, IDisposable
    {
        #region <<< Variaveis >>>
        protected Data.DsParametrizacoes _DsCache;
        protected int _Out;
        private MensagemSpbDAO _MensagemSpbDATA;
        private HistoricoSituacaoMensagemDAO _HistoricoMensagemDATA;
        protected DateTime? _DataOperacaoCambioSisbacen = new DateTime();
        protected string _CodigoClienteSisbacen = string.Empty;
        protected string _NumeroComandoOperacao = string.Empty;
        protected string _RegistroOperacaoCambial2 = string.Empty;
        protected MensagemSpbDAO.EstruturaMensagemSPB _EstruturaMensagemSPB = new MensagemSpbDAO.EstruturaMensagemSPB();
        protected A8NET.Mensagem.Conciliacao _ConciliacaoBO = new Conciliacao();
        protected XmlDocument _XmlOperacaoAux;
        protected XmlDocument _XmlOperacaoR0 = new XmlDocument();
        protected string _StatusOperacao = string.Empty;
        protected string _TipoMensagemRetorno = string.Empty;
        protected string _IndicadorAceite = string.Empty;
        protected string _EventoProcessamento = string.Empty;
        protected long _NumeroSequenciaOperacao = 0;
        protected int _TipoOperacao = 0;
        protected MensagemSPB _MensagemSPB;
        #endregion

        #region <<< Structs >>>
        public struct EstruturaStatusOperacaoMensagem
        {
            public int StatusOperacao;
            public int StatusMensagemSPB;
            public string SituacaoRecebida;
            public string SituacaoMensagemSPB;
        }
        #endregion

        #region <<< Construtores >>>
        public MensagemSPB(Data.DsParametrizacoes dataSetCache)
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion

            _MensagemSpbDATA = new MensagemSpbDAO();
            _HistoricoMensagemDATA = new HistoricoSituacaoMensagemDAO();
            _OperacaoDATA = new OperacaoDAO();

            base.DataSetCache = dataSetCache;
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~MensagemSPB()
        {
            this.Dispose();
        }
        #endregion

        #region <<< ProcessaMensagem >>>
        /// <summary>
        /// Mètodo trata o processamendo da mensagem oriunda do SPB, trata as mensagens CAMxxR1 e CAMxxR2
        /// </summary>
        /// <param name="nomeFila">nome da </param>
        /// <param name="mensagemRecebida"></param>
        public override void ProcessaMensagem(string nomeFila, string mensagemRecebida)
        {
            udtMensagem EntidadeMensagem = new udtMensagem();
            try
            {
                EntidadeMensagem.Parse(mensagemRecebida);
                _MensagemSPB = MensagemSPBFactory(EntidadeMensagem.TipoMensagem, base.DataSetCache);
                _MensagemSPB.GerenciaMensagem(EntidadeMensagem);
            }
            catch
            {
                throw;
            }
        }
        #endregion

        #region <<< GerenciaMensagem >>>
        /// <summary>
        /// Método trata a mensagem conforme o tipo da mensagem
        /// </summary>
        /// <param name="udtMsg">udt da mensagem (header, linha com os dados, xml original, fila)</param>
        public virtual void GerenciaMensagem(udt.udtMensagem entidadeMensagem)
        {
            Comum.Comum.EnumStatusMensagem StatusMensagem;
            int TipoBackoffice = 0;
            int? CodigoLocalLiquidacao;

            // Obtem dados genéricos para salvar na MensagemSPB
            StatusMensagem = ObterStatusMensagem(entidadeMensagem.TipoMensagem, entidadeMensagem.CodigoMensagem);
            TipoBackoffice = ObterTipoBackOffice(entidadeMensagem.CodigoMensagem);
            CodigoLocalLiquidacao = ObterCodigoLocalLiquidacao(entidadeMensagem.CodigoMensagem);

            // Obtem dados específicos para salvar na MensagemSPB
            ObterConteudoTagsEspecificas(entidadeMensagem.XmlMensagem);

            // Salva a Mensagem SPB
            Incluir(
                entidadeMensagem.XmlMensagem.InnerXml,
                entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ,
                TipoBackoffice,
                entidadeMensagem.CodigoMensagem,
                null,
                entidadeMensagem.CabecalhoMensagem.CodigoEmpresa,
                _NumeroComandoOperacao,
                null,
                StatusMensagem,
                A8NET.Comum.Comum.EnumInidicador.Nao,
                CodigoLocalLiquidacao,
                "", // CódigoVeiculoLegal
                entidadeMensagem.CabecalhoMensagem.SiglaSistemaEnviouNZ.Trim(),
                1,
                null,
                null,
                _RegistroOperacaoCambial2
                );

            // Gerencia chamada
            GerenciarChamada(entidadeMensagem);

        }
        #endregion

        #region <<< METODO FACTORY - MensagemSPBFactory >>>
        /// <summary>
        /// Cria a instancia do objeto de valorização conforme o tipo da mensagem
        /// </summary>
        /// <param name="codigoMensagem">codigo da mensagem do SPB</param>
        /// <returns>retorna as instancias conforme o tipo da mensagem(R1, R2 e Aviso Informação)</returns>
        public static MensagemSPB MensagemSPBFactory(string tipoMensagem, Data.DsParametrizacoes dataSetCache)
        {
            if (tipoMensagem == "R1")
            {
                return new MensagemSPBR1(dataSetCache);
            }
            else if (tipoMensagem == "R2")
            {
                return new MensagemSPBR2(dataSetCache);
            }
            else
            {
                return new MensagemAvisoInformacao(dataSetCache);
            }
        }
        #endregion

        #region <<< VerificarMensagemTratadaSLCC >>>
        /// <summary>
        /// Verificar se a mensagem é tratada pelo SLCC, consulta na tabela TB_TIPO_OPER se o tipo de mensagem é tratado pelo sistema
        /// </summary>
        /// <param name="codigoMensagem">codigo da mensagem </param>
        /// <returns>retorna true ou false </returns>
        public bool VerificarMensagemTratadaSLCC(string codigoMensagem)
        {
            string CodigoMensagem  = codigoMensagem.Substring(0, 7);
            if (DataSetCache.TB_TIPO_OPER.Select("CO_MESG_SPB_REGT_OPER='" + CodigoMensagem + "'").Length > 0)
            {
                if (!CodigoMensagem.Trim().Equals("BMC0012")) return true;
            }
            else
            {
                if (codigoMensagem.Substring(0, 3) == "CAM") return true;
                //if ("|SEL|CTP|BMC|CAM|".IndexOf("|" + CodigoMensagem.Trim() + "|") != -1) return true;
            }
            return false;
        }
        #endregion

        #region <<< ObterNomeTagSituacao >>>
        /// <summary>
        /// Retorna as TAGS do grupo
        /// </summary>
        /// <param name="codigoGrupo">codigo do grupo</param>
        /// <returns>aos nomes das tags</returns>
        protected DataRow[] SelecionarNomeTagSituacao(string codigoGrupo)
        {
            if (DataSetCache.TB_SITU_SPB_SITU_PROC.Select("SG_GRUP_MESG_SPB='" + codigoGrupo + "'").Length > 0)
            {
                object Retorno = DataSetCache.TB_SITU_SPB_SITU_PROC
                    .DefaultView
                    .ToTable(true, new string[] { "NO_TAG", "SG_GRUP_MESG_SPB" })
                    .Select("SG_GRUP_MESG_SPB='" + codigoGrupo + "'");

                return (DataRow[])Retorno;
            }
            else
            {
                if(codigoGrupo == "CAM") return null;
                else throw new Exception("3017"); //3017 - Grupo de mensagem sem Situação SPB X Situação Processamento cadastrado.
            }
        }
        #endregion

        #region <<< ObterStatus >>>
        /// <summary>
        /// Obter status da mensagem e da operação
        /// </summary>
        /// <param name="entidadeMensagem">mensagem</param>
        /// <returns>status e situação atuala da mensagem e da operação</returns>
        protected EstruturaStatusOperacaoMensagem ObterStatus(udt.udtMensagem entidadeMensagem)
        {
            EstruturaStatusOperacaoMensagem EstruturaStatus = new EstruturaStatusOperacaoMensagem();
            string NomeTipoTAG =  string.Empty;
            string NomeTAG =  string.Empty;
            bool TemTipoTAG = false;

            try
            {
                DataRow[] rowRetorno = this.SelecionarNomeTagSituacao(entidadeMensagem.CodigoGrupoMensagem);
                if (rowRetorno == null)
                {
                    EstruturaStatus.StatusMensagemSPB = (int)Comum.Comum.EnumStatusMensagem.Registrada;
                    EstruturaStatus.StatusOperacao = (int)Comum.Comum.EnumStatusOperacao.Registrada;
                    return EstruturaStatus;
                }
                // retorna o nome das TAGS que o grupo de acesso
                foreach (DataRow rowTagSituacao in rowRetorno)
                {
                    NomeTAG  = rowTagSituacao["NO_TAG"].ToString();
                    if (!entidadeMensagem.LinhaMensagem.Table.Columns.Contains(NomeTAG)) continue;
                    EstruturaStatus.SituacaoMensagemSPB = entidadeMensagem.LinhaMensagem[NomeTAG].ToString();

                    foreach (DsParametrizacoes.TB_SITU_SPB_SITU_PROCRow rowSITU in DataSetCache.TB_SITU_SPB_SITU_PROC
                        .Select(string.Format("SG_GRUP_MESG_SPB = '{0}' AND NO_TAG = '{1}'", 
                        entidadeMensagem.CodigoGrupoMensagem, NomeTAG)))
                    {
                        if (entidadeMensagem.LinhaMensagem[NomeTAG].ToString() == rowSITU.DE_DOMI)
                        {
                            EstruturaStatus.StatusOperacao = rowSITU.CO_SITU_PROC_OPER_ATIV;
                            EstruturaStatus.StatusMensagemSPB = rowSITU.CO_SITU_PROC_MESG_SPB;
                            break;
                        }
                        if(!rowSITU.IsSQ_TIPO_TAGNull() && TemTipoTAG == false)
                        {
                            // só precisa entrar neste IF uma unica vez, pois, na sequencia o mesmo repete o valor
                            TemTipoTAG = true;
                            NomeTipoTAG= rowSITU.SQ_TIPO_TAG;
                        }
                    }
                    if (TemTipoTAG)
                    {
                        EstruturaStatus.SituacaoRecebida = string.Concat(NomeTipoTAG.PadLeft(5, '0'), "|", EstruturaStatus.SituacaoMensagemSPB);
                    }
                }
                return EstruturaStatus;
            }
            catch (Exception ex)
            {
                if (int.TryParse(ex.Message,out  _Out)) throw ex;
                else throw new Exception("ObterStatus()" + ex.ToString());
            }
        }
        #endregion

        #region <<< Incluir >>>
        public DateTime Incluir(string xmlMensagem, string controleIF, int? tipoBackOffice, string codigoMensagem, int? seqOperacao, string codigoEmpresa, 
            string comandoDocumento, string situacaoMensagemSPB, Comum.Comum.EnumStatusMensagem statusMensagem, 
            Comum.Comum.EnumInidicador indicadorEntrarManual, int? codigoLocalLiquidacao, string codigoVeicLegal, string siglaSistema, int controladorRepeticao,
            DateTime? dataOperacaoCambioSisbacen, string codigoClienteSisbacen, string registroOperacaoCambial2)
        {
            HistoricoSituacaoMensagemDAO.EstruturaHistoricoSituacaoMsg ParametroHistoricoSituacaoMsg = new HistoricoSituacaoMensagemDAO.EstruturaHistoricoSituacaoMsg();
            MensagemSpbDAO.EstruturaMensagemSPB ParametroMensagemSPB = new MensagemSpbDAO.EstruturaMensagemSPB();
            TextXmlDAO TextXmlData = new TextXmlDAO();
            DateTime DataGravacao =  new DateTime();
            int CodigoTxtXML = 0;
            XmlDocument XmlR1 = new XmlDocument();
            long RegistroOperacaoCambial2 = 0;

            try 
            {   
                DataGravacao = _MensagemSpbDATA.ObterDataGravacao(controleIF);

                XmlR1.LoadXml(xmlMensagem);

                // Insere tabela de texto xml
                CodigoTxtXML = TextXmlData.InserirBase64(XmlR1.SelectSingleNode("//SISMSG").OuterXml);

                // preenche dados da tabela de mensagem
                ParametroMensagemSPB.NU_CTRL_IF = controleIF;
                ParametroMensagemSPB.DH_REGT_MESG_SPB = DataGravacao;
                ParametroMensagemSPB.NU_SEQU_CNTR_REPE = controladorRepeticao;
                ParametroMensagemSPB.NU_SEQU_OPER_ATIV = seqOperacao;
                ParametroMensagemSPB.TP_BKOF = (tipoBackOffice == (int)Comum.Comum.EnumTipoBackOffice.Todos || tipoBackOffice == 0 ? null : tipoBackOffice);
                ParametroMensagemSPB.CO_EMPR = codigoEmpresa;
                ParametroMensagemSPB.DH_RECB_ENVI_MESG_SPB = DateTime.Now;
                ParametroMensagemSPB.CO_MESG_SPB = codigoMensagem;
                ParametroMensagemSPB.NU_COMD_OPER = comandoDocumento;
                ParametroMensagemSPB.CO_SITU_MESG_SPB = situacaoMensagemSPB;
                ParametroMensagemSPB.CO_TEXT_XML = CodigoTxtXML;
                ParametroMensagemSPB.CO_ULTI_SITU_PROC = (int)statusMensagem;
                ParametroMensagemSPB.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroMensagemSPB.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                ParametroMensagemSPB.IN_ENTR_MANU = (int)indicadorEntrarManual;
                ParametroMensagemSPB.CO_LOCA_LIQU = codigoLocalLiquidacao; // (codigoLocalLiquidacao == 0 ? null : codigoLocalLiquidacao.ToString());
                ParametroMensagemSPB.CO_VEIC_LEGA = codigoVeicLegal;
                ParametroMensagemSPB.SG_SIST = siglaSistema.Trim();
                ParametroMensagemSPB.DT_OPER_CAMB_SISBACEN = dataOperacaoCambioSisbacen;
                ParametroMensagemSPB.CD_CLIE_SISBACEN = codigoClienteSisbacen;
                if (long.TryParse(registroOperacaoCambial2, out RegistroOperacaoCambial2) == true)
                    ParametroMensagemSPB.NR_OPER_CAMB_2 = RegistroOperacaoCambial2;

                // inserir a mensagem SPB
                _MensagemSpbDATA.Inserir(ParametroMensagemSPB);

                // preecher os dados do histórico da sitação da mensagem
                ParametroHistoricoSituacaoMsg.NU_CTRL_IF = controleIF;
                ParametroHistoricoSituacaoMsg.DH_REGT_MESG_SPB = DataGravacao;
                ParametroHistoricoSituacaoMsg.NU_SEQU_CNTR_REPE = controladorRepeticao;
                ParametroHistoricoSituacaoMsg.DH_SITU_ACAO_MESG_SPB = _HistoricoMensagemDATA.ObterDataGravacao(controleIF);
                ParametroHistoricoSituacaoMsg.CO_SITU_PROC = (int)statusMensagem;
                ParametroHistoricoSituacaoMsg.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroHistoricoSituacaoMsg.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;

                // inserir histórico
                _HistoricoMensagemDATA.Inserir(ParametroHistoricoSituacaoMsg);

                return DataGravacao;
                
            }
            catch (Exception ex)
            {
                throw new Exception("Método: MensagemSPB.Incluir() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterConteudoTagsEspecificas >>>
        public void ObterConteudoTagsEspecificas(XmlDocument xmlMensagemSPB)
        {
            byte Contador = 0;

            try
            {
                switch (xmlMensagemSPB.SelectSingleNode("//CodMsg").InnerXml)
                {
                    #region >>> dataOperacaoCambioSisbacen >>
                    case "CAM0043":
                        if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//DtEvtCAM") != null)
                            if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//DtEvtCAM").InnerXml != string.Empty)
                               _DataOperacaoCambioSisbacen = Comum.Comum.ConvertDtToDateTime(xmlMensagemSPB.DocumentElement.SelectSingleNode("//DtEvtCAM").InnerXml);
                        break;
                    #endregion

                    #region >>> codigoClienteSisbacen >>>
                    case "CAM0045R1":
                    case "CAM0046R1":
                        if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//IdentdPessoaCli") != null)
                            _CodigoClienteSisbacen = xmlMensagemSPB.DocumentElement.SelectSingleNode("//IdentdPessoaCli").InnerXml;
                        break;
                    #endregion

                    #region >>> numeroComandoOperacao(=RegistroOperacaoCambial) e RegistroOperacaoCambial2 >>>
                    case "CAM0021R1": case "CAM0021R2":
                    case "CAM0023R1": case "CAM0023R2":
                    case "CAM0024R2":
                    case "CAM0025R2":
                    case "CAM0027R2":
                    case "CAM0028R2":
                    case "CAM0029R2":
                    case "CAM0030R2":
                    case "CAM0031R2":
                    case "CAM0033R2":
                    case "CAM0034R2":
                    case "CAM0006R1": case "CAM0006R2":
                    case "CAM0009R1": case "CAM0009R2":
                    case "CAM0013R1": case "CAM0013R2":
                    case "CAM0007R2":
                    case "CAM0008R2":
                    case "CAM0010R2":
                    case "CAM0014R2":
                    case "CAM0012R1":
                    case "CAM0015"  :
                    case "CAM0055"  :
                    case "CAM0005R2":
                    case "CAM0054R1":
                        if (xmlMensagemSPB.DocumentElement.SelectNodes("//RegOpCaml") != null)

                            if (xmlMensagemSPB.DocumentElement.SelectNodes("//RegOpCaml").Count == 1)
                            {
                                // se o xml tiver apenas uma tag RegOpCaml então grava no campo NU_COMD_OPER
                                _NumeroComandoOperacao = xmlMensagemSPB.DocumentElement.SelectSingleNode("//RegOpCaml").InnerXml;
                            }
                            else if (xmlMensagemSPB.DocumentElement.SelectNodes("//RegOpCaml").Count == 2)
                            {
                                Contador = 0;
                                // se o xml tiver duas tags RegOpCaml então grava a primeira no campo NU_COMD_OPER e a segunda no campo NR_OPER_CAMB_2
                                foreach (XmlNode node in xmlMensagemSPB.DocumentElement.SelectNodes("//RegOpCaml"))
                                {
                                    Contador++;
                                    if (Contador == 1) _NumeroComandoOperacao = node.InnerText;
                                    if (Contador == 2) _RegistroOperacaoCambial2 = node.InnerText;
                                }
                            }

                        if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//RegOpCaml2") != null)
                        {
                            _RegistroOperacaoCambial2 = xmlMensagemSPB.DocumentElement.SelectSingleNode("//RegOpCaml2").InnerText;
                        }

                        break;
                    #endregion

                    #region >>> dataOperacaoCambioSisbacen e numeroComandoOperacao >>>
                    case "CAM0044R1":
                        if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//DtEvtCAM") != null)
                            if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//DtEvtCAM").InnerXml != string.Empty)
                                _DataOperacaoCambioSisbacen = Comum.Comum.ConvertDtToDateTime(xmlMensagemSPB.DocumentElement.SelectSingleNode("//DtEvtCAM").InnerXml);
                        if (xmlMensagemSPB.DocumentElement.SelectSingleNode("//RegOpCaml") != null)
                            _NumeroComandoOperacao = xmlMensagemSPB.DocumentElement.SelectSingleNode("//RegOpCaml").InnerXml;
                        break;
                    #endregion
                }
            }
            catch { }
        }
        #endregion

        #region >>> ObterTipoOperacao >>>
        protected Comum.Comum.EnumTipoOperacao ObterTipoOperacao(string codigoMensagem)
        {
            
            switch (codigoMensagem)
            {
                case "CAM0007R2": return Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemTelaCega;
                
                case "CAM0008R2": return Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemTelaCega;
                
                case "CAM0010R2": return Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancariaSemCamara;
                
                case "CAM0014R2": return Comum.Comum.EnumTipoOperacao.InformaOperacaoArbitragemParceiroPais;
                
                case "CAM0015"  : return Comum.Comum.EnumTipoOperacao.CAMInformaContratacaoInterbancarioViaLeilao;
                
                case "CAM0055":
                case "CAM0005R2":
                case "BMC0005":
                    {
                        if (_ConciliacaoBO.VerificaConciliacao(ref _EstruturaMensagemSPB, _IndicadorAceite, ref _NumeroSequenciaOperacao, true) == true)
                        {
                            // Atualiza TipoBackoffice/CodigoVeiculoLegal/CodigoLocalLiquidacao da MensagemSPB, de acordo com a Operacao Conciliada
                            if (_NumeroSequenciaOperacao != 0)
                            {
                                //_MensagemSPBConciliadaComOperacao = true;
                                _OperacaoDATA.ObterOperacao((int)_NumeroSequenciaOperacao);
                                return (A8NET.Comum.Comum.EnumTipoOperacao)int.Parse(_OperacaoDATA.TB_OPER_ATIV.TP_OPER.ToString());
                            }
                        }
                        return 0; // caso não tenha conseguido conciliar com uma Operacao, então retorno 0
                        
                    }
                default: // obtem TipoOperacao de acordo com cadastro da tabela TB_TIPO_OPER
                    DataSetCache.TB_TIPO_OPER.DefaultView.RowFilter = "TP_MESG_RECB_INTE='" + codigoMensagem + "'";
                    if (DataSetCache.TB_TIPO_OPER.DefaultView.Count == 0) return 0;
                    return (A8NET.Comum.Comum.EnumTipoOperacao)int.Parse(DataSetCache.TB_TIPO_OPER.DefaultView[0]["TP_OPER"].ToString());
            };
                
        }
        #endregion

        #region >>> ObterStatusMensagem >>>
        protected Comum.Comum.EnumStatusMensagem ObterStatusMensagem(string tipoMensagem, string codigoMensagem)
        {
            Comum.Comum.EnumStatusMensagem StatusMensagem;

            if (tipoMensagem == "R1")
                StatusMensagem = Comum.Comum.EnumStatusMensagem.R1;
            else if (tipoMensagem == "R2")
                StatusMensagem = Comum.Comum.EnumStatusMensagem.R2;
            else
            {
                DataRow[] Retorno = DataSetCache.TB_MENSAGEM.Select("CO_MESG='" + codigoMensagem.Trim() + "'");
                if (Retorno.Length == 0)
                {
                    StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Informação;
                }
                else
                {
                    if (int.Parse(Retorno[0]["SQ_TIPO_FLUX"].ToString()) == (int)Comum.Comum.EnumTipoFluxo.TipoFluxo5)
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Informação;
                    else
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Aviso;
                }
            }

            return StatusMensagem;
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

        #region >>> ObterCodigoLocalLiquidacao >>>
        /// <summary>
        /// Este método sempre tem que retornar algo, porque se a MensagemSPB for gravada sem CO_LOCA_LIQU ela não é visualizada na tela
        /// </summary>
        /// <param name="codigoMensagem">Código da Mensagem</param>
           protected static int? ObterCodigoLocalLiquidacao(string codigoMensagem)
        {
            switch (codigoMensagem.Trim())
            {
                case "CAM0013R2":
                case "CAM0014R2":
                    return (int)Comum.Comum.EnumLocalLiquidacao.BMC;
                
                default:
                    if (codigoMensagem.Substring(0, 3).Equals("CAM"))
                    {
                        return (int)Comum.Comum.EnumLocalLiquidacao.CAM;
                    }
                    else
                    {
                        return null;
                    }
            }
        }
        #endregion

        #region >>> METODO VIRTUAL - ObterEventoProcessamento >>>
        public virtual string ObterEventoProcessamento(string codigoMensagem) { return string.Empty; }
        #endregion

        #region <<< GerenciarChamada >>>
        /// <summary>
        /// O metodo basicamente altera o status da operação do R0 e envia a mensagem para o legado
        /// </summary>
        /// <param name="parametroOPER">daod da operação R0</param>
        /// <param name="entidadeMensagem">mensagem recebida</param>
        /// <param name="statusOperacao">status da operação</param>
        /// <param name="enumEstorno">indica o estorno</param>
        public virtual void GerenciarChamada(udt.udtMensagem entidadeMensagem)
        {
            _XmlOperacaoAux = entidadeMensagem.XmlMensagem;
            bool MensagemSPBConciliadaComOperacao = false;
            bool MensagemSPBIncluiuOperacao = false;
            bool EntradaManual = false;

            try
            {
                // Retorna a MensagemSPB que acabou de ser gravada no banco no método anterior GerenciaMensagem()
                _MensagemSpbDATA.SelecionarMensagensPorControleIF(entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ);
                if (_MensagemSpbDATA.Itens.Length == 0) return; // MensagemSPB não foi gravada, portanto encerrar processamento
                _EstruturaMensagemSPB = _MensagemSpbDATA.ObterMensagemLida();
                
                // Obtem dados necessários durante o processamento
                _EventoProcessamento = this.ObterEventoProcessamento(entidadeMensagem.CodigoMensagem);
                _TipoOperacao = (int)ObterTipoOperacao(entidadeMensagem.CodigoMensagem); if (_TipoOperacao == 0) return;
                _TipoMensagemRetorno = DataSetCache.TB_TIPO_OPER.FindByTP_OPER(_TipoOperacao).TP_MESG_RETN_INTE.ToString();

                // Obtem parametrização de processamento
                DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.RowFilter = string.Format(@"NO_PROC_OPER_ATIV ='{0}' 
                                                                                        AND TP_OPER={1}",
                                                                                        _EventoProcessamento,
                                                                                        _TipoOperacao.ToString());
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.Count == 0) return; // Mensagem sem parametrização de processamento, encerrar SEM DAR EXPLICAÇÕES

                // Verifica se MensagemSPB gera Operação
                if (entidadeMensagem.CodigoMensagem == "CAM0015")
                {
                    IncluirOperacao(entidadeMensagem, _TipoOperacao);
                    if (_NumeroSequenciaOperacao != 0) MensagemSPBIncluiuOperacao = true;
                    else return; // se _NumeroSequenciaOperacao = 0 então significa que não conseguiu gerar Operacao, portanto aborta processamento
                }

                #region >>> Verifica se faz Conciliacao >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_VERI_REGR_CNCL"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    // Lê tag IndicadorAceite. Conteúdo tem que ser "S" ou "N"
                    if (entidadeMensagem.XmlMensagem.SelectSingleNode("//IndrActe") != null)
                    {
                        _IndicadorAceite = entidadeMensagem.XmlMensagem.SelectSingleNode("//IndrActe").InnerText;
                        if (_IndicadorAceite != "S" && _IndicadorAceite != "N") return;
                    }

                    // Verifica conciliação MensagemSPB com Operação. Caso a conciliação esteja OK, 
                    // o status da Operação e da MensagemSPB já serão atualizados dentro de ConciliacaoBO.VerificaConciliacao()
                    if (_ConciliacaoBO.VerificaConciliacao(ref _EstruturaMensagemSPB, _IndicadorAceite, ref _NumeroSequenciaOperacao, false) == true) //, ref _StatusOperacao
                    {
                        if (_NumeroSequenciaOperacao != 0)
                        {
                            MensagemSPBConciliadaComOperacao = true;
                        }
                    }
                    else
                    {
                        // Conciliação da MensagemSPB não OK, portanto encerrar processamento
                        return;
                    }

                }
                #endregion

                // Atualiza NumeroSequenciaOperacao/TipoBackoffice/CodigoLocalLiquidacao da MensagemSPB, para ficarem iguais aos da Operacao Conciliada
                if (_NumeroSequenciaOperacao != 0)
                {
                    _OperacaoDATA.ObterOperacao((int)_NumeroSequenciaOperacao);

                    if (entidadeMensagem.CodigoMensagem.ToString().Trim() != "BMC0005")
                    {
                        _EstruturaMensagemSPB.CO_LOCA_LIQU = _OperacaoDATA.TB_OPER_ATIV.CO_LOCA_LIQU;
                        if (MensagemSPBIncluiuOperacao == true)
                        {
                            _EstruturaMensagemSPB.NU_SEQU_OPER_ATIV = _OperacaoDATA.TB_OPER_ATIV.NU_SEQU_OPER_ATIV;
                        }
                        if (MensagemSPBConciliadaComOperacao == true)
                        {
                            _EstruturaMensagemSPB.TP_BKOF = _OperacaoDATA.ObterTipoBackoffice(_OperacaoDATA, DataSetCache.TB_VEIC_LEGA);
                        }
                        base.AlterarMensagemSPB(ref _EstruturaMensagemSPB);
                    }

                    // verifica se a Operaçao foi gerada via Entrada Manual
                    if (_OperacaoDATA.TB_OPER_ATIV.IN_ENTR_MANU.ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                    {
                        EntradaManual = true;
                    }
                }

                #region >>> Verifica se Envia Retorno para o Legado >>>
                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_MESG_RETN"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {

                    // Se a Operação foi gerada via Entrada Manual então não envia retorno para o legado
                    if (EntradaManual == true)
                    {
                        return;
                    }
                    else
                    {
                        if (entidadeMensagem.CodigoMensagem.Equals("BMC0005"))
                        {
                            base.AlterarStatusOperacao(int.Parse(_OperacaoDATA.TB_OPER_ATIV.NU_SEQU_OPER_ATIV.ToString()), (int)Comum.Comum.EnumStatusOperacao.RegistradaAutomatica, 0, 0);
                            _OperacaoDATA.TB_OPER_ATIV.CO_ULTI_SITU_PROC = (int)Comum.Comum.EnumStatusOperacao.RegistradaAutomatica;
                        }
                    }

                    // Se houver uma Operacao associada/conciliada à Mensagem, então appenda algumas tags ref à Operacao, necessárias no retorno para o legado
                    if (_NumeroSequenciaOperacao != 0)
                    {
                        // Retorna a Operação para obter CO_OPER_ATIV
                        _OperacaoDATA.ObterOperacao((int)_NumeroSequenciaOperacao);

                        // Appenda tags no XML de Retorno para o legado
                        if (_XmlOperacaoAux.SelectSingleNode("//CO_OPER_ATIV") == null) Comum.Comum.AppendNode(ref _XmlOperacaoAux, "MESG", "CO_OPER_ATIV", _OperacaoDATA.TB_OPER_ATIV.CO_OPER_ATIV.ToString());
                        else _XmlOperacaoAux.SelectSingleNode("//CO_OPER_ATIV").InnerText = _OperacaoDATA.TB_OPER_ATIV.CO_OPER_ATIV.ToString();
                        if (_XmlOperacaoAux.SelectSingleNode("//CO_ULTI_SITU_PROC") == null) Comum.Comum.AppendNode(ref _XmlOperacaoAux, "MESG", "CO_ULTI_SITU_PROC", _OperacaoDATA.TB_OPER_ATIV.CO_ULTI_SITU_PROC.ToString());
                        else _XmlOperacaoAux.SelectSingleNode("//CO_ULTI_SITU_PROC").InnerText = _OperacaoDATA.TB_OPER_ATIV.CO_ULTI_SITU_PROC.ToString();
                    }

                    // Envia retorno legado
                    this.TratarRetorno(entidadeMensagem.XmlMensagem, _TipoMensagemRetorno, entidadeMensagem.CabecalhoMensagem.CodigoEmpresa, _NumeroSequenciaOperacao);

                    // Para algumas mensagens específicas mudar status da Mensagem para "EnviadaLegado"
                    if (entidadeMensagem.CodigoMensagem.Equals("CAM0006R2")
                     || entidadeMensagem.CodigoMensagem.Equals("CAM0009R2")
                     || entidadeMensagem.CodigoMensagem.Equals("CAM0013R2")
                     || entidadeMensagem.CodigoMensagem.Equals("CAM0015"))
                    {
                        base.AlterarStatusMensagemSPB(ref _EstruturaMensagemSPB, A8NET.Comum.Comum.EnumStatusMensagem.EnviadaLegado);
                    }

                }
                #endregion

            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPB.GerenciarChamada - " + ex.ToString());
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
        private void TratarRetorno(XmlDocument xmlMensagem, string tipoMsgRetornoInterno, string codigoEmpresa, long numeroSequenciaOperacao)
        {
            OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna ParametroMsgInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
            OperacaoMensagemInternaDAO OperacaoInternaDAO = new OperacaoMensagemInternaDAO();
            TextXmlDAO TextXmlData = new TextXmlDAO();
            string Protocolo = string.Empty;
            int FormatoSaidaMsg = 0;
            int CodigoTextXML = 0;

            try
            {

                #region >>> Monta mensagem de retorno e coloca na fila de entrada do A7NET >>>
                foreach (DataRow Row in base.DataSetCache.TB_REGR_SIST_DEST.Select("TP_MESG='" + tipoMsgRetornoInterno +
                                                                                   "' AND SG_SIST_ORIG='A8'" +
                                                                                   " AND CO_EMPR_ORIG=" + codigoEmpresa +
                                                                                   " AND SG_SIST_DEST<>'R2'"))
                {
                    Protocolo = string.Concat(tipoMsgRetornoInterno.PadLeft(9, '0'), "A8 ", Row["SG_SIST_DEST"].ToString().ToUpper().PadRight(3, ' '), codigoEmpresa.PadLeft(5, '0'));

                    xmlMensagem.DocumentElement.SelectSingleNode("TP_MESG").InnerXml = tipoMsgRetornoInterno.ToString().PadLeft(9, '0');
                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "DT_MESG", DateTime.Today.ToString("yyyyMMdd"));
                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "HO_MESG", DateTime.Now.ToString("HHmm"));
                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "CO_MESG_SPB", xmlMensagem.SelectSingleNode("//CodMsg").InnerText);
                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "TP_RETN", "1");

                    //corrige as tags SG_SIST_ORIG e SG_SIST_DEST
                    xmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml = Row["SG_SIST_ORIG"].ToString().ToUpper().Trim();
                    xmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml = Row["SG_SIST_DEST"].ToString().ToUpper().Trim();

                    // Os sistemas GPC e R2 precisam da tag CO_VEIC_LEGA, pois ela foi incluída erroneamente nos layouts que enviamos para eles
                    if (Row["SG_SIST_DEST"].ToString().ToUpper().Trim() == "GPC"
                     || Row["SG_SIST_DEST"].ToString().ToUpper().Trim() == "R2"
                     || Row["SG_SIST_DEST"].ToString().ToUpper().Trim() == "BOL"
                     || Row["SG_SIST_DEST"].ToString().ToUpper().Trim() == "HQ")
                    {
                        Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "CO_VEIC_LEGA", "46");
                    }

                    MQConnector MqConnector = null;
                    using (MqConnector = new MQConnector())
                    {
                        MqConnector.MQConnect();
                        MqConnector.MQQueueOpen("A7Q.E.ENTRADA_NET", MQConnector.enumMQOpenOptions.PUT);
                        MqConnector.Message = Protocolo + xmlMensagem.InnerXml;
                        MqConnector.MQPutMessage();
                        MqConnector.MQQueueClose();
                        MqConnector.MQEnd();
                    }

                    // Obtem FormatoMensagemSaida
                    DataRow[] RowREGR = DataSetCache.TB_REGR_SIST_DEST.Select("TP_MESG='" + tipoMsgRetornoInterno + "' AND SG_SIST_DEST='" + Row["SG_SIST_DEST"].ToString().ToUpper().Trim() + "'", "DH_INIC_VIGE_REGR_TRAP DESC");
                    if (RowREGR.Length > 0) int.TryParse(RowREGR[0]["TP_FORM_MESG_SAID"].ToString(), out FormatoSaidaMsg);

                }
                #endregion

                #region >>> Gera registro na tabela OperacaoMensagemInterna caso a MensagemSPB seja conciliada/associada com Operacao >>>
                if (numeroSequenciaOperacao != 0)
                {
                    // Armazena a mensagem original na tabela TB_TEXT_XML 
                    CodigoTextXML = TextXmlData.InserirBase64(xmlMensagem.OuterXml);

                    // preencher os dados da operação mensagem interna
                    ParametroMsgInterna.NU_SEQU_OPER_ATIV = numeroSequenciaOperacao;
                    ParametroMsgInterna.TP_MESG_INTE = tipoMsgRetornoInterno; //TipoMensagemOriginal;
                    ParametroMsgInterna.TP_FORM_MESG_SAID = FormatoSaidaMsg;
                    ParametroMsgInterna.TP_SOLI_MESG_INTE = (int)Comum.Comum.enumTipoSolicitacao.RetornoLegado;
                    ParametroMsgInterna.CO_TEXT_XML = CodigoTextXML;
                    ParametroMsgInterna.DH_MESG_INTE = OperacaoInternaDAO.ObterDataGravacao(numeroSequenciaOperacao.ToString()).AddSeconds(1);
                    // insere os dados no banco
                    OperacaoInternaDAO.Inserir(ParametroMsgInterna);
                }
                #endregion

            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPB.TratarRetorno() - " + ex.ToString());
            }
        }
        #endregion

        #region >>> IncluirOperacao >>>
        protected void IncluirOperacao(udt.udtMensagem entidadeMensagem, int tipoOperacao)
        {
            OperacaoDAO.EstruturaOperacao ParametroOperacao = new OperacaoDAO.EstruturaOperacao();
            DsTB_OPER_ATIV_MESG_INTE DsTB_OPER_ATIV_MESG_INTE = new DsTB_OPER_ATIV_MESG_INTE();
            HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao ParametroHistoricoSituacaoOperacao = new HistoricoSituacaoOperacaoDAO.EstruturaHistoricoSituacaoOperacao();
            TextXmlDAO TextXMLData = new TextXmlDAO();
            
            try
            {
                // Campos fixos, valores default (para garantir que não ocorrerá erro na inserção da Operação, no caso de não conseguir obter algum campo da MensagemSPB)
                ParametroOperacao.SG_SIST = "E2";
                ParametroOperacao.CO_VEIC_LEGA = "46";
                ParametroOperacao.CO_LOCA_LIQU = (int)Comum.Comum.EnumLocalLiquidacao.CAM;
                ParametroOperacao.DT_OPER_ATIV = Comum.Comum.ConvertDtToDateTime(DateTime.Today.ToString("yyyyMMdd"));
                ParametroOperacao.IN_ENTR_SAID_RECU_FINC = 0; //0=vazio, ou seja, não é nem 1=Entrada e nem 2=Saída
                ParametroOperacao.IN_OPER_DEBT_CRED = 0;
                ParametroOperacao.VA_OPER_ATIV = 0;
                ParametroOperacao.CO_ULTI_SITU_PROC = (decimal)Comum.Comum.EnumStatusOperacao.Registrada;
                ParametroOperacao.DH_ULTI_ATLZ = DateTime.Now;
                ParametroOperacao.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                ParametroOperacao.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroOperacao.NU_SEQU_OPER_ATIV = 0;
                ParametroOperacao.TP_OPER = tipoOperacao;
                ParametroOperacao.CO_EMPR = int.Parse(entidadeMensagem.CabecalhoMensagem.CodigoEmpresa);
                ParametroOperacao.CO_USUA_CADR_OPER = Comum.Comum.UsuarioSistema;
                ParametroOperacao.IN_DISP_CONS = (int)Comum.Comum.EnumInidicador.Sim;
                ParametroOperacao.IN_ENTR_MANU = (int)Comum.Comum.EnumInidicador.Nao;

                // Obtem conteúdo de campos específicos
                if (entidadeMensagem.CodigoMensagem == "CAM0015") ObterCamposCAM0015(entidadeMensagem, ref ParametroOperacao);

                // Insere Operação
                _NumeroSequenciaOperacao = _OperacaoDATA.Inserir(ParametroOperacao);
                if (_NumeroSequenciaOperacao == 0) return; // se NumeroSequenciaOperacao = 0 então significa que não conseguiu gerar operacao, portanto encerra processamento

                // OBS: o registro na tabela A8.TB_OPER_ATIV_MESG_INTE será gerado mais a frente quando a MensagemSPB for ser retornada para o legado
                
                #region <<< insere TB_HIST_SITU_ACAO_OPER_ATIV >>>
                // insere o Historico de Situacao da Operacao
                ParametroHistoricoSituacaoOperacao.NU_SEQU_OPER_ATIV = _NumeroSequenciaOperacao;
                ParametroHistoricoSituacaoOperacao.DH_SITU_ACAO_OPER_ATIV = DateTime.Now;
                ParametroHistoricoSituacaoOperacao.CO_SITU_PROC = ParametroOperacao.CO_ULTI_SITU_PROC.ToString();
                ParametroHistoricoSituacaoOperacao.CO_USUA_ATLZ = Comum.Comum.UsuarioSistema;
                ParametroHistoricoSituacaoOperacao.CO_ETCA_USUA_ATLZ = Comum.Comum.NomeMaquina;
                _HistSituacaoOperacaoDATA.Inserir(ParametroHistoricoSituacaoOperacao);
                #endregion

            }
            catch (Exception ex)
            {
                return;
            }
        }
        #endregion

        #region >>> ObterCamposCAM0015 >>>
        private void ObterCamposCAM0015(udt.udtMensagem entidadeMensagem, ref OperacaoDAO.EstruturaOperacao parametroOperacao)
        {

            try
            {
                // Obtem tag TpOpCAM e já faz De-Para, do domínio do Book para o domínio da base do A8
                string TipoOperacaoCAM = Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "TpOpCAM");
                if (TipoOperacaoCAM == "V") TipoOperacaoCAM = "1";
                else if (TipoOperacaoCAM == "C") TipoOperacaoCAM = "2";
                else TipoOperacaoCAM = string.Empty;

                parametroOperacao.CO_LOCA_LIQU = (int)Comum.Comum.EnumLocalLiquidacao.CAM;
                parametroOperacao.CO_ULTI_SITU_PROC = (decimal)Comum.Comum.EnumStatusOperacao.Registrada;
                parametroOperacao.NU_COMD_OPER = Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "RegOpCaml");
                parametroOperacao.VA_OPER_ATIV = decimal.Parse(Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "VlrMN"));
                parametroOperacao.VA_MOED_ESTR = decimal.Parse(Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "VlrME"));
                parametroOperacao.PE_TAXA_NEGO = decimal.Parse(Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "TaxCam"));
                parametroOperacao.DT_OPER_ATIV = Comum.Comum.ConvertDtToDateTime(Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "DtMovto"));
                parametroOperacao.DT_LIQU_OPER_ATIV = Comum.Comum.ConvertDtToDateTime(Comum.Comum.LerNode(entidadeMensagem.XmlMensagem, "DtLiquid"));
                if (TipoOperacaoCAM != string.Empty)
                {
                    parametroOperacao.IN_OPER_DEBT_CRED = int.Parse(TipoOperacaoCAM);
                    parametroOperacao.IN_ENTR_SAID_RECU_FINC = int.Parse(TipoOperacaoCAM);
                }
                    
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPB.ObterCamposCAM0015() - " + ex.ToString());
            }

        }
        #endregion


        #region >>> MontaMensagemNZEntradaA8 >>>
        protected string MontaMensagemNZEntradaA8(DataView dvTB_MESG_RECB_ENVI_SPB)
        {
            XmlDocument XMLMensagemNZEntradaA8 = new XmlDocument();
            string HeaderNZ = string.Empty;
            string XMLMensagemSPB = string.Empty;

            try
            {
                HeaderNZ = base.MontarHeaderMensagemNZ(dvTB_MESG_RECB_ENVI_SPB[0]["CO_MESG_SPB"].ToString().Trim()
                                                      , int.Parse(dvTB_MESG_RECB_ENVI_SPB[0]["CO_EMPR"].ToString())
                                                      , dvTB_MESG_RECB_ENVI_SPB[0]["NU_CTRL_IF"].ToString());

                XMLMensagemSPB = base.SelecionarTextoBase64(int.Parse(dvTB_MESG_RECB_ENVI_SPB[0]["CO_TEXT_XML"].ToString()));

                XMLMensagemNZEntradaA8.LoadXml("<MESG></MESG>");
                Comum.Comum.AppendNode(ref XMLMensagemNZEntradaA8, "MESG", "TP_MESG", "000001000");
                Comum.Comum.AppendNode(ref XMLMensagemNZEntradaA8, "MESG", "SG_SIST_ORIG", "NZ");
                Comum.Comum.AppendNode(ref XMLMensagemNZEntradaA8, "MESG", "SG_SIST_DEST", "A8");
                Comum.Comum.AppendNode(ref XMLMensagemNZEntradaA8, "MESG", "CO_EMPR", dvTB_MESG_RECB_ENVI_SPB[0]["CO_EMPR"].ToString().PadLeft(5, '0'));
                Comum.Comum.AppendNode(ref XMLMensagemNZEntradaA8, "MESG", "TX_MESG", HeaderNZ + XMLMensagemSPB);

                return XMLMensagemNZEntradaA8.OuterXml;
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPB.MontaMensagemNZEntradaA8() - " + ex.ToString());
            }
        }
        #endregion

    }
}
