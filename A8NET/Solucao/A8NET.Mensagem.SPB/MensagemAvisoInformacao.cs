using System;
using System.Collections.Generic;
using System.Text;
using A8NET.Data;
using A8NET.Data.DAO;
using System.Xml;
using System.Data;
using A8NET.ConfiguracaoMQ;

namespace A8NET.Mensagem.SPB
{
    public class MensagemAvisoInformacao : MensagemSPB
    {

        #region <<< Construtor >>>
        public MensagemAvisoInformacao(Data.DsParametrizacoes dsCache) : base(dsCache)
        {
            //_OperacaoDATA = new OperacaoDAO();
            //_MensagemSpbDATA = new MensagemSpbDAO();
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~MensagemAvisoInformacao()
        {
            this.Dispose();
        }
        #endregion

        //#region <<< GerenciaMensagem >>>
        ///// <summary>
        ///// Método trata a mensagem conforme o tipo da mensagem
        ///// </summary>
        ///// <param name="udtMsg">udt da mensagem (header, linha com os dados, xml original, fila)</param>
        //public override void GerenciaMensagem(udt.udtMensagem entidadeMensagem)
        //{
        //    Comum.Comum.EnumStatusMensagem StatusMensagem;
        //    int TipoBackoffice = 0;
        //    int? CodigoLocalLiquidacao;

        //    // Obtem dados genéricos para salvar na MensagemSPB
        //    StatusMensagem = ObterStatusMensagem(entidadeMensagem.CodigoMensagem);
        //    TipoBackoffice = ObterTipoBackOffice(entidadeMensagem.CodigoMensagem);
        //    CodigoLocalLiquidacao = ObterCodigoLocalLiquidacao(entidadeMensagem.CodigoMensagem);

        //    // Obtem dados específicos para salvar na MensagemSPB
        //    base.ObterConteudoTagsEspecificas(entidadeMensagem.XmlMensagem);

        //    // Salva a Mensagem SPB
        //    base.Incluir(
        //        entidadeMensagem.XmlMensagem.InnerXml,
        //        entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ,
        //        TipoBackoffice,
        //        entidadeMensagem.CodigoMensagem,
        //        null,
        //        entidadeMensagem.CabecalhoMensagem.CodigoEmpresa,
        //        _NumeroComandoOperacao,
        //        null,
        //        StatusMensagem,
        //        A8NET.Comum.Comum.EnumInidicador.Nao,
        //        CodigoLocalLiquidacao,
        //        "", // CódigoVeiculoLegal
        //        entidadeMensagem.CabecalhoMensagem.SiglaSistemaEnviouNZ,
        //        1,
        //        null,
        //        null,
        //        _RegistroOperacaoCambial2
        //        );

        //    // Gerencia chamada
        //    GerenciarChamada(entidadeMensagem);

        //}
        //#endregion

//        #region <<< GerenciarChamada >>>
//        /// <summary>
//        /// O metodo basicamente altera o status da operação do R0 e envia a mensagem para o legado
//        /// </summary>
//        /// <param name="parametroOPER">daod da operação R0</param>
//        /// <param name="entidadeMensagem">mensagem recebida</param>
//        /// <param name="statusOperacao">status da operação</param>
//        /// <param name="enumEstorno">indica o estorno</param>
//        public void GerenciarChamada(udt.udtMensagem entidadeMensagem)
//        {
//            _XmlOperacaoAux = entidadeMensagem.XmlMensagem;

//            try
//            {

//                // Obtem dados necessários durante o processamento
//                _EventoProcessamento = ObterEventoProcessamento(entidadeMensagem.CodigoMensagem);
//                _TipoOperacao = (int)ObterTipoOperacao(entidadeMensagem.CodigoMensagem); if (_TipoOperacao == 0) return;
//                _TipoMensagemRetorno = DataSetCache.TB_TIPO_OPER.FindByTP_OPER(_TipoOperacao).TP_MESG_RETN_INTE.ToString();

//                // Retorna a MensagemSPB que acabou de ser gravada no banco
//                _MensagemSpbDATA.SelecionarMensagensPorControleIF(entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ);
//                if (_MensagemSpbDATA.Itens.Length == 0) return; // MensagemSPB não foi gravada, portanto encerra processamento
//                _EstruturaMensagemSPB = _MensagemSpbDATA.ObterMensagemLida();

//                // Obtem parametrização de processamento
//                DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.RowFilter = string.Format(@"NO_PROC_OPER_ATIV ='{0}' 
//                                                                                        AND TP_OPER={1}",
//                                                                                        _EventoProcessamento,
//                                                                                        _TipoOperacao.ToString());
//                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.Count == 0) return; // Mensagem sem parametrização de processamento, encerrar SEM DAR EXPLICAÇÕES

//                // Verifica se MensagemSPB gera Operação
//                if (entidadeMensagem.CodigoMensagem == "CAM0015")
//                {
//                    _NumeroSequenciaOperacao = GerarOperacao(entidadeMensagem, _TipoOperacao);
//                    if (_NumeroSequenciaOperacao == 0) return; // senão conseguir gerar operacao então aborta processamento
//                }
    
//                #region >>> Verifica se faz Conciliacao >>>
//                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_VERI_REGR_CNCL"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
//                {
//                    // Verifica conciliação MensagemSPB com Operação. Caso a conciliação esteja OK, 
//                    // o status da Operação e da MensagemSPB já serão atualizados dentro de ConciliacaoBO.VerificaConciliacao()
//                    if (_ConciliacaoBO.VerificaConciliacao(_EstruturaMensagemSPB, _IndicadorAceite, ref _NumeroSequenciaOperacao, ref _StatusOperacao) == false)
//                    {
//                        return; // Conciliação não OK, portanto encerrar processamento
//                    }

//                }
//                #endregion

//                #region >>> Verifica se Envia Retorno para o Legado >>>
//                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_MESG_RETN"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
//                {
//                    // Se houver uma Operacao associada/conciliada à Mensagem, então appenda algumas tags ref à Operacao, necessárias no retorno para o legado
//                    if (_NumeroSequenciaOperacao != 0)
//                    {
//                        // Retorna a Operação para obter CO_OPER_ATIV
//                        _XmlOperacaoR0 = base.ObterOperacaoXML((int)_NumeroSequenciaOperacao);

//                        // Appenda tags no XML de Retorno para o legado
//                        if (_XmlOperacaoAux.SelectSingleNode("//CO_OPER_ATIV") == null) Comum.Comum.AppendNode(ref _XmlOperacaoAux, "MESG", "CO_OPER_ATIV", _XmlOperacaoR0.DocumentElement.SelectSingleNode("//CO_OPER_ATIV").InnerXml);
//                        else _XmlOperacaoAux.SelectSingleNode("//CO_OPER_ATIV").InnerText = _XmlOperacaoR0.DocumentElement.SelectSingleNode("//CO_OPER_ATIV").InnerXml;
//                        if (_XmlOperacaoAux.SelectSingleNode("//CO_ULTI_SITU_PROC") == null) Comum.Comum.AppendNode(ref _XmlOperacaoAux, "MESG", "CO_ULTI_SITU_PROC", _StatusOperacao);
//                        else _XmlOperacaoAux.SelectSingleNode("//CO_ULTI_SITU_PROC").InnerText = _StatusOperacao;
//                    }

//                    // Envia retorno legado
//                    this.TratarRetorno(entidadeMensagem.XmlMensagem, _TipoMensagemRetorno, entidadeMensagem.CabecalhoMensagem.CodigoEmpresa, _NumeroSequenciaOperacao);

//                    // Para algumas mensagens específicas mudar status da MensagemSPB para Status específico
//                    if (entidadeMensagem.CodigoMensagem.Equals("CAM0015"))
//                    {
//                        base.AlterarStatusMensagemSPB(_EstruturaMensagemSPB, A8NET.Comum.Comum.EnumStatusMensagem.EnviadaLegado);
//                    }
//                    else if (entidadeMensagem.CodigoMensagem.Equals("CAM0055"))
//                    {
//                        base.AlterarStatusMensagemSPB(_EstruturaMensagemSPB, A8NET.Comum.Comum.EnumStatusMensagem.Registrada);
//                    }

//                }
//                #endregion

//                // Atualiza TipoBackoffice/CodigoVeiculoLegal/CodigoLocalLiquidacao da MensagemSPB, de acordo com a Operacao Conciliada/Associada
//                //
//                //
//                //

//            }
//            catch (Exception ex)
//            {
//                throw new Exception("MensagemSPB2.GerenciarChamada - " + ex.ToString());
//            }
//        }
//        #endregion

//        #region <<< TratarRetorno >>>
//        /// <summary>
//        /// Enviar mensagem de retorno para o Legado
//        ///   - Montar protocolo de integração A7
//        ///   - Montar Remessa
//        ///   - Incluir a mensagem de retorno na tabela de Mensagem Interna
//        /// </summary>
//        /// <param name="parametroOPER"></param>
//        /// <param name="entidadeMensagem"></param>
//        /// <returns></returns>
//        private void TratarRetorno(XmlDocument xmlMensagem, string tipoMsgRetornoInterno, string codigoEmpresa, long NumeroSequenciaOperacao)
//        {
//            OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna ParametroMsgInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
//            OperacaoMensagemInternaDAO OperacaoInternaDAO = new OperacaoMensagemInternaDAO();
//            TextXmlDAO TextXmlData = new TextXmlDAO();
//            DsParametrizacoes.TB_TIPO_OPERRow RowTipoOPER = null;
//            string Protocolo = string.Empty;
//            string TipoMensagem = "0";
//            int FormatoSaidaMsg = 0;
//            int CodigoTextXML = 0;

//            try
//            {

//                #region >>> Monta mensagem de retorno e coloca na fila de entrada do A7NET >>>
//                foreach (DataRow Row in base.DataSetCache.TB_REGR_SIST_DEST.Select("TP_MESG='" + tipoMsgRetornoInterno +
//                                                                                   "' AND SG_SIST_ORIG='A8'" +
//                                                                                   " AND CO_EMPR_ORIG=" + codigoEmpresa +
//                                                                                   " AND SG_SIST_DEST<>'R2'"))
//                {
//                    Protocolo = string.Concat(tipoMsgRetornoInterno.PadLeft(9, '0'), "A8 ", Row["SG_SIST_DEST"].ToString().ToUpper().PadRight(3, ' '), codigoEmpresa.PadLeft(5, '0'));

//                    xmlMensagem.DocumentElement.SelectSingleNode("TP_MESG").InnerXml = tipoMsgRetornoInterno.ToString().PadLeft(9, '0');
//                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "DT_MESG", DateTime.Today.ToString("yyyyMMdd"));
//                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "HO_MESG", DateTime.Now.ToString("HHmm"));
//                    //Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "CO_VEIC_LEGA", "46");
//                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "CO_MESG_SPB", xmlMensagem.SelectSingleNode("//CodMsg").InnerText);
//                    Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "TP_RETN", "1");

//                    MQConnector MqConnector = null;
//                    using (MqConnector = new MQConnector())
//                    {
//                        MqConnector.MQConnect();
//                        MqConnector.MQQueueOpen("A7Q.E.ENTRADA_NET", MQConnector.enumMQOpenOptions.PUT);
//                        MqConnector.Message = Protocolo + xmlMensagem.InnerXml;
//                        MqConnector.MQPutMessage();
//                        MqConnector.MQQueueClose();
//                        MqConnector.MQEnd();
//                    }

//                    // Obtem FormatoMensagemSaida
//                    DataRow[] RowREGR = DataSetCache.TB_REGR_SIST_DEST.Select("TP_MESG='" + tipoMsgRetornoInterno + "' AND SG_SIST_DEST='" + Row["SG_SIST_DEST"].ToString().ToUpper().Trim() + "'", "DH_INIC_VIGE_REGR_TRAP DESC");
//                    if (RowREGR.Length > 0) int.TryParse(RowREGR[0]["TP_FORM_MESG_SAID"].ToString(), out FormatoSaidaMsg);

//                }
//                #endregion

//                #region >>> Gera registro na tabela OperacaoMensagemInterna caso a MensagemSPB seja conciliada/associada com Operacao >>>
//                if (NumeroSequenciaOperacao != 0)
//                {
//                    // Armazena a mensagem original na tabela TB_TEXT_XML 
//                    CodigoTextXML = TextXmlData.InserirBase64(xmlMensagem.OuterXml);

//                    // preencher os dados da operação mensagem interna
//                    ParametroMsgInterna.NU_SEQU_OPER_ATIV = NumeroSequenciaOperacao;
//                    ParametroMsgInterna.TP_MESG_INTE = tipoMsgRetornoInterno; //TipoMensagemOriginal;
//                    ParametroMsgInterna.TP_FORM_MESG_SAID = FormatoSaidaMsg;
//                    ParametroMsgInterna.TP_SOLI_MESG_INTE = (int)Comum.Comum.enumTipoSolicitacao.RetornoLegado;
//                    ParametroMsgInterna.CO_TEXT_XML = CodigoTextXML;
//                    ParametroMsgInterna.DH_MESG_INTE = OperacaoInternaDAO.ObterDataGravacao(NumeroSequenciaOperacao.ToString()).AddSeconds(1);
//                    // insere os dados no banco
//                    OperacaoInternaDAO.Inserir(ParametroMsgInterna);
//                }
//                #endregion

//            }
//            catch (Exception ex)
//            {
//                throw new Exception("MensagemSPB2.TratarRetorno() - " + ex.ToString());
//            }
//        }
//        #endregion

        #region >>> ObterEventoProcessamento >>>
        public override string ObterEventoProcessamento(string codigoMensagem)
        {
            DataRow[] Retorno = DataSetCache.TB_MENSAGEM.Select("CO_MESG='" + codigoMensagem.Trim() + "'");
            if (Retorno.Length == 0)
            {
                return "RecebimentoInformacao";
            }
            else
            {
                if (int.Parse(Retorno[0]["SQ_TIPO_FLUX"].ToString()) == (int)Comum.Comum.EnumTipoFluxo.TipoFluxo5)
                    return "RecebimentoInformacao";
                else
                    return "RecebimentoAviso";
            }
        }
        #endregion

    }
}
