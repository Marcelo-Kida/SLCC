using System;
using System.Collections.Generic;
using System.Text;
using A8NET.Data.DAO;
using A8NET.Historico;
using A8NET.ConfiguracaoMQ;
using A8NET.Data;
using A8NET.Mensagem;
using System.Data;
using System.Xml;
using System.Configuration;

namespace A8NET.Mensagem.SPB
{
    public class MensagemSPBR1 : MensagemSPB, IDisposable
    {

        #region <<< Variáveis >>>
        #endregion

        #region <<< Construtor >>>
        public MensagemSPBR1(Data.DsParametrizacoes dsCache): base(dsCache)
        {
            _MensagemSpbDATA = new MensagemSpbDAO();
            _OperacaoDATA = new OperacaoDAO();
            _ParametroOPER = new OperacaoDAO.EstruturaOperacao();
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~MensagemSPBR1()
        {
            this.Dispose();
        }
        #endregion

        #region <<< OVERRIDE - GerenciaMensagem >>>
        /// <summary>
        /// Método trata a mensagem conforme o tipo da mensagem
        /// </summary>
        /// <param name="udtMsg">udt da mensagem (header, linha com os dados, xml original, fila)</param>
        public override void GerenciaMensagem(udt.udtMensagem entidadeMensagem)
        {
            if (base.VerificarMensagemTratadaSLCC(entidadeMensagem.CodigoMensagem))
            {
                this.ProcessaMensagemTratada(entidadeMensagem);
            }
            else
            {
                this.ProcessaMensagemNaoTratada(entidadeMensagem);
            }
        }
        #endregion

        #region <<< ProcessaMensagem >>>
        public override void ProcessaMensagem(string nomeFila, string mensagemRecebida)
        {
            base.ProcessaMensagem(nomeFila, mensagemRecebida);
        }
        #endregion

        #region <<< ProcessaMensagemTratada >>>
        /// <summary>
        /// Processamento de mensagens R1 geradas por uma remessa do legado
        /// </summary>
        private void ProcessaMensagemTratada(udt.udtMensagem entidadeMensagem)
        {
            OperacaoDAO OperacaoDATA = new OperacaoDAO();
            MensagemSpbDAO.EstruturaMensagemSPB EstruturaMensagemSpbR0 = new MensagemSpbDAO.EstruturaMensagemSPB();
            TextXmlDAO TextXmlData = new TextXmlDAO();
            EstruturaStatusOperacaoMensagem EstruturaStatus = new EstruturaStatusOperacaoMensagem();
            XmlDocument XmlOperacaoR0 = new XmlDocument();
            Comum.Comum.EnumStatusMensagem StatusMensagem;
            DateTime DataGravacao = new DateTime();

            try
            {
                // selecionar as operações pelo controle do IF
                _MensagemSpbDATA.SelecionarMensagensPorControleIF(entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ);
                if (_MensagemSpbDATA.Itens.Length == 0)
                {
                    RemessaRejeitadaDAO.EstruturaRemessaRejeitada ParametroRemessaRejeitada = new RemessaRejeitadaDAO.EstruturaRemessaRejeitada();
                    Data.DAO.RemessaRejeitadaDAO _RemessaRejeitadaDATA = new Data.DAO.RemessaRejeitadaDAO();
                    int CodigoTextXML = 0;
                    CodigoTextXML = TextXmlData.InserirBase64(entidadeMensagem.XmlMensagem.InnerXml);
                    ParametroRemessaRejeitada.SG_SIST_ORIG_INFO = entidadeMensagem.XmlMensagem.SelectSingleNode("//SG_SIST_ORIG").InnerXml;
                    ParametroRemessaRejeitada.TP_MESG_INTE = int.Parse(entidadeMensagem.XmlMensagem.SelectSingleNode("//TP_MESG").InnerXml);
                    ParametroRemessaRejeitada.CO_EMPR = int.Parse(entidadeMensagem.CabecalhoMensagem.CodigoEmpresa);
                    ParametroRemessaRejeitada.CO_TEXT_XML_REJE = CodigoTextXML;
                    ParametroRemessaRejeitada.CO_TEXT_XML_RETN_SIST_ORIG = CodigoTextXML;
                    ParametroRemessaRejeitada.TX_XML_ERRO = Comum.Comum.Base64Encode(@"<Erro><Grupo_ErrorInfo>
                                                              <Number>3022</Number>
                                                              <Description>N&#250;mero de Controle IF da Mensagem R1 inv&#225;lido.</Description>
                                                              <ComputerName>" + Comum.Comum.NomeMaquina + @"</ComputerName>
                                                              <Source>A8NET.Mensagem.SPB.MensagemSPBR1.ProcessaMensagemTratada()</Source>
                                                              <ErrorType>1</ErrorType>
                                                              </Grupo_ErrorInfo></Erro>"); //<Time>" + DateTime.Now + @"</Time>
                    ParametroRemessaRejeitada.DH_REME_REJE = DateTime.Now;
                    _RemessaRejeitadaDATA.Inserir(ParametroRemessaRejeitada);

                    return;
                }
                
                // retorna a mensagem de R0 que já está no banco
                EstruturaMensagemSpbR0 = _MensagemSpbDATA.ObterMensagemLida();

                // retorna a operação da mensagem de R0 da R1
                XmlOperacaoR0 = base.ObterOperacaoXML(int.Parse(EstruturaMensagemSpbR0.NU_SEQU_OPER_ATIV.ToString()));
                
                if (DataSetCache.TB_TIPO_OPER.FindByTP_OPER(decimal.Parse(XmlOperacaoR0.DocumentElement.SelectSingleNode("TP_OPER").InnerXml)) != null)
                {              
                    Comum.Comum.AppendNode(ref XmlOperacaoR0, "MESG", "TP_MESG_RETN_INTE", DataSetCache.TB_TIPO_OPER.FindByTP_OPER(decimal.Parse(XmlOperacaoR0.DocumentElement.SelectSingleNode("TP_OPER").InnerXml)).TP_MESG_RETN_INTE);
                }

                //obtem conteúdo de tags específicas
                _DataOperacaoCambioSisbacen = null;
                base.ObterConteudoTagsEspecificas(entidadeMensagem.XmlMensagem); 

                // Inclui a Mensagem R1
                StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.R1;
                DataGravacao = base.Incluir(
                         entidadeMensagem.XmlMensagem.InnerXml,
                         entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ,
                         int.Parse(EstruturaMensagemSpbR0.TP_BKOF.ToString()),
                         entidadeMensagem.CodigoMensagem,
                         int.Parse(EstruturaMensagemSpbR0.NU_SEQU_OPER_ATIV.ToString()),
                         entidadeMensagem.CabecalhoMensagem.CodigoEmpresa,
                         _NumeroComandoOperacao,
                         null, // situação mensagem SPB
                         StatusMensagem,
                         Comum.Comum.EnumInidicador.Nao, // Indicador Entrada Manual
                         int.Parse(XmlOperacaoR0.DocumentElement.SelectSingleNode("CO_LOCA_LIQU").InnerXml),
                         XmlOperacaoR0.DocumentElement.SelectSingleNode("CO_VEIC_LEGA").InnerXml,
                         XmlOperacaoR0.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml,
                         1, // controle repetição
                         _DataOperacaoCambioSisbacen,
                         _CodigoClienteSisbacen,
                         _RegistroOperacaoCambial2
                         );
                        

                // Obtem os status e a situação da mensagem e da operação
                EstruturaStatus = base.ObterStatus(entidadeMensagem);

                switch (entidadeMensagem.CodigoMensagem.Trim())
                {
                    case "CAM0005R1":
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.Reativada;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Respondida;
                        if (int.Parse(XmlOperacaoR0.SelectSingleNode("//TP_OPER").InnerText) == (int)Comum.Comum.EnumTipoOperacao.InformaOperacaoArbitragemParceiroPais)
                        {
                            StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Reativada;
                        }
                        break;
                    case "CAM0006R1":
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.Respondida;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Respondida;
                        break;
                    case "CAM0007R1": case "CAM0010R1":
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.Confirmada;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Respondida;
                        break;
                    case "CAM0009R1":
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.Registrada;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Respondida;
                        break;
                    case "CAM0013R1": 
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.Registrada;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Registrada;
                        break;
                    case "CAM0014R1":
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.Confirmada;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Registrada;
                        break;
                    case "CAM0054R1":
                        EstruturaStatus.StatusOperacao = (int)A8NET.Comum.Comum.EnumStatusOperacao.AConciliarAceite;
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Respondida;
                        break;
                    default:
                        StatusMensagem = A8NET.Comum.Comum.EnumStatusMensagem.Registrada;
                        break;
                }

                EstruturaMensagemSpbR0.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                EstruturaMensagemSpbR0.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                base.AlterarStatusMensagemSPB(ref EstruturaMensagemSpbR0, 
                                                  StatusMensagem);

                // Atualizar o status da Operação enviar a mensagem pelo MQ
                this.GerenciarChamada(
                    XmlOperacaoR0,
                    entidadeMensagem,
                    EstruturaStatus.StatusOperacao,
                    Comum.Comum.EnumInidicador.Nao);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro:ProcessaMensagemTratada() - DESC:" + ex.ToString());
            }
            finally {
                OperacaoDATA.Dispose();
            }
        }
        #endregion

        #region <<< GerenciarChamada >>>
        /// <summary>
        /// O metodo basicamente altera o status da operação do R0 e envia a mensagem para o legado
        /// </summary>
        /// <param name="parametroOPER">daod da operação R0</param>
        /// <param name="entidadeMensagem">mensagem recebida</param>
        /// <param name="statusOperacao">status da operação</param>
        /// <param name="enumEstorno">indica o estorno</param>
        public void GerenciarChamada(XmlDocument xmlOperacao, udt.udtMensagem entidadeMensagem, int statusOperacao,
            Comum.Comum.EnumInidicador enumEstorno)
        {
            string TipoOperacao = xmlOperacao.DocumentElement.SelectSingleNode("TP_OPER").InnerXml;
            string CodigoOperacao = xmlOperacao.DocumentElement.SelectSingleNode("NU_SEQU_OPER_ATIV").InnerXml;
            string Mensagem = string.Empty;
            bool AtualizarOperacao = false;
            string XMLMensagemSPBConciliada = string.Empty;
            udt.udtMensagem EntidadeMensagemConciliada = new udt.udtMensagem();
            MensagemSpbDAO MensagemR1DATA = new MensagemSpbDAO();
            DsTB_MESG_RECB_ENVI_SPB DsMensagemSPB = new DsTB_MESG_RECB_ENVI_SPB();

            try
            {

                // Retorna a MensagemSPB R1 que acabou de ser gravada no banco no método anterior ProcessaMensagemTratada()
                MensagemR1DATA.SelecionarMensagensPorControleIF(entidadeMensagem.CabecalhoMensagem.ControleRemessaNZ);
                if (MensagemR1DATA.Itens.Length == 0) return; // MensagemSPB não foi gravada, portanto encerrar processamento
                _EstruturaMensagemSPB = MensagemR1DATA.ObterUltimaMensagemPorNumeroControleIF();

                // Alterar o status da Operação
                if (statusOperacao > 0 && base.AtualizaStatusOperacao(entidadeMensagem.CodigoMensagem, long.Parse(CodigoOperacao)) == true)
                {
                    _ParametroOPER.CO_ULTI_SITU_PROC = statusOperacao;
                    base.AlterarStatusOperacao(int.Parse(CodigoOperacao), statusOperacao, 0, 0);
                }

                #region >>> Grava Registro Operação Cambial e Registro Operação Cambial 2 na Operação >>>
                if (entidadeMensagem.CodigoMensagem == "CAM0006R1"
                 || entidadeMensagem.CodigoMensagem == "CAM0009R1"
                 || entidadeMensagem.CodigoMensagem == "CAM0013R1"
                 || entidadeMensagem.CodigoMensagem == "CAM0054R1")
                {
                    _OperacaoDATA.ObterOperacao(int.Parse(CodigoOperacao));
                    _ParametroOPER = _OperacaoDATA.TB_OPER_ATIV;
                    if (_NumeroComandoOperacao != string.Empty)
                    {
                        _ParametroOPER.NU_COMD_OPER = _NumeroComandoOperacao;
                        AtualizarOperacao = true;
                    }
                    if (_RegistroOperacaoCambial2 != string.Empty)
                    {
                        _ParametroOPER.NR_OPER_CAMB_2 = _RegistroOperacaoCambial2;
                        AtualizarOperacao = true;
                    }
                    if (AtualizarOperacao == true)
                    {
                        _ParametroOPER.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                        _ParametroOPER.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                        _OperacaoDATA.Atualizar(_ParametroOPER);
                    }
                }
                #endregion

                // Inclui/Atualiza conteúdo da tag CO_ULTI_SITU_PROC
                if (xmlOperacao.SelectSingleNode("//CO_ULTI_SITU_PROC") == null) Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "CO_ULTI_SITU_PROC", statusOperacao.ToString());
                else xmlOperacao.SelectSingleNode("//CO_ULTI_SITU_PROC").InnerText = statusOperacao.ToString();

                //Sistema verifica se a mensagem de requisição foi originada na entrada manual do A8 e se SistemaOrigem = GPC ou R2 (sistemas do Comex).
                //Se afirmativo, sistema retorna mensagem de requisição R0 ao sistema legado
                if (xmlOperacao.SelectSingleNode("//IN_ENTR_MANU") != null)
                {
                    if (xmlOperacao.DocumentElement.SelectSingleNode("IN_ENTR_MANU").InnerXml == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                    {
                        if (xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml.Trim() == "GPC" || xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml.Trim() == "R2" || xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml.Trim() == "BOL" || xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml.Trim() == "HQ")
                        {
                            EnviarLegadoOperacaoEntradaManual(xmlOperacao, entidadeMensagem.XmlMensagem);
                        }
                        else // sistema NÃO é nem GPC nem R2
                        {
                            return; // qualquer outro sistema que não seja o GPC ou R2 NÃO deve retornar nada para o legado quando a Operação foi gerada pela Entrada Manual
                        }
                    }

                }

                if (entidadeMensagem.CodigoGrupoMensagem == "CAM") statusOperacao = (int)Comum.Comum.EnumStatusOperacao.Inicial;
                // retorna a parametrização 
                DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.RowFilter = string.Format(@"NO_PROC_OPER_ATIV ='RecebimentoR1' AND TP_OPER={0} 
                        AND IN_ESTO_PJ_A6={1} AND CO_SITU_PROC={2}", TipoOperacao, (int)enumEstorno, statusOperacao);

                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView.Count == 0) return; //  sair do método SEM DAR EXPLICAÇÕES

                if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.DefaultView[0]["IN_ENVI_MESG_RETN"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {
                    Mensagem = this.TratarRetorno(ref xmlOperacao, entidadeMensagem.XmlMensagem);

                    //MQConnector.PutMensagem(ConfigurationManager.AppSettings["FilaA7Entrada"].ToString(), Mensagem);
                    //MQConnector.PutMensagem("A7Q.E.ENTRADA_NET", Mensagem);

                    MQConnector _MqConnector = null;
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

                //#region >>> Tratamento de mensagens processadas fora da ordem que era esperada pelo Fluxo >>>
                //// se a mensagem for CAM0009R1/CAM0013R1/CAM0054R1, então verifica se já foi recebida antes dela uma MensagemR2/Informacao do mesmo fluxo,
                //// se sim, processa novamente a MensagemR2/Informacao para que ela seja retornada para o legado
                //if (entidadeMensagem.CodigoMensagem == "CAM0013R1"
                // || entidadeMensagem.CodigoMensagem == "CAM0009R1"
                // || entidadeMensagem.CodigoMensagem == "CAM0054R1")
                //{
                //    if (_ConciliacaoBO.VerificaConciliacao(_EstruturaMensagemSPB, ref DsMensagemSPB) == true)
                //    {
                //        XMLMensagemSPBConciliada = base.MontaMensagemNZEntradaA8(DsMensagemSPB.TB_MESG_RECB_ENVI_SPB.DefaultView);
                //        EntidadeMensagemConciliada.Parse(XMLMensagemSPBConciliada);

                //        if (EntidadeMensagemConciliada.CodigoMensagem == "BMC0005")
                //        {
                //            MensagemAvisoInformacao MensagemAvisoInformacao = new MensagemAvisoInformacao(DataSetCache);
                //            MensagemAvisoInformacao.GerenciarChamada(EntidadeMensagemConciliada);
                //        }
                //        else
                //        {
                //            MensagemSPBR2 MensagemR2 = new MensagemSPBR2(DataSetCache);
                //            MensagemR2.GerenciarChamada(EntidadeMensagemConciliada);
                //        }
                //    }
                //}
                //#endregion

            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPB.GerenciarChamada - " + ex.ToString());
            }
            finally
            {
                #region >>> Tratamento de mensagens processadas fora da ordem que era esperada pelo Fluxo >>>
                // se a mensagem for CAM0009R1/CAM0013R1/CAM0054R1, então verifica se já foi recebida antes dela uma MensagemR2/Informacao do mesmo fluxo,
                // se sim, processa novamente a MensagemR2/Informacao para que ela seja retornada para o legado
                if (entidadeMensagem.CodigoMensagem == "CAM0013R1"
                 || entidadeMensagem.CodigoMensagem == "CAM0009R1"
                 || entidadeMensagem.CodigoMensagem == "CAM0054R1")
                {
                    if (_ConciliacaoBO.VerificaConciliacao(_EstruturaMensagemSPB, ref DsMensagemSPB) == true)
                    {
                        XMLMensagemSPBConciliada = base.MontaMensagemNZEntradaA8(DsMensagemSPB.TB_MESG_RECB_ENVI_SPB.DefaultView);
                        EntidadeMensagemConciliada.Parse(XMLMensagemSPBConciliada);

                        if (EntidadeMensagemConciliada.CodigoMensagem == "BMC0005")
                        {
                            MensagemAvisoInformacao MensagemAvisoInformacao = new MensagemAvisoInformacao(DataSetCache);
                            MensagemAvisoInformacao.GerenciarChamada(EntidadeMensagemConciliada);
                        }
                        else
                        {
                            MensagemSPBR2 MensagemR2 = new MensagemSPBR2(DataSetCache);
                            MensagemR2.GerenciarChamada(EntidadeMensagemConciliada);
                        }
                    }
                }
                #endregion
            }
        }
        #endregion

        #region <<< TratarRetorno() >>>
        /// <summary>
        /// Enviar mensagem de retorno para o Legado
        ///   - Montar protocolo de integração A7
        ///   - Montar Remessa
        ///   - Incluir a mensagem de retorno na tabela de Mensagem Interna
        /// </summary>
        /// <param name="parametroOPER"></param>
        /// <param name="entidadeMensagem"></param>
        /// <returns></returns>
        private string TratarRetorno(ref XmlDocument xmlOperacao, XmlDocument xmlMensagem)
        {
            OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna ParametroMsgInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
            OperacaoMensagemInternaDAO OperacaoInternaDAO = new OperacaoMensagemInternaDAO();
            TextXmlDAO TextXmlData = new TextXmlDAO();
            DsParametrizacoes.TB_TIPO_OPERRow RowTipoOPER = null;
            string TipoMensagemOriginal = xmlOperacao.DocumentElement.SelectSingleNode("TP_OPER").InnerXml;
            string CodigoOPER = xmlOperacao.DocumentElement.SelectSingleNode("NU_SEQU_OPER_ATIV").InnerXml;
            string Protocolo = string.Empty;
            int FormatoSaidaMsg = 0;
            string TipoMensagem = "0";
            int CodigoTextXML = 0;

            try
            {
                if (xmlOperacao.SelectSingleNode("//CO_ERRO1") != null)
                {
                    if (xmlOperacao.SelectSingleNode("//CO_ERRO1").InnerXml == "4007")// 'Veículo Legal inválido"
                    {
                        xmlOperacao.SelectSingleNode("//CO_ERRO1").InnerXml = null;
                        if (xmlOperacao.SelectSingleNode("//DE_ERRO1") != null)
                        {
                            xmlOperacao.SelectSingleNode("//DE_ERRO1").InnerXml = null;
                        }
                    }
                }

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
                Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "DT_MESG", DateTime.Today.ToString("yyyyMMdd"));
                Comum.Comum.AppendNode(ref xmlMensagem, "MESG", "HO_MESG", DateTime.Now.ToString("HHmm"));

                Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "MESG", xmlMensagem.DocumentElement.InnerXml);
                Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "TP_RETN", "1");

                if (xmlOperacao.SelectSingleNode("//CO_MESG_SPB") == null) Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "CO_MESG_SPB", xmlMensagem.SelectSingleNode("//CodMsg").InnerText);
                else xmlOperacao.SelectSingleNode("//CO_MESG_SPB").InnerText = xmlMensagem.SelectSingleNode("//CodMsg").InnerText;

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

        #region <<< ProcessaMensagemNaoTratada >>>
        private void ProcessaMensagemNaoTratada(udt.udtMensagem entidadeMensagem)
        {
        }
        #endregion

        #region <<< EnviarLegadoOperacaoEntradaManual >>>
        /// <summary>
        /// O metodo basicamente altera o status da operação do R0 e envia a mensagem para o legado
        /// </summary>
        /// <param name="parametroOPER">dado da operação R0</param>
        /// <param name="entidadeMensagem">mensagem recebida</param>
        /// <param name="statusOperacao">status da operação</param>
        /// <param name="enumEstorno">indica o estorno</param>
        public void EnviarLegadoOperacaoEntradaManual(XmlDocument xmlOperacao, XmlDocument xmlMensagem)
        {
            string Mensagem = string.Empty;
            string Protocolo = string.Empty;
            string TipoMensagem = xmlOperacao.DocumentElement.SelectSingleNode("TP_MESG").InnerXml;
            int CodigoTextXML = 0;
            OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna ParametroMsgInterna = new OperacaoMensagemInternaDAO.EstruturaOperacaoMensagemInterna();
            OperacaoMensagemInternaDAO OperacaoInternaDAO = new OperacaoMensagemInternaDAO();
            TextXmlDAO TextXmlData = new TextXmlDAO();
            string CodigoOPER = xmlOperacao.DocumentElement.SelectSingleNode("NU_SEQU_OPER_ATIV").InnerXml;
            int FormatoSaidaMsg = 0;

            Protocolo = string.Concat(TipoMensagem.PadLeft(9, '0'),
                    "A8 ",
                    xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml.Trim().PadRight(3, ' '),
                    xmlOperacao.DocumentElement.SelectSingleNode("CO_EMPR").InnerXml.PadLeft(5, '0'));

            DataRow[] RowREGR = DataSetCache.TB_REGR_SIST_DEST.Select("TP_MESG='" + TipoMensagem + "' AND SG_SIST_DEST='" + xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml + "'", "DH_INIC_VIGE_REGR_TRAP DESC");
            if (RowREGR.Length > 0) int.TryParse(RowREGR[0]["TP_FORM_MESG_SAID"].ToString(), out FormatoSaidaMsg);

            // Corrige tags SG_SIST_ORIG e SG_SIST_DEST
            xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml = xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml;
            xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml = "A8";

            // Armazena a mensagem original na tabela TB_TEXT_XML 
            CodigoTextXML = TextXmlData.InserirBase64(xmlOperacao.OuterXml);

            // Preencher os dados da operação mensagem interna
            ParametroMsgInterna.NU_SEQU_OPER_ATIV = CodigoOPER;
            ParametroMsgInterna.TP_MESG_INTE = TipoMensagem; //TipoMensagemOriginal;
            ParametroMsgInterna.TP_FORM_MESG_SAID = FormatoSaidaMsg;
            ParametroMsgInterna.TP_SOLI_MESG_INTE = (int)Comum.Comum.enumTipoSolicitacao.RetornoLegado;
            ParametroMsgInterna.CO_TEXT_XML = CodigoTextXML;
            ParametroMsgInterna.DH_MESG_INTE = OperacaoInternaDAO.ObterDataGravacao(CodigoOPER);
            // inseri os dados no banco
            OperacaoInternaDAO.Inserir(ParametroMsgInterna);

            MQConnector _MqConnector = null;
            using (_MqConnector = new MQConnector())
            {
                _MqConnector.MQConnect();
                _MqConnector.MQQueueOpen("A7Q.E.ENTRADA_NET", MQConnector.enumMQOpenOptions.PUT);
                _MqConnector.Message = Protocolo + xmlOperacao.OuterXml;
                _MqConnector.MQPutMessage();
                _MqConnector.MQQueueClose();
                _MqConnector.MQEnd();
            }

            //volta o conteúdo original das tags SG_SIST_ORIG e SG_SIST_DEST 
            xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerXml = xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml;
            xmlOperacao.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerXml = "A8";
        }
        #endregion

        #region >>> ObterEventoProcessamento >>>
        public override string ObterEventoProcessamento(string codigoMensagem)
        {
            return string.Empty;
        }
        #endregion

    }
}
