using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using A7NET.Data;
using A7NET.Comum;
using A7NET.ConfiguracaoMQ;

namespace A7NET.Mensagem
{
    public abstract class Mensagem
    {
        #region <<< Metodos Abstract >>>
        public abstract void ProcessaMensagem(string nomeFila, string mensagemRecebida, string messageId);
        #endregion

        #region >>> Construtor >>>
        public Mensagem()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion

            _XmlMensagem = new XmlDocument();

        }
        #endregion

        #region <<< Variaveis >>>
        protected DateTime _DataHoraInicioRegra;
        protected XmlDocument _XmlMensagem;
        protected A7NET.Mensagem.udt.udtProtocoloMensagem _ProtocoloMensagem;
        protected bool _UtilizaRepeticao;
        protected int _TipoFormatoMensagemSaida;
        private A7NET.ConfiguracaoMQ.MQConnector _MqConnector = null;
        private DsParametrizacoes _DataSetCache;
        private DataRow _Regra;
        private bool _XmlValido = true;
        private string _XmlParseError;
        #endregion

        #region <<< Propriedades >>>
        public DsParametrizacoes DataSetCache
        {
            get { return _DataSetCache; }
            set { _DataSetCache = value; }
        }
        #endregion

        #region <<< HandleValidationError >>>
        private void HandleValidationError(object sender, ValidationEventArgs e)
        {
            _XmlValido = false;
            _XmlParseError = e.Message;

        }
        #endregion

        #region <<< ObterPropriedades >>>
        protected string ObterPropriedades(string nomeXml)
        {
            XmlDocument Propriedades;
            XmlNode NodePropriedades;

            try
            {
                Propriedades = new XmlDocument();
                NodePropriedades = Propriedades.CreateElement(nomeXml);
                Propriedades.AppendChild(NodePropriedades);

                foreach (DataColumn Coluna in DataSetCache.TB_MESG.Columns)
                {
                    NodePropriedades = null;
                    NodePropriedades = Propriedades.CreateElement(Coluna.ColumnName);
                    Propriedades.DocumentElement.AppendChild(NodePropriedades);
                }

                return Propriedades.InnerXml;

            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        #endregion

        #region <<< AutenticarMensagem >>>
        protected bool AutenticarMensagem()
        {
            ValidaMensagem AutenticaMensagem;
            string DetalheOcorrenciaErro;
            string RetornoErro = "";
            int OutN;

            try
            {
                AutenticaMensagem = new ValidaMensagem(DataSetCache);

                if (!AutenticaMensagem.ValidaTipoMensagem(_TipoFormatoMensagemSaida, _ProtocoloMensagem.TipoMensagem))
                {
                    DetalheOcorrenciaErro = "Tipo de Mensagem não cadastrado." + "\r\n" +
                                            "Código do Tipo de Mensagem : " + _ProtocoloMensagem.TipoMensagem;

                    _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                    return false;

                }

                if (!int.TryParse(_ProtocoloMensagem.CodigoEmpresa, out OutN))
                {
                    DetalheOcorrenciaErro = "Código de Empresa inválido." + "\r\n" +
                                            "Código da Empresa : " + _ProtocoloMensagem.CodigoEmpresa;

                    _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                    return false;
                }

                if (!AutenticaMensagem.ValidaEmpresa(_ProtocoloMensagem.CodigoEmpresa))
                {
                    DetalheOcorrenciaErro = "Código de Empresa inválido." + "\r\n" +
                                            "Código da Empresa : " + _ProtocoloMensagem.CodigoEmpresa;

                    _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                    return false;
                }

                if (!AutenticaMensagem.ValidaSistema(_ProtocoloMensagem.CodigoEmpresa, _ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper()))
                {
                    DetalheOcorrenciaErro = "Sistema Origem não cadastrado." + "\r\n" +
                                            "Código da Empresa : " + _ProtocoloMensagem.CodigoEmpresa + "\r\n" +
                                            "Sigla Sistema     : " + _ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper();

                    _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                    return false;
                }

                if (!AutenticaMensagem.ValidaSistema(_ProtocoloMensagem.CodigoEmpresa, _ProtocoloMensagem.SiglaSistemaDestino))
                {
                    DetalheOcorrenciaErro = "Sistema Destino não cadastrado." + "\r\n" +
                                            "Código da Empresa : " + _ProtocoloMensagem.CodigoEmpresa + "\r\n" +
                                            "Sigla Sistema     : " + _ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper();

                    _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                    return false;
                }

                if (!AutenticaMensagem.ObtemRegraTransporte(ref _Regra, _ProtocoloMensagem.TipoMensagem, _ProtocoloMensagem.CodigoEmpresa,
                                                            _ProtocoloMensagem.SiglaSistemaOrigem, _ProtocoloMensagem.SiglaSistemaDestino, ref RetornoErro))
                {
                    if (RetornoErro == "")
                    {
                        DetalheOcorrenciaErro = "Regra de Transporte não cadastrada. (" + _DataSetCache.TB_REGR_TRAP_MESG.Rows.Count + ") \r\n" +
                                                "Código do Tipo de Mensagem  : " + _ProtocoloMensagem.TipoMensagem + "\r\n" +
                                                "Sigla Sistema               : " + _ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() + "\r\n" +
                                                "Código da Empresa           : " + _ProtocoloMensagem.CodigoEmpresa;
                    }
                    else
                    {
                        DetalheOcorrenciaErro = RetornoErro;
                    }

                    _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                    return false;
                }

                _TipoFormatoMensagemSaida = int.Parse(_Regra["TP_FORM_MESG_SAID"].ToString());
                _XmlMensagem.DocumentElement.SelectSingleNode("TP_FORM_MESG_SAID").InnerText = _Regra["TP_FORM_MESG_SAID"].ToString();
                _DataHoraInicioRegra = Convert.ToDateTime(_Regra["DH_INIC_VIGE_REGR_TRAP"]);
                _XmlMensagem.DocumentElement.SelectSingleNode("DH_INIC_VIGE_REGR_TRAP").InnerText = string.Format("{0:yyyyMMddHHmmss}", _DataHoraInicioRegra);

                return true;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< TraduzirMensagem >>>
        protected bool TraduzirMensagem()
        {
            A7NET.Data.MensagemDAO MensagemDAO = new A7NET.Data.MensagemDAO();
            ConfiguraMensagem MontaMensagem = new ConfiguraMensagem();
            XmlDocument XmlRegra;
            XmlDocument XmlAux;
            XmlDocument XmlMensagemEntrada;
            XmlDocument XmlMensagemTraduzida;
            XmlDocument XmlMensagemSaida;
            XmlDocument XmlParametroGeral;
            XmlSchemaSet SchemaSet;
            TextReader SchemaReader;
            StringBuilder Protocolo = new StringBuilder();
            ValidationEventHandler EventHandler;
            string DetalheOcorrenciaErro;
            string MensagemTraduzida;
            string MensagemErro;
            string CampoErro;
            string ParametroGeral;
            string TipoMensagem;
            string HeaderNZ;
            int OutN;

            try
            {
                XmlRegra = new XmlDocument();
                XmlRegra.LoadXml(MontaMensagem.MontaXmlRegraTransporte(int.Parse(_Regra["CO_TEXT_XML_REGR"].ToString()),
                                                                       int.Parse(_Regra["CO_TEXT_XML_MESG"].ToString())));

                if (_Regra["IN_EXIS_REGR_TRNF"].ToString() == "S")
                {
                    if (XmlRegra.DocumentElement.SelectSingleNode("TX_REGR_TRNF_MESG").ToString().Trim() == string.Empty)
                    {
                        DetalheOcorrenciaErro = "Regra não cadastrada ou corrompida para a mensagem." + "\r\n" +
                                                "Tipo Mensagem   : " + _ProtocoloMensagem.TipoMensagem + "\r\n" +
                                                "Código Empresa  : " + _ProtocoloMensagem.CodigoEmpresa + "\r\n" +
                                                "Sistema Origem  : " + _ProtocoloMensagem.SiglaSistemaOrigem + "\r\n" +
                                                "Sistema Destino : " + _ProtocoloMensagem.SiglaSistemaDestino;

                        _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                        return false;

                    }

                    XmlAux = new XmlDocument();
                    A7NET.Comum.Comum.AppendNode(ref XmlAux, "", "Documento", "");
                    A7NET.Comum.Comum.AppendNode(ref XmlAux, "Documento", "Mensagem", "");
                    A7NET.Comum.Comum.AppendAttribute(ref XmlAux, "Mensagem", "Tipo", A7NET.Comum.Comum.TipoEntradaToSTR(int.Parse(_Regra["TP_FORM_MESG_ENTR"].ToString())));
                    A7NET.Comum.Comum.AppendAttribute(ref XmlAux, "Mensagem", "Delimitador", _Regra["TP_CTER_DELI_SAID"].ToString());

                    //Se a mensagem de entrada for do tipo XML
                    if (int.Parse(_Regra["TP_FORM_MESG_ENTR"].ToString()) == (int)A7NET.Comum.Comum.EnumTipoEntradaMensagem.EntradaXML)
                    {
                        #region Mensagem XML
                        //Faz Load da Mensagem de Entrada
                        XmlMensagemEntrada = new XmlDocument();
                        try
                        {
                            XmlMensagemEntrada.LoadXml(_XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_ENTR").InnerXml.Substring(20));
                        }
                        catch (Exception ex)
                        {
                            DetalheOcorrenciaErro = "XML de entrada inválido." + "\r\n" +
                                                    "Parser Error Reason: " + ex.ToString();

                            _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                            return false;
                            
                        }

                        // Para Mensagens MDR
                        if (XmlMensagemEntrada.ChildNodes[0].NodeType.ToString().ToUpper() == "ELEMENT")
                        {
                            A7NET.Comum.Comum.AppendNode(ref XmlAux, "Documento/Mensagem", XmlMensagemEntrada.ChildNodes[0].Name, XmlMensagemEntrada.ChildNodes[0].InnerXml);
                        }
                        else
                        {
                            for (int Index = 0; Index < XmlMensagemEntrada.ChildNodes.Count; Index++)
                            {
                                if (XmlMensagemEntrada.ChildNodes[Index].NodeType.ToString().ToUpper() == "ELEMENT")
                                {
                                    A7NET.Comum.Comum.AppendNode(ref XmlAux, "Documento/Mensagem", XmlMensagemEntrada.ChildNodes[Index].Name, XmlMensagemEntrada.ChildNodes[Index].InnerXml);
                                    break;
                                }
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region Mensagem String
                        //Se a Mensagem de entrada = Tipo String
                        XmlAux.SelectSingleNode("//Documento/Mensagem").InnerText = _XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_ENTR").InnerXml.Substring(20);
                        #endregion
                    }

                    A7NET.Comum.Comum.AppendNode(ref XmlAux, "Documento", "TX_REGR_TRNF_MESG", XmlRegra.DocumentElement.SelectSingleNode("TX_REGR_TRNF_MESG").InnerXml);

                    if (XmlAux.DocumentElement.SelectSingleNode("Formato/@FilaDestino") != null)
                    {
                        A7NET.Comum.Comum.AppendNode(ref _XmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", XmlAux.DocumentElement.SelectSingleNode("Formato/@FilaDestino").InnerText);
                    }
                    else
                    {
                        A7NET.Comum.Comum.AppendNode(ref _XmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", "");
                    }

                    //Carrega Informacoes para validacao XSD
                    XmlMensagemTraduzida = new XmlDocument();
                    SchemaSet = new XmlSchemaSet();
                    SchemaReader = new StringReader(XmlRegra.DocumentElement.SelectSingleNode("TX_VALID_SAID_MESG").InnerXml);
                    XmlReader XmlSchemaReader = XmlReader.Create(SchemaReader);
                    XmlSchemaReader.Read();
                    SchemaSet.Add("", XmlSchemaReader);
                    XmlMensagemTraduzida.Schemas = SchemaSet;

                    //Traduzir Mensagem
                    MensagemTraduzida = Traduzir(XmlAux);

                    #region Valida Xml Traduzido
                    XmlMensagemTraduzida.LoadXml(MensagemTraduzida);
                    EventHandler = new ValidationEventHandler(HandleValidationError);
                    XmlMensagemTraduzida.Validate(EventHandler);

                    // tratamento para aceitar TP_SOLI com 2 dígitos (=10), para não precisar alterar TP_SOLI de todos os layouts para Numérico(2), 
                    // pois haveria o risco de impactar o funcionamento de praticamente todos os layouts de Operação processados pelo A7
                    if (_XmlValido == false || _XmlParseError != null)
                    {
                        if (_XmlParseError.ToUpper().IndexOf("'TP_SOLI'") != -1 || _XmlParseError.IndexOf("'10'") != -1)
                        {
                            _XmlValido = true;
                            _XmlParseError = null;
                        }
                    }

                    //Xml Inválido
                    if (!_XmlValido)
                    {
                        CampoErro = _XmlParseError.Substring((_XmlParseError.IndexOf("'", 0)), (_XmlParseError.IndexOf("'", _XmlParseError.IndexOf("'", 0) + 1)) - (_XmlParseError.IndexOf("'", 0)) + 1);
                        DetalheOcorrenciaErro = string.Concat("O campo ", CampoErro, " está inválido ou não foi informado.");

                        _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerXml = DetalheOcorrenciaErro;

                        //Retorna a Mensagem Traduzida mas nao valida
                        _XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_SAID").InnerXml = XmlMensagemTraduzida.OuterXml;

                        //Monta Mensagem Erro
                        MensagemErro = MontaMensagem.MontaMensagemErro(_XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml, "A7Q.E.ENTRADA_NET", DetalheOcorrenciaErro);

                        //Put na fila de Erro
                        using (_MqConnector = new A7NET.ConfiguracaoMQ.MQConnector())
                        {
                            _MqConnector.MQConnect();
                            _MqConnector.MQQueueOpen("A7Q.E.ERRO", MQConnector.enumMQOpenOptions.PUT);
                            _MqConnector.Message = MensagemErro;
                            _MqConnector.MQPutMessage();
                            _MqConnector.MQQueueClose();
                            _MqConnector.MQEnd();
                        }

                        return false;

                    }
                    #endregion

                    //Seta Protocolo em uma unica string
                    Protocolo.Append(_ProtocoloMensagem.TipoMensagem);
                    Protocolo.Append(_ProtocoloMensagem.SiglaSistemaOrigem);
                    Protocolo.Append(_ProtocoloMensagem.SiglaSistemaDestino);
                    Protocolo.Append(_ProtocoloMensagem.CodigoEmpresa);

                    switch (int.Parse(_Regra["TP_FORM_MESG_SAID"].ToString()))
                    {
                        #region Saida String
                        case (int)A7NET.Comum.Comum.EnumTipoSaidaMensagem.SaidaString:

                            if (_ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "DV"
                            ||  _ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "BG"
                            ||  _ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "PZ")
                            {
                                MensagemTraduzida = XmlMensagemTraduzida.SelectSingleNode("//SaidaSTR").InnerText;
                            }
                            else
                            {
                                MensagemTraduzida = string.Concat(Protocolo.ToString(), XmlMensagemTraduzida.SelectSingleNode("//SaidaSTR").InnerText);
                            }
                            break;
                        #endregion

                        #region Saida String + XML
                        case (int)A7NET.Comum.Comum.EnumTipoSaidaMensagem.SaidaStringXML:

                            MensagemTraduzida = XmlMensagemTraduzida.SelectSingleNode("//SaidaSTR").InnerText;
                            MensagemTraduzida = string.Concat(MensagemTraduzida, MontaSaidaXml(_Regra["NO_TITU_MESG"].ToString(), XmlMensagemTraduzida));
                            break;
                        #endregion

                        #region Saida XML
                        case (int)A7NET.Comum.Comum.EnumTipoSaidaMensagem.SaidaXML:

                            MensagemTraduzida = MontaSaidaXml(_Regra["NO_TITU_MESG"].ToString(), XmlMensagemTraduzida);

                            XmlMensagemSaida = new XmlDocument();
                            XmlMensagemSaida.LoadXml(MensagemTraduzida);

                            if (int.Parse(XmlMensagemSaida.SelectSingleNode("//TP_MESG").InnerText) == (int)A7NET.Comum.Comum.EnumTipoMensagemLQS.RetornoCompromissada)
                            {
                                if (XmlMensagemSaida.SelectSingleNode("//SG_SIST_ORIG").InnerText == "YS")
                                {
                                    ParametroGeral = MensagemDAO.SelecionaMensagemParametrizacao();

                                    XmlParametroGeral = new XmlDocument();
                                    XmlParametroGeral.LoadXml(ParametroGeral);

                                    if (XmlParametroGeral.SelectSingleNode("//ALTERAR_EMPRESA_VOLTA_SAC") != null)
                                    {
                                        if (XmlParametroGeral.SelectSingleNode("//ALTERAR_EMPRESA_VOLTA_SAC/EMPRESA/@OBRIG").InnerText == "S")
                                        {
                                            XmlMensagemSaida.SelectSingleNode("//CO_EMPR").InnerText = XmlParametroGeral.SelectSingleNode("//ALTERAR_EMPRESA_VOLTA_SAC/EMPRESA").InnerText;
                                        }
                                    }
                                }
                            }

                            MensagemTraduzida = XmlMensagemSaida.OuterXml;
                            break;
                        #endregion

                        default:
                            break;

                    }

                    //Retorna o Codigo de composicao atributo identificador da mensagem
                    _XmlMensagem.DocumentElement.SelectSingleNode("CO_CMPO_ATRB_IDEF_MESG").InnerText = XmlMensagemTraduzida.SelectSingleNode("//SaidaID").InnerText;

                    if (int.Parse(_Regra["TP_NATZ_MESG"].ToString()) == (int)A7NET.Comum.Comum.EnumNaturezaMensagem.MensagemEnvio)
                    {
                        //Retorna a Mensagem Traduzida
                        _XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_SAID").InnerXml = MensagemTraduzida;
                    }
                    else if (int.Parse(_Regra["TP_NATZ_MESG"].ToString()) == (int)A7NET.Comum.Comum.EnumNaturezaMensagem.MensagemConsulta)
                    {
                        //Retorna a Mensagem Consulta
                         _XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_SAID").InnerXml = MensagemDAO.SelecionaRespostaMensagemConsulta(XmlMensagemTraduzida.SelectSingleNode("//SaidaID").InnerText);
                    }
                }
                else
                {
                    if (int.Parse(_Regra["TP_NATZ_MESG"].ToString()) == (int)A7NET.Comum.Comum.EnumNaturezaMensagem.MensagemECO)
                    {
                        _XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_SAID").InnerXml = _XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml;
                    }
                    else if (int.Parse(_Regra["TP_NATZ_MESG"].ToString()) == (int)A7NET.Comum.Comum.EnumNaturezaMensagem.MensagemEnvio)
                    {
                        MensagemTraduzida = MontaMensagemSaidaSemTraducao();

                        if (int.TryParse(_ProtocoloMensagem.TipoMensagem, out OutN))
                        {
                            TipoMensagem = Convert.ToString(OutN);
                        }
                        else
                        {
                            TipoMensagem = _ProtocoloMensagem.TipoMensagem.Trim();
                        }

                        if (TipoMensagem == "1002")
                        {
                            HeaderNZ = MensagemTraduzida.Substring(0, MensagemTraduzida.IndexOf("<") - 1);

                            XmlMensagemSaida = new XmlDocument();
                            XmlMensagemSaida.LoadXml(MensagemTraduzida.Substring(MensagemTraduzida.IndexOf("<")));
                            TraduzBook(ref XmlMensagemSaida, A7NET.Comum.Comum.EnumTipoTraducaoBookSPB.BookAntigo_BookNovo);

                            MensagemTraduzida = string.Concat(HeaderNZ, XmlMensagemSaida.OuterXml);
                        }

                        _XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_SAID").InnerXml = MensagemTraduzida;

                        _XmlMensagem.DocumentElement.SelectSingleNode("CO_CMPO_ATRB_IDEF_MESG").InnerText = MontaMensagem.ObterCodigoMensagemNZ(_XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml, _ProtocoloMensagem.TipoMensagem);

                        if (XmlRegra.DocumentElement.SelectSingleNode("TX_REGR_TRNF_MESG").SelectSingleNode("Formato/@FilaDestino") != null)
                        {
                            A7NET.Comum.Comum.AppendNode(ref _XmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", XmlRegra.DocumentElement.SelectSingleNode("TX_REGR_TRNF_MESG").SelectSingleNode("Formato/@FilaDestino").InnerText);
                        }
                        else
                        {
                            A7NET.Comum.Comum.AppendNode(ref _XmlMensagem, "Grupo_Mensagem", "NO_FILA_MQSE_DEST", "");
                        }
                    }
                }
                
                return true;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< Traduzir >>>
        protected string Traduzir(XmlDocument xmlMensagem)
        {
            A7NET.Comum.Comum.EnumTipoParteSaida IdTipoSaida = A7NET.Comum.Comum.EnumTipoParteSaida.SemParte;
            ConfiguraMensagem MontaMensagem = new ConfiguraMensagem();
            XmlDocument XmlSaida;
            string TipoSaida = "";
            string TagFormato = "";
            bool TraduzTipo;

            try
            {
                //Inicializa a variavel auxiliar para utilizacao de repeticoes
                _UtilizaRepeticao = false;

                XmlSaida = new XmlDocument();
                XmlSaida.LoadXml("<Saida></Saida>");

                if (xmlMensagem.SelectSingleNode("//Documento/Mensagem/@Tipo").InnerText == "XML")
                {
                    MontaMensagem.PrepararXML(ref xmlMensagem);
                }

                for (int Index = 0; Index <= 2; Index++)
                {
                    TraduzTipo = false;

                    switch (Index)
                    {
                        case 0:
                            if (xmlMensagem.SelectNodes("//Documento/*/Formato/IDOutPut").Count > 0)
                            {
                                IdTipoSaida = A7NET.Comum.Comum.EnumTipoParteSaida.ParteId;
                                TipoSaida = "SaidaID";
                                TagFormato = "IDOutPut";
                                TraduzTipo = true;
                            }
                            break;

                        case 1:
                            if (xmlMensagem.SelectNodes("//Documento/*/Formato/STROutPut").Count > 0)
                            {
                                IdTipoSaida = A7NET.Comum.Comum.EnumTipoParteSaida.ParteSTR;
                                TipoSaida = "SaidaSTR";
                                TagFormato = "STROutPut";
                                TraduzTipo = true;
                            }
                            break;

                        case 2:
                            if (xmlMensagem.SelectNodes("//Documento/*/Formato/XMLOutPut").Count > 0)
                            {
                                IdTipoSaida = A7NET.Comum.Comum.EnumTipoParteSaida.ParteXML;
                                TipoSaida = "SaidaXML";
                                TagFormato = "XMLOutPut";
                                TraduzTipo = true;
                            }
                            break;

                        default:
                            TraduzTipo = false;
                            break;
                    }

                    if (TraduzTipo)
                    {
                        A7NET.Comum.Comum.AppendNode(ref XmlSaida, "Saida", TipoSaida, "");
                        Converter(ref XmlSaida, xmlMensagem.SelectSingleNode("//Documento/Mensagem"), 
                                  xmlMensagem.SelectSingleNode("//" + TagFormato), IdTipoSaida, TipoSaida, 0);
                    }

                }

                return XmlSaida.InnerXml;

            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        #endregion

        #region <<< Converter >>>
        protected void Converter(ref XmlDocument xmlSaida, XmlNode nodeMensagem, XmlNode nodeFormato, 
                                 A7NET.Comum.Comum.EnumTipoParteSaida tipoParteSaida, string tipoSaida, int numeroIteracao)
        {
            XmlDocument XmlAux;
            XmlNode NodeContexto;
            XmlAttribute Attribute;
            int OutN;
            int TamanhoPai = 0;
            int TamanhoAvo = 0;
            int IteracaoAvo = 0;
            int Repeticoes = 0;
            string XPath;
            string Valor = "";
            string TargetTag;
            string TagPai;
            bool UsarNumeroIteracaoPai = false;
            bool LimparIndicesRepeticao = false;

            try
            {
                if (nodeMensagem == null) return;

                foreach (XmlNode Node in nodeFormato.SelectNodes("./*"))
                {
                    switch (nodeMensagem.SelectSingleNode("//Documento/Mensagem/@Tipo").InnerText)
                    {
                        #region XML
                        case "XML":
                            if (Node.SelectSingleNode("@Tipo").InnerText != "Grupo"
                            && Node.SelectSingleNode("@TargetTag").InnerText != "")
                            {
                                if (Node.ParentNode.Name.IndexOf("OutPut") == -1)
                                {
                                    XPath = ".//" + Node.SelectSingleNode("@TargetTag").InnerText;

                                    //Verificacao para nao impactar nas mensagens existentes
                                    if (Node.SelectSingleNode("@RepetTag") != null)
                                    {
                                        if (Node.SelectSingleNode("@RepetTag").InnerText != "0")
                                        {
                                            XPath = XPath + "[@Posicao>'" + Node.SelectSingleNode("@UltimaPosicao").InnerText + "']";
                                        }
                                    }

                                    //Esta verificacao e usada para itens que tem pai (possiveis repeticoes)
                                    //Verifica se existe algum no apos a ultima posicao pesquisada
                                    if (nodeMensagem.SelectSingleNode(XPath) != null)
                                    {
                                        Valor = nodeMensagem.SelectSingleNode(XPath).InnerText;

                                        //Verificacao para nao impactar nas mensagens existentes
                                        if (Node.SelectSingleNode("@UltimaPosicao") != null)
                                        {
                                            //Guarda a ultima posicao
                                            Node.SelectSingleNode("@UltimaPosicao").InnerText = nodeMensagem.SelectSingleNode(XPath + "/@Posicao").InnerText;
                                        }
                                    }
                                    else
                                    {
                                        //Nao existem mais nos
                                        Valor = "";
                                    }
                                }
                                else
                                {
                                    //Primeiro Nivel do XML
                                    if (nodeMensagem.SelectSingleNode(".//" + Node.SelectSingleNode("@TargetTag").InnerText) != null)
                                    {
                                        Valor = nodeMensagem.SelectSingleNode(".//" + Node.SelectSingleNode("@TargetTag").InnerText).InnerText;
                                    }
                                    else
                                    {
                                        Valor = "";
                                    }
                                }
                            }
                            else
                            {
                                Valor = "";
                            }

                            if (Node.Name.IndexOf("CO_FORM_LIQU") == -1
                            &&  Node.Name.IndexOf("CO_BANC") == -1
                            &&  Node.Name.IndexOf("CO_AGEN") == -1
                            &&  Node.Name.IndexOf("NU_CC") == -1
                            &&  Node.Name.IndexOf("VA_") == -1
                            &&  Node.Name.IndexOf("PU_") == -1
                            &&  Node.Name.IndexOf("PE_") == -1
                            &&  Node.Name.IndexOf("QT_") == -1)
                            {
                                if (Node.SelectSingleNode("@Default") != null)
                                {
                                    if (Node.SelectSingleNode("@DefaultObrigatorio").InnerText == "1")
                                    {
                                        Valor = Node.SelectSingleNode("@Default").InnerText;
                                    }
                                    else
                                    {
                                        if (Node.SelectSingleNode("@Default").InnerText.Trim() != "")
                                        {
                                            if (Valor.Trim() == "")
                                            {
                                                Valor = Node.SelectSingleNode("@Default").InnerText;
                                            }
                                            else
                                            {
                                                if (int.TryParse(Valor, out OutN))
                                                {
                                                    if (OutN == 0)
                                                    {
                                                        Valor = Node.SelectSingleNode("@Default").InnerText;
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                            }
                            break;
                        #endregion

                        #region String
                        case "String":
                            if (Node.SelectSingleNode("./@Tipo").InnerText == "Grupo")
                            {
                                //Tratamento para Item de grupo
                                Valor = "";
                            }
                            else
                            {
                                if (Node.ParentNode.SelectSingleNode("./@TamanhoOriginal") == null)
                                {
                                    TamanhoPai = 0;
                                }
                                else
                                {
                                    TamanhoPai = int.Parse(Node.ParentNode.SelectSingleNode("./@TamanhoOriginal").InnerText);
                                }

                                if (Node.ParentNode.ParentNode.ParentNode.SelectSingleNode("./@TamanhoOriginal") == null)
                                {
                                    TamanhoAvo = 0;
                                }
                                else
                                {
                                    TamanhoPai = int.Parse(Node.ParentNode.ParentNode.ParentNode.SelectSingleNode("./@TamanhoOriginal").InnerText);
                                    if (Node.ParentNode.ParentNode.ParentNode.ParentNode.SelectSingleNode("./@Iteracao") == null)
                                    {
                                        IteracaoAvo = 0;
                                    }
                                    else
                                    {
                                        IteracaoAvo = (int.Parse(Node.ParentNode.ParentNode.ParentNode.ParentNode.SelectSingleNode("./@Iteracao").InnerText) - 1);
                                    }
                                }

                                if (Node.SelectSingleNode("@Inicio").InnerText.Trim() != ""
                                && Node.SelectSingleNode("@TamanhoOriginal").InnerText.Trim() != "")
                                {
                                    Valor = nodeMensagem.InnerText.Substring(int.Parse(Node.SelectSingleNode("@Inicio").InnerText + (numeroIteracao * TamanhoPai) + (IteracaoAvo * TamanhoAvo)),
                                                                             int.Parse(Node.SelectSingleNode("@TamanhoOriginal").InnerText));
                                }
                                else
                                {
                                    Valor = "";
                                }
                            }

                            if (Node.Name.IndexOf("CO_FORM_LIQU") == -1
                            &&  Node.Name.IndexOf("CO_BANC") == -1
                            &&  Node.Name.IndexOf("CO_AGEN") == -1
                            &&  Node.Name.IndexOf("NU_CC") == -1
                            &&  Node.Name.IndexOf("VA_") == -1
                            &&  Node.Name.IndexOf("PU_") == -1
                            &&  Node.Name.IndexOf("PE_") == -1
                            &&  Node.Name.IndexOf("QT_") == -1)
                            {
                                if (Node.SelectSingleNode("@Default") != null)
                                {
                                    

                                    if (Node.SelectSingleNode("@DefaultObrigatorio").InnerText == "1")
                                    {
                                        Valor = Node.SelectSingleNode("@Default").InnerText;
                                    }
                                    else
                                    {
                                        if (Node.SelectSingleNode("@Default").InnerText.Trim() != "")
                                        {
                                            if (Valor.Trim() == "")
                                            {
                                                Valor = Node.SelectSingleNode("@Default").InnerText;
                                            }
                                            else
                                            {
                                                if (int.TryParse(Valor, out OutN))
                                                {
                                                    if (OutN == 0)
                                                    {
                                                        Valor = Node.SelectSingleNode("@Default").InnerText;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                        #endregion

                        default:
                            break;

                    }

                    if (tipoParteSaida == A7NET.Comum.Comum.EnumTipoParteSaida.ParteId
                    ||  tipoParteSaida == A7NET.Comum.Comum.EnumTipoParteSaida.ParteSTR)
                    {
                        if (Node.SelectSingleNode("@Tipo").InnerText != "Grupo")
                        {
                            //Tratamento para item de grupo
                            //Formatar como CopyFixo
                            Valor = A7NET.Comum.Comum.ValorToSTR(Valor, Node.SelectSingleNode("@Tipo").InnerText, int.Parse("0" + Node.SelectSingleNode("@Tamanho").InnerText),
                                                                 int.Parse("0" + Node.SelectSingleNode("@Decimais").InnerText), Node.SelectSingleNode("@Obrigatorio").InnerText, Node.Name);

                        }
                    }
                    else if (tipoParteSaida == A7NET.Comum.Comum.EnumTipoParteSaida.ParteXML)
                    {

                        if (nodeMensagem.SelectSingleNode("//Documento/Mensagem/@Tipo").InnerText == "String")
                        {
                            Valor = A7NET.Comum.Comum.ValorToXML(Valor, Node.SelectSingleNode("@Tipo").InnerText, int.Parse("0" + Node.SelectSingleNode("@Tamanho").InnerText),
                                                                 int.Parse("0" + Node.SelectSingleNode("@Decimais").InnerText), Node.SelectSingleNode("@Obrigatorio").InnerText);
                        }

                        if (Node.SelectSingleNode("@Tipo").InnerText.ToUpper().Equals("NUMBER")
                        &&  int.Parse("0" + Node.SelectSingleNode("@Tamanho").InnerText) <= 15)
                        {
                            if (Node.SelectSingleNode("@TargetTag") != null)
                            {
                                TargetTag = "|CO_PARP_CAMR|CO_CNPT_CAMR|CO_CLIE_CAMR|CO_BANC_LIQU_CAMR|CO_ANUE_CAMR|CO_CEDE_CAMR|CO_ADQU_CAMR";
                                if (TargetTag.IndexOf(string.Concat("|", Node.SelectSingleNode("@TargetTag").InnerText, "|")) == -1)
                                {
                                    if (Node.SelectSingleNode("@Decimais").InnerText == "0")
                                    {
                                        if (int.TryParse(Valor, out OutN))
                                        {
                                            Valor = Convert.ToString(OutN);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (Node.SelectSingleNode("@Decimais").InnerText == "0")
                                {
                                    if (int.TryParse(Valor, out OutN))
                                    {
                                        Valor = Convert.ToString(OutN);
                                    }
                                }
                            }
                        }
                    }

                    //Montar mensagem de saida
                    if (Node.SelectSingleNode("child::*") == null)
                    {
                        A7NET.Comum.Comum.AppendNode(ref xmlSaida, tipoSaida, Node.Name, Valor);
                    }
                    else
                    {
                        //Tag com Filho
                        A7NET.Comum.Comum.AppendNode(ref xmlSaida, tipoSaida, Node.Name, "");

                        //Inclui atributo para controle de iteracao, utilizado somente para string
                        Attribute = Node.OwnerDocument.CreateAttribute("Iteracao");
                        Node.Attributes.SetNamedItem(Attribute);

                        UsarNumeroIteracaoPai = false;

                        if (int.Parse("0" + Node.SelectSingleNode("@Repeticoes").InnerText) == 0)
                        {
                            Repeticoes = 1;
                            UsarNumeroIteracaoPai = true;

                            //Deve-se limpar os indices pois a nao repeticao e garantida pelo contexto do pai
                            LimparIndicesRepeticao = true;
                        }
                        else
                        {
                            Repeticoes = int.Parse(Node.SelectSingleNode("@Repeticoes").InnerText);
                            _UtilizaRepeticao = true;
                            //Zera Indices da Repeticao 
                            foreach (XmlNode ZeraNode in Node.SelectNodes(".//*"))
                            {
                                ZeraNode.SelectSingleNode("@UltimaPosicao").InnerText = "0";
                            }
                            LimparIndicesRepeticao = false;
                        }

                        //Trecho inserido pois o sistema não está identificando a quantidade correta de repetições.
                        //Ocorreu em producao com a CTP9015 enviada pelo A8 (Lancamento de PU).
                        if (Repeticoes == 1)
                        {
                            if (Node.SelectSingleNode("@TargetTag") != null)
                            {
                                if (Node.SelectSingleNode("@TargetTag").InnerText != "")
                                {
                                    if (nodeMensagem.SelectNodes("//" + Node.SelectSingleNode("@TargetTag").InnerText + "/*").Count > Repeticoes
                                    && Node.Name.IndexOf("REPET_") != 0
                                    && Node.Name.IndexOf("GRUPO_") == 0)
                                    {
                                        Repeticoes = nodeMensagem.SelectNodes("//" + Node.SelectSingleNode("@TargetTag").InnerText + "/*").Count;
                                        _UtilizaRepeticao = true;
                                        //Zera Indices da Repeticao 
                                        foreach (XmlNode ZeraNode in Node.SelectNodes(".//*"))
                                        {
                                            ZeraNode.SelectSingleNode("@UltimaPosicao").InnerText = "0";
                                        }
                                        LimparIndicesRepeticao = false;
                                    }
                                }
                            }
                        }

                        //Verificar se o pai ja tem contexto marcado
                        XPath = ".//";

                        //Armazena o contexto a ser utilizado para a busca das informacoes
                        NodeContexto = nodeMensagem;

                        //Armazenar o indice da repeticao para uso em caso de repeticao dentro da repeticao
                        //Importante: caso não haja tag pai de repeticao configurada o tradutor ignora o contexto e procura a partir da raiz
                        //caso haja tag pai configurada mas ela nao seja encontrada o tradutor ignora o contexto e procura a partir da raiz
                        //e finalmente, se a tag pai estiver configurada e presente na mensagem de entrada o contexto e utilizado
                        if (Node.SelectSingleNode("@UltimaPosicao") != null)
                        {
                            if (int.Parse("0" + Node.SelectSingleNode("@UltimaPosicao").InnerText) == 0)
                            {
                                if (Node.SelectSingleNode("@TargetTag").InnerText.Trim() != "")
                                {
                                    if (nodeMensagem.SelectSingleNode(XPath + Node.SelectSingleNode("@TargetTag").InnerText) != null)
                                    {
                                        Node.SelectSingleNode("@UltimaPosicao").InnerText = nodeMensagem.SelectSingleNode(XPath + Node.SelectSingleNode("@TargetTag").InnerText + "/@Posicao").InnerText;
                                        NodeContexto = null;
                                        NodeContexto = nodeMensagem.SelectSingleNode(XPath + Node.SelectSingleNode("@TargetTag").InnerText);
                                    }
                                    else
                                    {
                                        //Caso nao exista configuracao de TargetTag para tags de repeticao ignorar o fechamento do contexto
                                        //Assim a procura nao fica restrita a uma tag pai
                                        Node.SelectSingleNode("@UltimaPosicao").InnerText = "0";

                                        //Nao deve-se limpar os indices pois a tag pai nao existe e deve ser controlada pelo indice na busca
                                        LimparIndicesRepeticao = false;
                                    }
                                }
                                else
                                {
                                    //Caso nao exista configuracao de TargetTag para tags de repeticao ignorar o fechamento do contexto
                                    //Assim a procura nao fica restrita a uma tag pai
                                    Node.SelectSingleNode("@UltimaPosicao").InnerText = "0";

                                    //Nao deve-se limpar os indices pois a tag pai nao existe e deve ser controlada pelo indice na busca
                                    LimparIndicesRepeticao = false;
                                }
                            }
                            else
                            {
                                if (nodeMensagem.SelectSingleNode(XPath + Node.SelectSingleNode("@TargetTag").InnerText + "[@Posicao>'" + Node.SelectSingleNode("@UltimaPosicao").InnerText + "']") != null)
                                {
                                    Node.SelectSingleNode("@UltimaPosicao").InnerText = nodeMensagem.SelectSingleNode(XPath + Node.SelectSingleNode("@TargetTag").InnerText + "[@Posicao>'" + Node.SelectSingleNode("@UltimaPosicao").InnerText + "']" + "/@Posicao").InnerText;
                                    NodeContexto = null;
                                    NodeContexto = nodeMensagem.SelectSingleNode(XPath + Node.SelectSingleNode("@TargetTag").InnerText + "[@Posicao>'" + Convert.ToString(int.Parse(Node.SelectSingleNode("@UltimaPosicao").InnerText) - 1) + "']");
                                }
                                else
                                {
                                    //Ultrapassa o limite porque chegou ao fim das repeticoes
                                    Node.SelectSingleNode("@UltimaPosicao").InnerText = "99999999";
                                    NodeContexto = null;
                                }
                            }
                        }

                        TagPai = Node.Name;

                        for (int Index = 1; Index <= Repeticoes; Index++)
                        {
                            //Armazena o numero da iteracao para uso no deslocamento string (repeticao string)
                            Node.SelectSingleNode("@Iteracao").InnerText = Convert.ToString(Index);

                            //Usar Index - 1 para pegar a primeira posicao
                            XmlAux = new XmlDocument();
                            XmlAux.InnerXml = xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + Node.Name)[xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + Node.Name).Count - 1].OuterXml;

                            Converter(ref XmlAux, NodeContexto, Node, tipoParteSaida, Node.Name,
                                      (UsarNumeroIteracaoPai ? numeroIteracao : Index - 1));

                            xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + Node.Name)[xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + Node.Name).Count - 1].InnerXml = XmlAux.SelectNodes(Node.Name)[0].InnerXml;
                        }

                        if (LimparIndicesRepeticao)
                        {
                            //Limpar os indices armazenados para o contexto pois na proxima iteracao o indice do pai ira mudar
                            foreach (XmlNode ZeraNode in Node.SelectNodes(".//*"))
                            {
                                ZeraNode.SelectSingleNode("@UltimaPosicao").InnerText = "0";
                            }
                        }

                        if (tipoParteSaida == A7NET.Comum.Comum.EnumTipoParteSaida.ParteXML)
                        {
                            // Caso seja XML, remove o Item de Agrupamento se as tags estiverem vazias, string e id sao posicionais
                            if (xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + TagPai)[xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + TagPai).Count - 1].InnerText == "")
                            {
                                //Remove Tag
                                xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + TagPai)[xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + TagPai).Count - 1].ParentNode.RemoveChild(xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + TagPai)[xmlSaida.SelectSingleNode("//" + tipoSaida).SelectNodes("//" + TagPai).Count - 1]);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }
        #endregion

        #region <<< MontaSaidaXml >>>
        protected string MontaSaidaXml(string nomeTagPrincipal, XmlDocument xmlMensagemTraduzida)
        {
            XmlDocument XmlMensagemSaida;
            string MensagemSaida;
            int OutN;

            try
            {
                if (nomeTagPrincipal.Trim() == "")
                {
                    nomeTagPrincipal = "MESG";
                }
                else
                {
                    nomeTagPrincipal = nomeTagPrincipal.Trim();
                }

                XmlMensagemSaida = new XmlDocument();

                //Se TipoMensagem = Numerico
                if (int.TryParse(_ProtocoloMensagem.TipoMensagem, out OutN))
                {
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, "", nomeTagPrincipal, "");
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, "TP_MESG", _ProtocoloMensagem.TipoMensagem);
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, "SG_SIST_ORIG", _ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper());
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, "SG_SIST_DEST", _ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper());
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, "CO_EMPR", Convert.ToString(int.Parse(_ProtocoloMensagem.CodigoEmpresa)));

                    foreach (XmlNode Node in xmlMensagemTraduzida.SelectNodes("//SaidaXML/*"))
                    {
                        A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, Node.Name, Node.InnerXml);
                    }

                    if (_ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "NZ")
                    {
                        TraduzBook(ref XmlMensagemSaida, A7NET.Comum.Comum.EnumTipoTraducaoBookSPB.BookNovo_BookAntigo);
                    }

                    MensagemSaida = XmlMensagemSaida.OuterXml;

                }
                else
                {
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, "", "SISMSG", "");
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, "SISMSG", nomeTagPrincipal, "");

                    foreach (XmlNode Node in xmlMensagemTraduzida.SelectNodes("//SaidaXML/*"))
                    {
                        if (Node.Name.ToUpper().Substring(0, 2) == "DT"
                        || Node.Name.ToUpper().Substring(0, 2) == "DH")
                        {
                            if (Node.InnerText != "0")
                            {
                                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, Node.Name, Node.InnerXml);
                            }
                            else
                            {
                                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, Node.Name, "");
                            }
                        }
                        else
                        {
                            A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, nomeTagPrincipal, Node.Name, Node.InnerXml);
                        }
                    }

                    if (_ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "NZ")
                    {
                        TraduzBook(ref XmlMensagemSaida, A7NET.Comum.Comum.EnumTipoTraducaoBookSPB.BookAntigo_BookNovo);
                    }

                    MensagemSaida = XmlMensagemSaida.OuterXml;

                }

                return MensagemSaida;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< MontaMensagemSaidaSemTraducao >>>
        protected string MontaMensagemSaidaSemTraducao()
        {
            XmlDocument XmlMensagemSaidaSemTraducao;

            try
            {
                if (_ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "A8")
                {
                    XmlMensagemSaidaSemTraducao = new XmlDocument();
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaidaSemTraducao, "", "MESG", "");
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaidaSemTraducao, "MESG", "TP_MESG", _ProtocoloMensagem.TipoMensagem);
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaidaSemTraducao, "MESG", "SG_SIST_ORIG", _ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper());
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaidaSemTraducao, "MESG", "SG_SIST_DEST", _ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper());
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaidaSemTraducao, "MESG", "CO_EMPR", Convert.ToString(int.Parse(_ProtocoloMensagem.CodigoEmpresa)));
                    A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaidaSemTraducao, "MESG", "TX_MESG", _XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml);

                    if (_ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "NZ")
                    {
                        TraduzBook(ref XmlMensagemSaidaSemTraducao, A7NET.Comum.Comum.EnumTipoTraducaoBookSPB.BookNovo_BookAntigo);
                    }

                    return XmlMensagemSaidaSemTraducao.OuterXml;
                }
                else
                {
                    return _XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml.Substring(20);
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< TraduzBook >>>
        //-------------------------------------------------------------------------------------
        //Descricao das alteracoes do BOOK SPB
        //-------------------------------------------------------------------------------------
        //Tipo                               BookAntigo               BookNovo
        //-------------------------------------------------------------------------------------
        //Numerico com decimais              '0,345'                  '0.345'
        //Data                               '20051231'               '2005-12-31'
        //AnoMes                             '200512'                 '2005-12'
        //Hora                               '185959'                 '18:59:59'
        //Data Hora                          '20051231185959'         '2005-12-31T18:59:59'
        //-------------------------------------------------------------------------------------
        protected void TraduzBook(ref XmlDocument xmlMensagem, A7NET.Comum.Comum.EnumTipoTraducaoBookSPB tipoTraducao)
        {
            string CodigoMensagem;

            try
            {
                if (xmlMensagem.SelectSingleNode("//CodMsg") == null)
                {
                    return;
                }

                if (xmlMensagem.SelectSingleNode("//CodMsg").InnerText.Trim().ToUpper().Substring(7) == "E")
                {
                    CodigoMensagem = xmlMensagem.SelectSingleNode("//CodMsg").InnerText.Trim().Substring(0, 7);
                }
                else
                {
                    CodigoMensagem = xmlMensagem.SelectSingleNode("//CodMsg").InnerText;
                }

                foreach (DataRow Row in _DataSetCache.TB_MENSAGEM_SPB.Select("CO_MESG='" + CodigoMensagem + "'"))
                {
                    //Para cada linha da mensagem
                    foreach (XmlNode Node in xmlMensagem.SelectNodes("//" + Row["NO_TAG"].ToString()))
                    {
                        if (tipoTraducao == A7NET.Comum.Comum.EnumTipoTraducaoBookSPB.BookNovo_BookAntigo)
                        {
                            //Traducao do formato novo para o formato antigo
                            if (Node.InnerText == "")
                            {
                                //Se campo vazio, remove-o do XML
                                if (Node.Attributes.Count == 0)
                                {
                                    Node.ParentNode.RemoveChild(Node);
                                }
                            }
                            else
                            {
                                if (int.Parse(Row["QT_CASA_DECI"].ToString()) > 0)
                                {
                                    //Campo numerico com casas decimais
                                    Node.InnerText = Node.InnerText.Replace(".", ",");
                                }
                                else
                                {
                                    //Campo tipo data ou hora
                                    Node.InnerText = Node.InnerText.Replace("T", "").Replace("-", "").Replace(":", "");
                                }
                            }
                        }
                        else if (tipoTraducao == A7NET.Comum.Comum.EnumTipoTraducaoBookSPB.BookAntigo_BookNovo)
                        {
                            //Traducao do formato antigo para o formato novo
                            if (Node.InnerText == "")
                            {
                                //Se campo vazio, remove-o do XML
                                if (Node.Attributes.Count == 0)
                                {
                                    Node.ParentNode.RemoveChild(Node);
                                }
                            }
                            else
                            {
                                if (int.Parse(Row["QT_CASA_DECI"].ToString()) > 0)
                                {
                                    //Campo numerico com casas decimais
                                    Node.InnerText = Node.InnerText.Replace(",", ".");
                                }
                                else
                                {
                                    switch (Row["NO_TIPO_TAG"].ToString())
                                    {
                                        case "AnoMes":
                                            if (Node.InnerText.Length == 6)
                                            {
                                                Node.InnerText = string.Format("{0:yyyy-MM}", DateTime.ParseExact(Node.InnerText, "yyyyMM", System.Threading.Thread.CurrentThread.CurrentCulture));
                                            }
                                            break;

                                        case "Data":
                                            if (Node.InnerText.Length == 8)
                                            {
                                                Node.InnerText = string.Format("{0:yyyy-MM-dd}", DateTime.ParseExact(Node.InnerText, "yyyyMMdd", System.Threading.Thread.CurrentThread.CurrentCulture));
                                            }
                                            break;

                                        case "Data Hora":
                                            if (Node.InnerText.Length == 14)
                                            {
                                                Node.InnerText = string.Format("{0:yyyy-MM-ddTHH:mm:ss}", DateTime.ParseExact(Node.InnerText, "yyyyMMddHHmmss", System.Threading.Thread.CurrentThread.CurrentCulture));
                                            }
                                            break;

                                        case "Hora":
                                            if (Node.InnerText.Length == 6)
                                            {
                                                Node.InnerText = string.Format("{0:HH:mm:ss}", DateTime.ParseExact(Node.InnerText, "HHmmss", System.Threading.Thread.CurrentThread.CurrentCulture));
                                            }
                                            break;

                                        default:
                                            break;

                                    }
                                }
                            }
                        }
                    }
                }

                //Remove todas as TAGS sem conteudo
                foreach (XmlNode NodeDel in xmlMensagem.SelectNodes("//*[.='']"))
                {
                    if (NodeDel.Attributes.Count == 0)
                    {
                        NodeDel.ParentNode.RemoveChild(NodeDel);
                    }
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< IncluirAvisoMonitor >>>
        protected void IncluirAvisoMonitor(string codigoEmpresa, string controleRemessa, string codigoMensagem,
                                           A7NET.Comum.Comum.EnumFluxoMonitor tipoFluxoMensagem,
                                           A7NET.Comum.Comum.EnumStatusMonitor situacaoMensagem,
                                           string sistemaOrigem, string sistemaDestino, string nomeFilaDestino,
                                           DateTime? dataMovimento, string valor, string situacaoLancamento,
                                           string codigoISPB_ErroPKI, DateTime? dataHora_Ext_R1)
        {
            StringBuilder Carimbo = new StringBuilder();
            string NomeFilaMonitor = "NZQ.E.MONITOR";

            try
            {
                #region Campo CO_EMPR
                Carimbo.Append(codigoEmpresa.Trim().PadLeft(5, '0'));
                #endregion

                #region Campo NU_CTRL_REME
                if (controleRemessa.Length > 20)
                {
                    Carimbo.Append(controleRemessa.Trim().Substring(0, 20));
                }
                else
                {
                    Carimbo.Append(controleRemessa.Trim().PadRight(20, ' '));
                }
                #endregion

                #region Campo CO_MESG
                if (codigoMensagem.Length > 9)
                {
                    Carimbo.Append(codigoMensagem.Trim().Substring(0, 9));
                }
                else
                {
                    Carimbo.Append(codigoMensagem.Trim().PadRight(9, ' '));
                }
                #endregion

                #region Campo TP_FLUX_MESG
                switch (tipoFluxoMensagem)
                {
                    case A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_NORMAL:
                        Carimbo.Append("N");
                        break;

                    case A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_LEGADINHO:
                        Carimbo.Append("L");
                        break;

                    case A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_CONT_MAXIMA:
                        Carimbo.Append("C");
                        break;

                    case A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_MSG_EXTERNA:
                        Carimbo.Append("E");
                        break;

                    case A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_MSG_CARGA:
                        Carimbo.Append("K");
                        break;

                    default:
                        break;

                }
                #endregion

                #region Campo SITU_MESG
                Carimbo.Append(Convert.ToString((int)situacaoMensagem).PadLeft(3, '0'));
                #endregion

                #region Campo ORIG_MESG
                if (sistemaOrigem.Length > 3)
                {
                    Carimbo.Append(sistemaOrigem.Trim().Substring(0, 3));
                }
                else
                {
                    Carimbo.Append(sistemaOrigem.Trim().PadRight(3, ' '));
                }
                #endregion

                #region Campo DEST_MESG
                if (sistemaDestino.Length > 3)
                {
                    Carimbo.Append(sistemaDestino.Trim().Substring(0, 3));
                }
                else
                {
                    Carimbo.Append(sistemaDestino.Trim().PadRight(3, ' '));
                }
                #endregion

                #region Campo NO_FILA_DEST
                if (nomeFilaDestino.Length > 48)
                {
                    Carimbo.Append(nomeFilaDestino.Trim().Substring(0, 48));
                }
                else
                {
                    Carimbo.Append(nomeFilaDestino.Trim().PadRight(48, ' '));
                }
                #endregion

                #region Campos ID_SINA_VALO / TAG_VALOR
                if (valor == "")
                {
                    Carimbo.Append("0"); //ID_SINA_VALO
                    Carimbo.Append(valor.Trim().PadLeft(18, '0')); //TAG_VALOR
                }
                else
                {
                    //ID_SINA_VALO
                    if (Convert.ToDecimal(valor) < 0)
                    {
                        Carimbo.Append("-");
                    }
                    else
                    {
                        Carimbo.Append("0");
                    }

                    //TAG_VALOR
                    if (valor.Length > 18)
                    {
                        Carimbo.Append(string.Format("{0:N2}", Convert.ToDecimal(valor)).Substring(0, 18).Replace(",", "").Replace(".", ""));
                    }
                    else
                    {
                        Carimbo.Append(string.Format("{0:N2}", Convert.ToDecimal(valor)).Replace(",", "").Replace(".", "").PadLeft(18, '0'));
                    }
                }
                #endregion

                #region Campo TAG_SITLANC
                if (situacaoLancamento.Length > 20)
                {
                    Carimbo.Append(situacaoLancamento.Trim().Substring(0, 20));
                }
                else
                {
                    Carimbo.Append(situacaoLancamento.Trim().PadRight(20, ' '));
                }
                #endregion

                #region Campo CO_ISPB_PKI
                if (codigoISPB_ErroPKI.Length > 8)
                {
                    Carimbo.Append(codigoISPB_ErroPKI.Trim().Substring(0, 8));
                }
                else
                {
                    Carimbo.Append(codigoISPB_ErroPKI.Trim().PadRight(8, ' '));
                }
                #endregion

                #region Campo CO_ISPB_PKI
                if (dataHora_Ext_R1 == null)
                {
                    Carimbo.Append("".PadRight(14, '0'));
                }
                else
                {
                    Carimbo.Append(string.Format("{0:yyyyMMddHHmmss}", dataHora_Ext_R1));
                }
                #endregion

                #region Campo TAG_DTMOVTO
                if (dataMovimento == null)
                {
                    Carimbo.Append("".PadRight(8, '0'));
                }
                else
                {
                    Carimbo.Append(string.Format("{0:yyyyMMdd}", dataMovimento));
                }
                #endregion

                #region Campo FILLER
                Carimbo.Append("".PadRight(39, ' '));
                #endregion

                //Put na fila NZQ.E.MONITOR
                using (_MqConnector = new A7NET.ConfiguracaoMQ.MQConnector())
                {
                    _MqConnector.MQConnect();
                    _MqConnector.MQQueueOpen(NomeFilaMonitor, MQConnector.enumMQOpenOptions.PUT);
                    _MqConnector.Message = Carimbo.ToString();
                    _MqConnector.MQPutMessage();
                    _MqConnector.MQQueueClose();
                    _MqConnector.MQEnd();
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< PostarMensagemTraduzida >>>
        protected void PostarMensagemTraduzida()
        {
            XmlDocument MensagemSPB;
            string NomeFilaDestino;
            string NumCtrlIF;
            string CodigoMensagem;
            string Valor;
            int outN;

            try
            {
                if (_XmlMensagem.DocumentElement.SelectSingleNode("NO_FILA_MQSE_DEST").InnerText.Trim() == "")
                {
                    NomeFilaDestino = _DataSetCache.TB_ENDE_FILA_MQSE.Select("SG_SIST_DEST='" + _XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerText.Trim().ToUpper() + "'")[0]["NO_FILA_MQSE"].ToString();
                }
                else
                {
                    NomeFilaDestino = _XmlMensagem.DocumentElement.SelectSingleNode("NO_FILA_MQSE_DEST").InnerText.Trim();
                }

                if (_XmlMensagem.SelectSingleNode("//SG_SIST_DEST").InnerText == "NZ")
                {
                    //Enviar carimbo
                    //Obtem a mensagem de entrada pois esta contem as tags necessarias com nomes padronizados pelo A8
                    MensagemSPB = new XmlDocument();
                    MensagemSPB.LoadXml(_XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml.Substring(20));

                    //Obtem o Numero de Controle IF
                    if (MensagemSPB.SelectSingleNode("//NU_CTRL_IF") != null)
                    {
                        NumCtrlIF = MensagemSPB.SelectSingleNode("//NU_CTRL_IF").InnerText;
                    }
                    else
                    {
                        NumCtrlIF = "";
                    }

                    //Obtem o Codigo da Mensagem
                    if (MensagemSPB.SelectSingleNode("//CO_MESG") != null)
                    {
                        CodigoMensagem = MensagemSPB.SelectSingleNode("//CO_MESG").InnerText;
                    }
                    else if (!int.TryParse(_XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml.Substring(0, 9), out outN))
                    {
                        CodigoMensagem = _XmlMensagem.SelectSingleNode("//TX_CNTD_ENTR").InnerXml.Substring(0, 9);
                    }
                    else
                    {
                        CodigoMensagem = "";
                    }

                    //Obtem o Valor da Mensagem
                    if (MensagemSPB.SelectSingleNode("//VA_OPER_ATIV") != null)
                    {
                        Valor = MensagemSPB.SelectSingleNode("//VA_OPER_ATIV").InnerText;
                    }
                    else
                    {
                        Valor = "";
                    }

                    IncluirAvisoMonitor(_XmlMensagem.SelectSingleNode("//CO_EMPR_ORIG").InnerText, NumCtrlIF, CodigoMensagem,
                                        A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_NORMAL, A7NET.Comum.Comum.EnumStatusMonitor.MONITOR_ENVIADA_NZ,
                                        "A8", "NZ", NomeFilaDestino, DateTime.Today, Valor, "", "", null);

                }

                PostarMQSeries(NomeFilaDestino);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< PostarMQSeries >>>
        protected void PostarMQSeries(string nomeFilaDestino)
        {
            XmlDocument MensagemPostada;
            string MessageId;
            string MensagemTraduzida;

            try
            {
                MensagemPostada = new XmlDocument();

                try
                {
                    MensagemPostada.LoadXml(_XmlMensagem.SelectSingleNode("//TX_CNTD_SAID").InnerXml);
                    if (MensagemPostada.SelectSingleNode("//TP_SOLI") != null)
                    {
                        if (MensagemPostada.SelectSingleNode("//SG_SIST_DEST").InnerText == "A8")
                        {
                            nomeFilaDestino = "A8Q.E.ENTRADA_NET";
                        }
                    }
                    else
                    {
                        if (MensagemPostada.SelectSingleNode("//TP_MESG") != null)
                        {
                            switch (int.Parse(MensagemPostada.SelectSingleNode("//TP_MESG").InnerText))
                            {
                                case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemNZA8:
                                case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemErroNZA8:
                                case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZErro:
                                case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZR1:
                                case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZR2:
                                case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemSTR0010R2PZA8:
                                    nomeFilaDestino = "A8Q.E.ENTRADA_NET";
                                    break;

                                default:
                                    break;

                            }
                        }
                    }
                }
                catch
                {
                    if (MensagemPostada.SelectSingleNode("//TP_MESG") != null)
                    {
                        switch (int.Parse(MensagemPostada.SelectSingleNode("//TP_MESG").InnerText))
                        {
                            case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemNZA8:
                            case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemErroNZA8:
                            case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZErro:
                            case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZR1:
                            case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZR2:
                            case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemSTR0010R2PZA8:
                                nomeFilaDestino = "A8Q.E.ENTRADA_NET";
                                break;

                            default:
                                break;

                        }
                    }
                }

                MensagemTraduzida = _XmlMensagem.SelectSingleNode("//TX_CNTD_SAID").InnerXml;

                //Put na fila
                using (_MqConnector = new A7NET.ConfiguracaoMQ.MQConnector())
                {
                    _MqConnector.MQConnect();
                    _MqConnector.MQQueueOpen(nomeFilaDestino, MQConnector.enumMQOpenOptions.PUT);
                    _MqConnector.ReplyToQueueName = "A7Q.E.REPORT";
                    _MqConnector.Message = MensagemTraduzida;
                    _MqConnector.Prioridade = (int)MQConnector.PrioridadeMensagem.Maxima;
                    _MqConnector.MQPutMessage();
                    MessageId = _MqConnector.MessageIdHex;
                    _MqConnector.MQQueueClose();
                    _MqConnector.MQEnd();
                }

                _XmlMensagem.DocumentElement.SelectSingleNode("CO_MESG_MQSE").InnerText = MessageId;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< EnviarMensagemRejeicaoLegado >>>
        protected void EnviarMensagemRejeicaoLegado()
        {
            XmlDocument XmlMensagemSaida;
            string MensagemSaida;
            string SiglaSistemaOrigem;
            string TipoMensagem;
            string TipoMensagemRetorno;
            string CodigoEmpresaRetorno;
            string ProtocoloRetornoLegado;
            int OutN;

            try
            {
                XmlMensagemSaida = new XmlDocument();
                XmlMensagemSaida.LoadXml(_XmlMensagem.SelectSingleNode("//TX_CNTD_SAID//SaidaXML").OuterXml);

                SiglaSistemaOrigem = _XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerText.Trim().ToUpper();

                switch (SiglaSistemaOrigem)
                {
                    case "WZ":
                        if (XmlMensagemSaida.SelectSingleNode("//TP_SOLI") != null)
                        {
                            if (int.TryParse(XmlMensagemSaida.SelectSingleNode("//TP_SOLI").InnerText, out OutN))
                            {
                                if (OutN == (int)A7NET.Comum.Comum.EnumTipoSolicitacao.Inclusao)
                                {
                                    return;
                                }
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            return;
                        }
                        break;

                    case "A8": case "NZ": case "BG": case "DV":
                        return;

                    default:
                        break;

                }

                TipoMensagem = _XmlMensagem.DocumentElement.SelectSingleNode("TP_MESG").InnerText;
                TipoMensagemRetorno = _DataSetCache.TB_TIPO_OPER.Select("TP_MESG_RECB_INTE='" + (int.TryParse(TipoMensagem, out OutN) ? Convert.ToString(OutN) : TipoMensagem.Trim()) + "'")[0]["TP_MESG_RETN_INTE"].ToString();

                if (SiglaSistemaOrigem == "WZ")
                {
                    TipoMensagemRetorno = Convert.ToString(int.Parse(TipoMensagemRetorno) + 2000);
                }

                TipoMensagemRetorno = TipoMensagemRetorno.PadLeft(9, '0');
                CodigoEmpresaRetorno = _XmlMensagem.DocumentElement.SelectSingleNode("CO_EMPR_ORIG").InnerText.PadLeft(5, '0');

                // Protocolo Retorno Legado
                // TipoMensagem        //String * 9 = TipoMensagemRetorno
                // SiglaSistemaOrigem  //String * 3 = "A8 "
                // SiglaSistemaDestino //String * 3 = SiglaSistemaOrigem
                // CodigoEmpresa       //String * 5 = CodigoEmpresaRetorno
                ProtocoloRetornoLegado = string.Concat(TipoMensagemRetorno, "A8 ", SiglaSistemaOrigem.PadRight(3, ' '), CodigoEmpresaRetorno);

                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, XmlMensagemSaida.DocumentElement.Name, "TP_MESG", _XmlMensagem.DocumentElement.SelectSingleNode("//TP_MESG").InnerText);
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, XmlMensagemSaida.DocumentElement.Name, "TP_RETN", "2");
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, XmlMensagemSaida.DocumentElement.Name, "DtMovto", _XmlMensagem.DocumentElement.SelectSingleNode("//DT_MOVI").InnerText);
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, XmlMensagemSaida.DocumentElement.Name, "CO_ULTI_SITU_PROC", "99");
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, XmlMensagemSaida.DocumentElement.Name, "CO_ERRO1", "1017");
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemSaida, XmlMensagemSaida.DocumentElement.Name, "DE_ERRO1", _XmlMensagem.SelectSingleNode("//TX_DTLH_OCOR_ERRO").InnerText);

                MensagemSaida = string.Concat(ProtocoloRetornoLegado, XmlMensagemSaida.OuterXml);

                //Put na fila A7Q.E.ENTRADA
                using (_MqConnector = new A7NET.ConfiguracaoMQ.MQConnector())
                {
                    _MqConnector.MQConnect();
                    _MqConnector.MQQueueOpen("A7Q.E.ENTRADA", MQConnector.enumMQOpenOptions.PUT);
                    _MqConnector.Message = MensagemSaida;
                    _MqConnector.MQPutMessage();
                    _MqConnector.MQQueueClose();
                    _MqConnector.MQEnd();
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< SalvarMensagem >>>
        protected void SalvarMensagem(A7NET.Comum.Comum.EnumOcorrencia codigoOcorrencia)
        {
            A7NET.Data.MensagemDAO MensagemDAO = new A7NET.Data.MensagemDAO();
            A7NET.Data.MensagemDAO.EstruturaMensagem DadosMensagem = new A7NET.Data.MensagemDAO.EstruturaMensagem();
            string DetalheOcorrenciaErro;
            string TipoMensagem;
            int CodigoMensagem;
            int OutN;

            try
            {
                //Seta Parametros para Gravar no Banco de Dados
                DadosMensagem.SiglaSistema  = _XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerText;

                if (_XmlMensagem.SelectSingleNode("//Grupo_Mensagem/CO_MESG_MQSE") != null)
                {
                    DadosMensagem.MessageId = _XmlMensagem.SelectSingleNode("//Grupo_Mensagem/CO_MESG_MQSE").InnerText;
                }

                TipoMensagem = _XmlMensagem.DocumentElement.SelectSingleNode("TP_MESG").InnerText;
                DadosMensagem.TipoMensagem = int.TryParse(TipoMensagem, out OutN) ? Convert.ToString(OutN) : TipoMensagem.Trim();

                DadosMensagem.CodigoEmpresaOrigem = int.Parse(_XmlMensagem.DocumentElement.SelectSingleNode("CO_EMPR_ORIG").InnerText);

                DadosMensagem.DataInicioRegra = _DataHoraInicioRegra;
                
                if (_XmlMensagem.SelectSingleNode("//Grupo_Mensagem/CO_CMPO_ATRB_IDEF_MESG") != null)
                {
                    DadosMensagem.CodigoOperacaoAtiva = _XmlMensagem.SelectSingleNode("//Grupo_Mensagem/CO_CMPO_ATRB_IDEF_MESG").InnerText;
                }

                //Grava Mensagem Entrada Base64
                DadosMensagem.CodigoXmlEntrada = MensagemDAO.PersisteMensagemBase64(_XmlMensagem.SelectSingleNode("//Grupo_Mensagem/TX_CNTD_ENTR").InnerXml);

                //Grava Mensagem Saida Base64
                if (_XmlMensagem.SelectSingleNode("//Grupo_Mensagem/TX_CNTD_SAID") != null)
                {
                    if (_XmlMensagem.SelectSingleNode("//Grupo_Mensagem/TX_CNTD_SAID").InnerText != "")
                    {
                        DadosMensagem.CodigoXmlSaida = MensagemDAO.PersisteMensagemBase64(_XmlMensagem.SelectSingleNode("//Grupo_Mensagem/TX_CNTD_SAID").InnerXml);
                    }
                }
            
                DadosMensagem.TipoFormatoMensagemSaida = int.Parse(_XmlMensagem.DocumentElement.SelectSingleNode("TP_FORM_MESG_SAID").InnerText);

                //Grava Mensagem no Banco de Dados
                CodigoMensagem = MensagemDAO.PersisteMensagem(DadosMensagem);

                //Seta Codigo da Mensagem
                _XmlMensagem.DocumentElement.SelectSingleNode("CO_MESG").InnerText = Convert.ToString(CodigoMensagem);

                DetalheOcorrenciaErro = _XmlMensagem.DocumentElement.SelectSingleNode("TX_DTLH_OCOR_ERRO").InnerText;

                //Grava Situacao da Mensagem
                MensagemDAO.PersisteSituacaoMensagem(CodigoMensagem, DetalheOcorrenciaErro, codigoOcorrencia);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< SalvarMensagemRejeitada >>>
        protected void SalvarMensagemRejeitada(string mensagemRecebida, string nomeFila, string messageId)
        {
            A7NET.Data.MensagemDAO.EstruturaMensagemRejeitada DadosMensagem = new A7NET.Data.MensagemDAO.EstruturaMensagemRejeitada();
            A7NET.Data.MensagemDAO MensagemDAO = new A7NET.Data.MensagemDAO();
            ConfiguraMensagem MontaMensagem = new ConfiguraMensagem();
            string MensagemErro;

            try
            {
                DadosMensagem.MessageId = messageId;
                DadosMensagem.CodigoOcorrencia = (int)A7NET.Comum.Comum.EnumOcorrencia.RejeicaoNaoAutenticidade;

                //Grava Mensagem Erro Base64
                DadosMensagem.CodigoXml = MensagemDAO.PersisteMensagemBase64(mensagemRecebida);

                if (_ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() != "")
                {
                    DadosMensagem.SistemaOrigem = _ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper();
                }
                else
                {
                    DadosMensagem.SistemaOrigem = "ORIGEM NAO IDENTIFICADA";
                }

                DadosMensagem.DetalheOcorrencia = _XmlMensagem.SelectSingleNode("//TX_DTLH_OCOR_ERRO").InnerText;
                
                //Grava Mensagem Rejeitada
                MensagemDAO.PersisteMensagemRejeitada(DadosMensagem);

                //Monta Mensagem Erro
                MensagemErro = MontaMensagem.MontaMensagemErro(mensagemRecebida, nomeFila, DadosMensagem.DetalheOcorrencia);

                //Put na fila Erro
                using (_MqConnector = new A7NET.ConfiguracaoMQ.MQConnector())
                {
                    _MqConnector.MQConnect();
                    _MqConnector.MQQueueOpen("A7Q.E.ERRO", MQConnector.enumMQOpenOptions.PUT);
                    _MqConnector.Message = MensagemErro;
                    _MqConnector.MQPutMessage();
                    _MqConnector.MQQueueClose();
                    _MqConnector.MQEnd();
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< IncluiRepeticao >>>
        protected bool IncluiRepeticao(string nomeTagMensagem)
        {
            Repeticao ControleRepeticao = new Repeticao();
            XmlDocument XmlMesgSaida;
            XmlDocument XmlMesgAux;
            XmlDocument XmlRepeticao;
            StringBuilder Repeticao = new StringBuilder();
            string MesgAux;
            string MesgSaida;
            string CabecalhoSaida = "";
            string NomeTagPrincipal;

            try
            {
                //Pega o XML da Mensagem
                MesgAux = _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml;
                MesgSaida = MesgAux;

                //Faz o Load do XML
                XmlMesgSaida = new XmlDocument();

                try
                {
                    XmlMesgSaida.LoadXml(MesgSaida);
                }
                catch
                {
                    try
                    {
                        CabecalhoSaida = MesgAux.Substring(0, 20);
                        MesgSaida = MesgAux.Substring(20);
                        XmlMesgSaida.LoadXml(MesgSaida);
                    }
                    catch
                    {
                        return false;
                    }
                }

                //Pega Nome da Tag Principal da Mensagem
                NomeTagPrincipal = XmlMesgSaida.DocumentElement.Name;

                //Cria Mensagem Auxiliar para incluir a Repeticao
                XmlMesgAux = new XmlDocument();
                A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, "", NomeTagPrincipal, "");

                foreach (XmlNode Node in XmlMesgSaida.DocumentElement.SelectSingleNode("//" + NomeTagPrincipal).ChildNodes)
                {
                    if (Node.Name.Substring(0, 3).ToUpper() != "GR_")
                    {
                        A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal, Node.Name, Node.InnerXml);
                    }
                    else
                    {
                        if (XmlMesgAux.SelectSingleNode("//" + NomeTagPrincipal + "/" + "REPE" + Node.Name.Substring(2)) != null)
                        {
                            XmlRepeticao = new XmlDocument();
                            Repeticao.Remove(0, Repeticao.Length);
                            Repeticao.Append(Node.OuterXml);
                            XmlRepeticao.LoadXml(Repeticao.ToString());
                            A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal + "/" + "REPE" + Node.Name.Substring(2), Node.Name, XmlRepeticao.DocumentElement.InnerXml);
                            XmlRepeticao = null;
                        }
                        else
                        {
                            XmlRepeticao = new XmlDocument();
                            Repeticao.Remove(0, Repeticao.Length);
                            Repeticao.Append(Node.OuterXml);
                            XmlRepeticao.LoadXml(Repeticao.ToString());
                            ControleRepeticao.IncluiRepeticaRecursivo(ref XmlRepeticao);
                            A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal, XmlRepeticao.DocumentElement.Name, XmlRepeticao.DocumentElement.InnerXml);
                            XmlRepeticao = null;
                        }
                    }
                }

                _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml = CabecalhoSaida + XmlMesgAux.OuterXml;

                return true;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< RetiraRepeticao >>>
        protected void RetiraRepeticao(string nomeTagMensagem)
        {
            Repeticao ControleRepeticao = new Repeticao();
            XmlDocument XmlMesgSaida;
            XmlDocument XmlMesgAux;
            XmlDocument XmlRepeticao;
            StringBuilder Repeticao = new StringBuilder();
            string MesgAux;
            string MesgSaida;
            string CabecalhoSaida = "";
            string NomeTagPrincipal;

            try
            {
                //Pega o XML da Mensagem
                MesgAux = _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml;
                MesgSaida = MesgAux;

                //Faz o Load do XML
                XmlMesgSaida = new XmlDocument();

                try
                {
                    XmlMesgSaida.LoadXml(MesgSaida);
                }
                catch
                {
                    try
                    {
                        CabecalhoSaida = MesgAux.Substring(0, 20);
                        MesgSaida = MesgAux.Substring(20);
                        XmlMesgSaida.LoadXml(MesgSaida);
                    }
                    catch
                    {
                        return;
                    }
                }

                //Pega Nome da Tag Principal da Mensagem
                NomeTagPrincipal = XmlMesgSaida.DocumentElement.Name;

                //Cria Mensagem Auxiliar para incluir a Repeticao
                XmlMesgAux = new XmlDocument();
                A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, "", NomeTagPrincipal, "");

                foreach (XmlNode Node in XmlMesgSaida.DocumentElement.SelectSingleNode("//" + NomeTagPrincipal).ChildNodes)
                {
                    if (Node.Name.Length >= 5)
                    {
                        if (Node.Name.Substring(0, 5).ToUpper() != "REPE_")
                        {
                            A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal, Node.Name, Node.InnerXml);
                        }
                        else
                        {
                            XmlRepeticao = new XmlDocument();
                            Repeticao.Remove(0, Repeticao.Length);
                            Repeticao.Append(Node.OuterXml);
                            XmlRepeticao.LoadXml(Repeticao.ToString());
                            ControleRepeticao.RetiraRepeticaRecursivo(ref XmlRepeticao);
                            A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal, XmlRepeticao.DocumentElement.Name, XmlRepeticao.DocumentElement.InnerXml);
                            XmlRepeticao = null;
                        }
                    }
                    else
                    {
                        A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal, Node.Name, Node.InnerXml);
                    }
                }

                XmlMesgAux.InnerXml = XmlMesgAux.OuterXml.Replace("<REPET>", string.Empty).Replace("</REPET>", string.Empty);

                _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml = CabecalhoSaida + XmlMesgAux.OuterXml;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< IncluiGrupo >>>
        protected bool IncluiGrupo(string nomeTagMensagem, string codigoMensagem)
        {
            Repeticao ControleGrupo = new Repeticao();
            XmlDocument XmlMesgSaida;
            XmlDocument XmlMesgAux;
            XmlDocument XmlGrupo;
            StringBuilder Grupo = new StringBuilder();
            string MesgAux;
            string MesgSaida;
            string CabecalhoSaida = "";
            string NomeTagPrincipal;

            try
            {
                //Pega o XML da Mensagem
                MesgAux = _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml;
                MesgSaida = MesgAux;

                //Faz o Load do XML
                XmlMesgSaida = new XmlDocument();

                try
                {
                    XmlMesgSaida.LoadXml(MesgSaida);
                }
                catch
                {
                    try
                    {
                        CabecalhoSaida = MesgAux.Substring(0, 200);
                        MesgSaida = MesgAux.Substring(200);
                        XmlMesgSaida.LoadXml(MesgSaida);
                    }
                    catch
                    {
                        return false;
                    }
                }

                //Pega Nome da Tag Principal da Mensagem
                NomeTagPrincipal = XmlMesgSaida.DocumentElement.Name;

                //Cria Mensagem Auxiliar para incluir a Repeticao
                XmlMesgAux = new XmlDocument();
                A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, "", NomeTagPrincipal, "");
                A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal, codigoMensagem, "");

                foreach (XmlNode Node in XmlMesgSaida.DocumentElement.SelectSingleNode("//" + NomeTagPrincipal + "/" + codigoMensagem).ChildNodes)
                {
                    if (Node.Name.Length >= 6)
                    {
                        if (Node.Name.Substring(0, 6).ToUpper() != "REPET_")
                        {
                            A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal + "/" + codigoMensagem, Node.Name, Node.InnerXml);
                        }
                        else
                        {
                            if (XmlMesgSaida.SelectSingleNode("//" + NomeTagPrincipal + "/" + codigoMensagem + "/" + Node.Name + "/Grupo_" + Node.Name.Substring(6)) == null)
                            {
                                XmlGrupo = new XmlDocument();
                                Grupo.Remove(0, Grupo.Length);
                                Grupo.Append(Node.OuterXml);
                                XmlGrupo.LoadXml(Grupo.ToString());
                                ControleGrupo.IncluiGrupoRecursivo(ref XmlGrupo);
                                A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal + "/" + codigoMensagem, Node.Name, XmlGrupo.DocumentElement.InnerXml);
                                XmlGrupo = null;
                            }
                            else
                            {
                                A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal + "/" + codigoMensagem, Node.Name, Node.InnerXml);
                            }
                        }
                    }
                    else
                    {
                        A7NET.Comum.Comum.AppendNode(ref XmlMesgAux, NomeTagPrincipal + "/" + codigoMensagem, Node.Name, Node.InnerXml);
                    }
                }

                _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml = CabecalhoSaida + XmlMesgAux.OuterXml;

                return true;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< RetiraGrupo >>>
        protected void RetiraGrupo(string nomeTagMensagem, string codigoMensagem)
        {
            string MesgAux;
            string MesgSaida;

            try
            {
                //Pega o XML da Mensagem
                MesgAux = _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml;
                MesgSaida = MesgAux;


                switch (codigoMensagem)
                {
                    case "CAM0021": case "CAM0022": case "CAM0023": case "CAM0024": case "CAM0025": case "CAM0026":
                    case "CAM0021R2": case "CAM0022R2": case "CAM0023R2": case "CAM0024R2": case "CAM0025R2": case "CAM0026R2":

                        MesgSaida = MesgSaida.Replace("<Grupo_" + codigoMensagem + "_CodClausEspfcoIF>", "");
                        MesgSaida = MesgSaida.Replace("</Grupo_" + codigoMensagem + "_CodClausEspfcoIF>", "");
                        break;

                    case "CAM0028":

                        MesgSaida = MesgSaida.Replace("<Grupo_" + codigoMensagem + "_NumDespc>", "");
                        MesgSaida = MesgSaida.Replace("</Grupo_" + codigoMensagem + "_NumDespc>", "");
                        break;

                    case "CAM0030": case "CAM0031": case "CAM0032": case "CAM0030R2": case "CAM0031R2": case "CAM0032R2":

                        MesgSaida = MesgSaida.Replace("<Grupo_" + codigoMensagem + "_NumDespc>", "");
                        MesgSaida = MesgSaida.Replace("</Grupo_" + codigoMensagem + "_NumDespc>", "");
                        MesgSaida = MesgSaida.Replace("<Grupo_" + codigoMensagem + "_CodClausEspfcoIF>", "");
                        MesgSaida = MesgSaida.Replace("</Grupo_" + codigoMensagem + "_CodClausEspfcoIF>", "");
                        break;

                    case "CAM0033":

                        MesgSaida = MesgSaida.Replace("<Grupo_" + codigoMensagem + "_RegOpCamlVincd>", "");
                        MesgSaida = MesgSaida.Replace("</Grupo_" + codigoMensagem + "_RegOpCamlVincd>", "");
                        break;

                    case "CAM0039": case "CAM0039R2":

                        MesgSaida = MesgSaida.Replace("<Grupo_" + codigoMensagem + "_RegOpCaml>", "");
                        MesgSaida = MesgSaida.Replace("</Grupo_" + codigoMensagem + "_RegOpCaml>", "");
                        break;

                    default:
                        break;

                }

                _XmlMensagem.DocumentElement.SelectSingleNode(nomeTagMensagem).InnerXml = MesgSaida;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

    }
}
