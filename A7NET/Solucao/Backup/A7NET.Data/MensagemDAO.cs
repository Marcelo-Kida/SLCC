using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Text;
using System.Xml;
using A7NET.Comum;

namespace A7NET.Data
{
    public class MensagemDAO : BaseDAO
    {
        #region <<< Variaveis >>>

        #endregion

        #region <<< Estrutura >>>
        public struct EstruturaMensagem
        {
            public string SiglaSistema;
            public string MessageId;
            public string TipoMensagem;
            public int CodigoEmpresaOrigem;
            public DateTime DataInicioRegra;
            public string CodigoOperacaoAtiva;
            public int CodigoXmlEntrada;
            public int? CodigoXmlSaida;
            public int TipoFormatoMensagemSaida;
        }

        public struct EstruturaMensagemRejeitada
        {
            public string MessageId;
            public int CodigoOcorrencia;
            public int CodigoXml;
            public string SistemaOrigem;
            public string DetalheOcorrencia;
        }
        #endregion

        #region <<< SelecionaMensagemBase64 >>>
        public string SelecionaMensagemBase64(int sequencial)
        {
            string MensagemBase64 = "";

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    DataSet DsTextoMensagem = new DataSet();

                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_TEXT_XML.SPS_TB_TEXT_XML";

                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_TEXT_XML(sequencial, ParameterDirection.Input));

                    _OraDA.Fill(DsTextoMensagem);

                    foreach (DataRow Row in DsTextoMensagem.Tables[0].Rows)
                    {
                        MensagemBase64 = MensagemBase64 + Row["TX_XML"].ToString();
                    }

                    if (MensagemBase64.Trim() != "")
                    {
                        MensagemBase64 = A7NET.Comum.Comum.Base64Decode(MensagemBase64);
                    }

                    return MensagemBase64;

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.SelecionaMensagemBase64() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< SelecionaMensagemParametrizacao >>>
        public string SelecionaMensagemParametrizacao()
        {
            string MensagemParametrizacao = "";

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    DataSet DsTextoMensagem = new DataSet();

                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_TEXT_XML.SPS_TB_TEXT_XML_A8";

                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CURSOR());

                    _OraDA.Fill(DsTextoMensagem);

                    foreach (DataRow Row in DsTextoMensagem.Tables[0].Rows)
                    {
                        MensagemParametrizacao = MensagemParametrizacao + Row["TX_XML"].ToString();
                    }

                    return MensagemParametrizacao;

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.SelecionaMensagemParametrizacao() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< SelecionaRespostaMensagemConsulta >>>
        public string SelecionaRespostaMensagemConsulta(string identificadorMensagem)
        {
            XmlDocument XmlResposta;
            string MensagemEntrada;
            string Protocolo;
            string MensagemResposta = "";

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    DataSet DsTextoMensagem = new DataSet();

                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_MESG.SPS_TB_MESG";

                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_CMPO_ATRB_IDEF_MESG(identificadorMensagem, ParameterDirection.Input));

                    _OraDA.Fill(DsTextoMensagem);

                    XmlResposta = new XmlDocument();

                    if (DsTextoMensagem.Tables[0].Rows.Count == 0)
                    {
                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "", "MESG", "");
                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "MESG", "DE_OCOR_MESG", "Mensagem não Encontrada");
                    }
                    else
                    {
                        DataRow Row = DsTextoMensagem.Tables[0].Rows[0];
                        MensagemEntrada = SelecionaMensagemBase64(int.Parse(Row["CO_TEXT_XML_ENTR"].ToString()));
                        Protocolo = MensagemEntrada.Substring(0, 20);

                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "", "MESG", "");
                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "MESG", "TX_PRTC", Protocolo);
                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "MESG", "CO_MESG", Row["CO_MESG"].ToString());
                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "MESG", "DE_OCOR_MESG", Row["DE_OCOR_MESG"].ToString());
                        A7NET.Comum.Comum.AppendNode(ref XmlResposta, "MESG", "DE_ABRV_OCOR_MESG", Row["DE_ABRV_OCOR_MESG"].ToString().Trim());
                    }

                    MensagemResposta = XmlResposta.OuterXml;

                    return MensagemResposta;

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.SelecionaRespostaMensagemConsulta() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< PersisteMensagemBase64 >>>
        public int PersisteMensagemBase64(string mensagemXml)
        {
            string MensagemBase64 = "";
            int CodigoRetorno = 0;
            int Ordem = 1;
            int PosicaoFinal = 0;

            try
            {
                MensagemBase64 = A7NET.Comum.Comum.Base64Encode(mensagemXml);

                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_TEXT_XML.SPI_TB_TEXT_XML";

                    for (int i = 0; i <= MensagemBase64.Length - 1; i += 4000)
                    {
                        if (MensagemBase64.Length - i > 4000)
                        {
                            PosicaoFinal = 4000;
                        }
                        else
                        {
                            PosicaoFinal = MensagemBase64.Length - i;
                        }

                        OracleParameter ParametroInOut = A7NETOracleParameter.CO_TEXT_XML(CodigoRetorno, ParameterDirection.InputOutput);
                        _OracleCommand.Parameters.Add(ParametroInOut);
                        _OracleCommand.Parameters.Add(A7NETOracleParameter.NU_SEQU_TEXT_XML(Ordem, ParameterDirection.Input));
                        _OracleCommand.Parameters.Add(A7NETOracleParameter.TX_XML(MensagemBase64.Substring(i, PosicaoFinal), ParameterDirection.Input));

                        _OracleCommand.ExecuteNonQuery();

                        CodigoRetorno = int.Parse(ParametroInOut.Value.ToString());

                        Ordem++;
                        _OracleCommand.Parameters.Clear();

                    }

                    return CodigoRetorno;

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.PersisteMensagemBase64() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< PersisteMensagem >>>
        public int PersisteMensagem(EstruturaMensagem dadosMensagem)
        {
            string CodigoMensagem;

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_MESG.SPI_TB_MESG";

                    //Parametros
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_MESG_OUT());
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.SG_SIST_ORIG(dadosMensagem.SiglaSistema.Trim().ToUpper(), ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_MESG_MQSE(dadosMensagem.MessageId.Trim().Equals(string.Empty) ? string.Empty : dadosMensagem.MessageId.Trim(), ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.TP_MESG(dadosMensagem.TipoMensagem, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_EMPR_ORIG(dadosMensagem.CodigoEmpresaOrigem, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.DH_INIC_VIGE_REGR_TRAP(dadosMensagem.DataInicioRegra, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_CMPO_ATRB_IDEF_MESG(dadosMensagem.CodigoOperacaoAtiva.Trim().Equals(string.Empty) ? string.Empty : dadosMensagem.CodigoOperacaoAtiva.Trim(), ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_TEXT_XML_ENTR(dadosMensagem.CodigoXmlEntrada, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_TEXT_XML_SAID(dadosMensagem.CodigoXmlSaida, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.TP_FORM_MESG_SAID(dadosMensagem.TipoFormatoMensagemSaida, ParameterDirection.Input));

                    _OracleCommand.ExecuteNonQuery();

                    CodigoMensagem = _OracleCommand.Parameters[A7NETOracleParameter.CO_MESG_OUT().ParameterName].Value.ToString();

                    return int.Parse(CodigoMensagem);

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.PersisteMensagem() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< PersisteMensagemRejeitada >>>
        public void PersisteMensagemRejeitada(EstruturaMensagemRejeitada dadosMensagem)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_MESG_REJE.SPI_TB_MESG_REJE";

                    //Parametros
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_MESG_MQSE(dadosMensagem.MessageId.Trim(), ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_OCOR_MESG(dadosMensagem.CodigoOcorrencia, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_TEXT_XML(dadosMensagem.CodigoXml, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.NO_ARQU_ENTR_FILA_MQSE(dadosMensagem.SistemaOrigem, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.TX_DTLH_OCOR_ERRO(dadosMensagem.DetalheOcorrencia, ParameterDirection.Input));

                    _OracleCommand.ExecuteNonQuery();

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.PersisteMensagemRejeitada() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< PersisteSituacaoMensagem >>>
        public void PersisteSituacaoMensagem(int codigoMensagem, string detalheOcorrenciaErro, A7NET.Comum.Comum.EnumOcorrencia codigoOcorrencia)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A7PROC.PKG_A7NET_TB_SITU_MESG.SPI_TB_SITU_MESG";

                    //Parametros
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_MESG(codigoMensagem, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.CO_OCOR_MESG((int)codigoOcorrencia, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A7NETOracleParameter.TX_DTLH_OCOR_ERRO(detalheOcorrenciaErro.Trim().Equals(string.Empty) ? string.Empty : detalheOcorrenciaErro.Trim(), ParameterDirection.Input));

                    _OracleCommand.ExecuteNonQuery();

                }
            }
            catch (Exception ex)
            {

                throw new Exception("MensagemDAO.PersisteSituacaoMensagem() - " + ex.ToString());
            }
        }
        #endregion

    }
}
