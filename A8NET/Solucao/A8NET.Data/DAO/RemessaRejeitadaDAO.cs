using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;

namespace A8NET.Data.DAO
{
    public class RemessaRejeitadaDAO : BaseDAO
    {

        #region <<< Estrutura >>>
        public struct EstruturaRemessaRejeitada
        {
            public object SG_SIST_ORIG_INFO;
            public object TP_MESG_INTE;
            public object CO_EMPR;
            public object CO_TEXT_XML_REJE;
            public object CO_TEXT_XML_RETN_SIST_ORIG;
            public object TX_XML_ERRO;
            public object DH_REME_REJE;
        }
        #endregion

        #region <<< Inserir >>>
        public void Inserir(EstruturaRemessaRejeitada parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_REME_REJE.SPI_TB_REME_REJE";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
					    A8NETOracleParameter.SG_SIST_ORIG_INFO(parametro.SG_SIST_ORIG_INFO, ParameterDirection.Input),
					    A8NETOracleParameter.TP_MESG_INTE(parametro.TP_MESG_INTE, ParameterDirection.Input),
					    A8NETOracleParameter.CO_EMPR(parametro.CO_EMPR, ParameterDirection.Input),
					    A8NETOracleParameter.CO_TEXT_XML_REJE(parametro.CO_TEXT_XML_REJE, ParameterDirection.Input),
					    A8NETOracleParameter.CO_TEXT_XML_RETN_SIST_ORIG(parametro.CO_TEXT_XML_RETN_SIST_ORIG, ParameterDirection.Input),
					    A8NETOracleParameter.TX_XML_ERRO(parametro.TX_XML_ERRO, ParameterDirection.Input),
					    A8NETOracleParameter.DH_REME_REJE(parametro.DH_REME_REJE, ParameterDirection.Input),});
                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("RemessaRejeitadaDAO.Inserir() - " + ex.ToString());
            }
        }
        #endregion

    }
}
