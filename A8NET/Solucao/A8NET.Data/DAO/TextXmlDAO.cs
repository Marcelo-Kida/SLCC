using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;

namespace A8NET.Data.DAO
{
    public class TextXmlDAO : BaseDAO
    {
        #region <<< Estrutor >>>
        public struct EstruturaTextXml
        {
            public int CO_TEXT_XML;
            public int NU_SEQU_TEXT_XML;
            public string TX_XML;
        }
        #endregion

        #region <<< InserirBase64 >>>
        public int InserirBase64(string mensagemXml)
        {
            string MensagemXml64 = "";
            int Ordem = 1;
            int CodigoRetorno = 0;
            int PosicaoFinal = 0;

            try
            {
                MensagemXml64 = Comum.Comum.Base64Encode(mensagemXml);

                using (OracleConnection OracleConn = new OracleConnection(new BaseDAO().GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();
                    OracleCommand Comand = new OracleCommand("A8PROC.PKG_A8_TB_TEXT_XML.SPI_TB_TEXT_XML", OracleConn);
                    Comand.CommandType = CommandType.StoredProcedure;

                    for (int i = 0; i <= MensagemXml64.Length - 1; i += 4000)
                    {
                        if (MensagemXml64.Length - i > 4000) PosicaoFinal = 4000;
                        else PosicaoFinal = MensagemXml64.Length - i;

                        Inserir(ref Comand, ref CodigoRetorno, Ordem, MensagemXml64.Substring(i, PosicaoFinal));

                        Ordem++;
                        Comand.Parameters.Clear();
                    }
                }
                return CodigoRetorno;
            }
            catch (Exception ex)
            {

                throw new Exception("TextXmlDAO.InserirBase64() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< Inserir >>>
        public void Inserir(ref OracleCommand OraCommand, ref int codigoTextoXml, int ordemSequecia, string xml)
        {
            try
            {
                OracleParameter ParametroInOut = A8NETOracleParameter.CO_TEXT_XML(codigoTextoXml, ParameterDirection.InputOutput);
                
                OraCommand.Parameters.Add(ParametroInOut);
                OraCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_TEXT_XML(ordemSequecia, ParameterDirection.Input));
                OraCommand.Parameters.Add(A8NETOracleParameter.TX_XML(xml, ParameterDirection.Input));

                OraCommand.ExecuteNonQuery();

                codigoTextoXml = int.Parse(ParametroInOut.Value.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("TextXmlDAO.Inserir()" + ex.ToString());
            }
        }
        #endregion

        #region <<< SelecionarTextoXML >>>
        public OracleDataReader SelecionarTextoXML(int codigoTextoXml)
        {
            try
            {
                OracleConnection OracleConn = new OracleConnection(base.GetStringConnection());
                OracleConn.Open();
                _OracleCommand.Connection = OracleConn;
                _OracleCommand.Parameters.Clear();
                _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_TEXT_XML.SPS_TB_TEXT_XML";
                _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_TEXT_XML(codigoTextoXml, ParameterDirection.Input));
                _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());

                return _OracleCommand.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                throw new Exception("TextXML.Inserir()" + ex.ToString());
            }
        }
        #endregion

    }
}
