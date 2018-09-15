using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;

namespace A8NET.Data.DAO
{
    public class TextXML
    {
        private void Inserir(string xmlBase64)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(AcessaBD.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    OracleCommand OraCommand = new OracleCommand();

                    OraCommand.Connection = OracleConn;
                    OraCommand.CommandType = CommandType.StoredProcedure;
                    OraCommand.CommandText = "A8PROC.PKG_A8_TB_TEXT_XML.SPI_TB_TEXT_XML";

                    OraCommand.Parameters.Add(A8NETOracleParameter.CO_TEXT_XML(codigoOperacao, ParameterDirection.Input));
                    OraCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_TEXT_XML(codigoOperacao, ParameterDirection.Input));
                    OraCommand.Parameters.Add(A8NETOracleParameter.TX_XML(codigoOperacao, ParameterDirection.Input));
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ObterOperacao()" + ex.ToString());
            }
        }
    }
}
