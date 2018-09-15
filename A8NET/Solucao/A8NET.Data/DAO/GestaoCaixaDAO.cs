using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Text;

namespace A8NET.Data.DAO
{
    public class GestaoCaixaDAO : BaseDAO
    {
        #region <<< Variaveis >>>
        
        #endregion

        #region <<< Estrutura >>>
        
        #endregion

        #region <<< Obter Identificador Remessa PJ >>>
        public string ObterIdentificadorRemessaPJ(string dataMovimento)
        {
            OracleParameter ParametroRetorno;
            string IdentificadorRemessaPJ;

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_INTEGRACAO_PJ.SPS_SQ_A8_NU_SEQU_REME_PJ";

                    //Adicionar parametro retorno
                    ParametroRetorno = A8NETOracleParameter.RETORNO();
                    _OracleCommand.Parameters.Add(ParametroRetorno);

                    _OracleCommand.ExecuteNonQuery();

                    IdentificadorRemessaPJ = string.Concat(dataMovimento, "A8 1", ParametroRetorno.Value.ToString().PadLeft(8, '0'));

                    return IdentificadorRemessaPJ;

                }
            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixaDAO.ObterIdentificadorRemessaPJ() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< Verificar Envio Maiores Valores >>>
        public bool VerificarEnvioMaioresValores(int codigoEmpresa, int codigoProduto, decimal valorOperacao)
        {
            OracleParameter ParametroRetorno;
            decimal ValorMinMaioresValores = 0;

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_INTEGRACAO_PJ.SPS_TB_PRODUTO";

                    ParametroRetorno = A8NETOracleParameter.VA_MINI_MAIR_VALO(ValorMinMaioresValores, ParameterDirection.Output);
                    _OracleCommand.Parameters.Add(ParametroRetorno);
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_EMPR(codigoEmpresa, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_PROD(codigoProduto, ParameterDirection.Input));

                    _OracleCommand.ExecuteNonQuery();

                    if (ParametroRetorno.Value == null
                    ||  ParametroRetorno.Value.ToString() == String.Empty)
                    {
                        return true;
                    }
                    else
                    {
                        ValorMinMaioresValores = decimal.Parse(ParametroRetorno.Value.ToString());
                        if (ValorMinMaioresValores > valorOperacao)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixaDAO.ObterIdentificadorRemessaPJ() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< Inserir Registro >>>>>>
        /// <summary>
        ///	Método responsável em inserir um registro na tabela TB_HIST_ENVI_INFO_GEST_CAIX.
        /// </summary>
        /// <param name="registro<<NomeClasseBO>>">Valores para inclusão</param>
        public void Inserir(decimal sequenciaOperacao, A8NET.Comum.Comum.EnumTipoMovimento tipoMovimento, int codigoTextXml)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_INTEGRACAO_PJ.SPI_TB_HIST_ENVI_INFO_GEST";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]
                    {
			            A8NETOracleParameter.NU_SEQU_OPER_ATIV(sequenciaOperacao, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SITU_MOVI_GEST_CAIX((int)tipoMovimento, ParameterDirection.Input),
                        A8NETOracleParameter.CO_TEXT_XML(codigoTextXml, ParameterDirection.Input)
                    }
                    );

                    _OracleCommand.ExecuteNonQuery();

                }
            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixaDAO.Inserir() - " + ex.ToString());
            }
        }
        #endregion
    }
}
