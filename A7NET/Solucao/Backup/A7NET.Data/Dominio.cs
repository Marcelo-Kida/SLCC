using System;
using System.Collections;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Xml;
using A7NET.Comum;

namespace A7NET.Data
{
    public class Dominio
    {
        #region <<< Variaveis >>>
        private DsParametrizacoes _DataSetCache;
        #endregion

        #region <<< Propriedades >>>
        public DsParametrizacoes DataSetCache
        {
            get { return _DataSetCache; }
            set { _DataSetCache = value; }
        }
        #endregion

        #region <<< CarregaDominio() >>>
        /// <summary>
        /// Método carrega o dataSet na memoria com todos os dados de dominio
        /// </summary>
        public void CarregaDominio()
        {
            try
            {
                if (_DataSetCache != null) return;
                A7NET.Data.BaseDAO Base = new A7NET.Data.BaseDAO();
                using (OracleConnection OracleConn = new OracleConnection(Base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    OracleCommand OraCommand = new OracleCommand();
                    OraCommand.Connection = OracleConn;
                    OraCommand.CommandType = CommandType.StoredProcedure;

                    //Adicionar parametro cursor
                    OraCommand.Parameters.Add(A7NETOracleParameter.CURSOR());

                    _DataSetCache = new DsParametrizacoes();
                    _DataSetCache.EnforceConstraints = false;

                    OracleDataAdapter OraDA = new OracleDataAdapter(OraCommand);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_MESG";
                    OraDA.Fill(_DataSetCache.TB_MESG);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_TIPO_MESG";
                    OraDA.Fill(_DataSetCache.TB_TIPO_MESG);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_EMPRESA_HO";
                    OraDA.Fill(_DataSetCache.TB_EMPRESA_HO);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_SIST";
                    OraDA.Fill(_DataSetCache.TB_SIST);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_REGR_TRAP_MESG";
                    OraDA.Fill(_DataSetCache.TB_REGR_TRAP_MESG);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_MENSAGEM_SPB";
                    OraDA.Fill(_DataSetCache.TB_MENSAGEM_SPB);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_ENDE_FILA_MQSE";
                    OraDA.Fill(_DataSetCache.TB_ENDE_FILA_MQSE);

                    OraCommand.CommandText = "A7PROC.PKG_A7NET_CACHE_DOMINIO.SPS_TB_TIPO_OPER";
                    OraDA.Fill(_DataSetCache.TB_TIPO_OPER);

                }
            }
            catch
            {
                throw;
            }
        }
        #endregion

    }
}
