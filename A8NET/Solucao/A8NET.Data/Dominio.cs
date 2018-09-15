using System;
using System.Collections;
using System.Threading;
using System.Data;
using System.Data.OracleClient;
using System.Xml;
using System.Text;
using System.Diagnostics;
using A8NET.Comum;

namespace A8NET.Data
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
            MBS_Client.BA AcessoMBS;
            DataSet DsUsuarioAcesso;
            long CodigoRetorno;

            try
            {
                if (_DataSetCache != null) return;
                DAO.BaseDAO Base = new A8NET.Data.DAO.BaseDAO();
                using (OracleConnection OracleConn = new OracleConnection(Base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    OracleCommand OraCommand = new OracleCommand();
                    OraCommand.Connection = OracleConn;
                    OraCommand.CommandType = CommandType.StoredProcedure;

                    //Adicionar parametro cursor
                    OraCommand.Parameters.Add(A8NETOracleParameter.CURSOR());

                    _DataSetCache = new DsParametrizacoes();
                    _DataSetCache.EnforceConstraints = false;

                    OracleDataAdapter OraDA = new OracleDataAdapter(OraCommand);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_CTRL_PROC_OPER_ATIV";
                    OraDA.Fill(_DataSetCache.TB_CTRL_PROC_OPER_ATIV);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_TIPO_OPER";
                    OraDA.Fill(_DataSetCache.TB_TIPO_OPER);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_FCAO_SIST_TIPO_OPER";
                    OraDA.Fill(_DataSetCache.TB_FCAO_SIST_TIPO_OPER);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_PARM_FCAO_SIST";
                    OraDA.Fill(_DataSetCache.TB_PARM_FCAO_SIST);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_PARM_FCAO_SIST_EXCE";
                    OraDA.Fill(_DataSetCache.TB_PARM_FCAO_SIST_EXCE);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_MENSAGEM";
                    OraDA.Fill(_DataSetCache.TB_MENSAGEM);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_SITU_SPB_SITU_PROC";
                    OraDA.Fill(_DataSetCache.TB_SITU_SPB_SITU_PROC);

                    // Carrega os Usuarios Cadastrados no MBS ou via contingencia.
                    try
                    {
                        OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_MBS_GRUPO";
                        OraDA.Fill(_DataSetCache.MBS_GRUPO);
                    }
                    catch
                    {
                        _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHGPCBX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                        _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPGPCBX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                        
                        // NICK Incluindo BOL
                        _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHBOL01", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                        _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPBOL01", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");

                        // NICK Incluindo HQ
                        _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHHQ0001", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                        _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPHQ0001", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");

                        try
                        {
                            AcessoMBS = new MBS_Client.BA();
                            DsUsuarioAcesso = new DataSet();

                            CodigoRetorno = AcessoMBS.recuperaUsuarioGrupo("A8_GRUPOVEICLEGAL_WORKFLOW", ref DsUsuarioAcesso, 0);

                            if (CodigoRetorno != 0)
                            {
                                // NICK Incluindo BOL
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHHQ0001", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHBOL01", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHGPC0BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHR20BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");

                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPHQ0001", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPBOL01", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPGPC0BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPR20BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            }
                            else
                            {
                                foreach (DataRow RowItemAcesso in DsUsuarioAcesso.Tables[0].Rows)
                                {
                                    _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow(RowItemAcesso["CD_USR"].ToString(), 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                                }
                            }
                        }
                        catch
                        {
                            // NICK Incluindo BOL
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHHQ0001", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHBOL01", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHGPC0BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SHR20BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");

                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPHQ0001", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPBOL01", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPGPC0BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                            _DataSetCache.MBS_GRUPO.AddMBS_GRUPORow("SPR20BX", 5480, "A8_GRUPOVEICLEGAL_WORKFLOW");
                        }
                    }
                    
                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_VEIC_LEGA";
                    OraDA.Fill(_DataSetCache.TB_VEIC_LEGA);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_CTRL_DOMI";
                    OraDA.Fill(_DataSetCache.TB_CTRL_DOMI);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_PRODUTO";
                    OraDA.Fill(_DataSetCache.TB_PRODUTO);

                    OraCommand.CommandText = "A7PROC.PKG_A7_CACHE_DOMINIO.SPS_TB_REGR_SIST_DEST";
                    OraDA.Fill(_DataSetCache.TB_REGR_SIST_DEST);

                    OraCommand.CommandText = "A8PROC.PKG_A8_CACHE_DOMINIO.SPS_TB_TIPO_OPER_CNTD_ATRB";
                    OraDA.Fill(_DataSetCache.TB_TIPO_OPER_CNTD_ATRB);
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
