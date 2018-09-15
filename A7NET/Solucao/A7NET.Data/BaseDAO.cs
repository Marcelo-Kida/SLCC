using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Runtime.InteropServices;
using System.Text;

namespace A7NET.Data
{
    public class BaseDAO
    {
        protected System.Collections.ArrayList _Lista = new System.Collections.ArrayList();
        protected System.Data.DataView _DView;
        protected OracleCommand _OracleCommand;
        protected OracleDataAdapter _OraDA;

        public BaseDAO()
        {
            _OracleCommand = new OracleCommand();
            _OracleCommand.CommandType = CommandType.StoredProcedure;
            _OraDA = new OracleDataAdapter(_OracleCommand);
        }

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~BaseDAO()
        {
            this.Dispose();
        }
        #endregion

        #region <<< WIN API >>>
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);
        #endregion

        #region <<< Variaveis >>>
        private static string _StringConexao = null;
        #endregion

        #region <<< GetStringConnection >>>

        public string GetStringConnection()
        {
            string retorno = "";

            if (_StringConexao == null || _StringConexao == "")
            {

                string filename = Environment.GetEnvironmentVariable("SLCC_Ambiente", EnvironmentVariableTarget.Machine);

                if (filename == null || filename.Trim().Equals(string.Empty))
                {
                    filename = Environment.GetEnvironmentVariable("SLCC_Ambiente", EnvironmentVariableTarget.User);
                }

                int chars = 256;
                StringBuilder buffer = new StringBuilder(chars);

                if (GetPrivateProfileString("OleDB", "Provider", "", buffer, chars, filename) != 0)
                {
                    retorno = buffer.ToString();
                    retorno = retorno.Replace("MSDAORA.1;", "");
                }

                _StringConexao = retorno;
            }
            else
            {
                retorno = _StringConexao;
            }

            return retorno + ";Enlist=False";

        }

        #endregion

        #region <<< ExecuteNonQuery >>>
        public void ExecuteNonQuery(OracleCommand OraCommand)
        {

            OracleConnection OracleConn = null;

            try
            {

                using (OracleConn = new OracleConnection(_StringConexao))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    OraCommand.Connection = OracleConn;

                    OraCommand.ExecuteNonQuery();
                }


            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                OraCommand.Dispose();

                if (OracleConn != null && OracleConn.State == ConnectionState.Open)
                {
                    OracleConn.Close();
                    OracleConn.Dispose();
                }


            }
        }
        #endregion
    }
}
