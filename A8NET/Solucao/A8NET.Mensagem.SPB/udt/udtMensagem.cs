using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Data;

namespace A8NET.Mensagem.SPB.udt
{
    public class udtMensagem:IDisposable
    {

        #region <<< Variáveis >>>
        // VARIAVEIS COMUNS
        private System.Data.DataRow _LinhaMensagem;
        private string _NomeFila = string.Empty;
        private string _RemessaMensagem = string.Empty;
        private string _CodigoMensagem = string.Empty;
        private string _CodigoGrupoMensagem = string.Empty;
        private string _TipoMensagem = string.Empty;
        private udtCabecalhoMensagem _CabecalhoMensagem;
        private int _CodigoErro;
        XmlDocument _XmlMensagem;
        #endregion

        #region <<< Propriedades >>>
        public System.Data.DataRow LinhaMensagem
        {
            get { return _LinhaMensagem; }
            set { _LinhaMensagem = value; }
        }
        public XmlDocument XmlMensagem
        {
            get { return _XmlMensagem; }
            set { _XmlMensagem = value; }
        }
        public int CodigoErro
        {
            get { return _CodigoErro; }
            set { _CodigoErro = value; }
        }

        public string CodigoMensagem
        {
            get { return _CodigoMensagem; }
            set { _CodigoMensagem = value; }
        }

        public string CodigoGrupoMensagem
        {
            get { return _CodigoGrupoMensagem; }
            set { _CodigoGrupoMensagem = value; }
        }
        public string TipoMensagem
        {
            get { return _TipoMensagem; }
            set { _TipoMensagem = value; }
        }
        public udtCabecalhoMensagem CabecalhoMensagem
        {
            get
            {
                return _CabecalhoMensagem;
            }
            set
            {
                _CabecalhoMensagem = value;
            }
        }
        #endregion

        #region <<< Construtores >>>
        public udtMensagem()
        {

        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~udtMensagem()
        {
            this.Dispose();
        }
        #endregion

        #region <<< Parse >>>
        public bool Parse(string mensagemRemessa)
        {
            StringBuilder sb = new StringBuilder();
            XmlWriter xw = new XmlTextWriter(new System.IO.StringWriter(sb));
            DataSet dsRetorno = new DataSet();
            CabecalhoMensagem = new udtCabecalhoMensagem();

            try
            {
                XmlMensagem = new XmlDocument();
                // carrega o xml 
                XmlMensagem.LoadXml(mensagemRemessa);

                // seta o valor do header
                CabecalhoMensagem.Parse(XmlMensagem.DocumentElement.SelectSingleNode("TX_MESG").InnerText.Substring(0, 200));
                
                // pega o valor do corpo do xml
                XmlMensagem.DocumentElement.SelectSingleNode("TX_MESG/*").WriteTo(xw);

                xw.Close();
                // carrega o corpo do xml da mensagem no datarow
                dsRetorno.ReadXml(new System.IO.StringReader(sb.ToString()));
                this._LinhaMensagem = dsRetorno.Tables[0].Rows[0];

                CodigoMensagem = _LinhaMensagem["CodMsg"].ToString();
                CodigoGrupoMensagem = CodigoMensagem.Substring(0, 3);
                TipoMensagem = CodigoMensagem.Substring(7, 2);
                return true;
            }
            catch
            {
                //XML mal formado
                CodigoErro = 204;
                return false;
            }
        }
        #endregion

    }
}