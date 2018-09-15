using System;
using System.Xml;
using System.Xml.XPath;
using System.Reflection;
using System.Data;
using System.Text;

namespace A8NET.Mensagem.Operacao
{
    public class udtOperacao
    {

        #region >>> Variáveis >>>
       // private XmlDocument _XmlUdtOperacao;
        private int codigoErro = 0;

        // VARIAVEIS COMUNS
        private XmlDocument xmlOperacao;
        private string remessaOperacao;
        private System.Data.DataRow _RowOperacao;
        //private string nomeFila;
        //private string headerMensagem;
        private DataSet dataSetMensagemXML;
        //private int tipoSessao;
        #endregion

        #region >>> Propriedades >>>
        public DataSet DataSetMensagemXML
        {
            get { return dataSetMensagemXML; }
            set { dataSetMensagemXML = value; }
        }
        public System.Data.DataRow RowOperacao
        {
            get { return _RowOperacao; }
            set { _RowOperacao = value; }
        }
        public XmlDocument XmlOperacao
        {
            get { return xmlOperacao; }
            set { xmlOperacao = value; }
        }
        public int CodigoErro
        {
            get { return codigoErro; }
            set { codigoErro = value; }
        }
        #endregion

        #region <<< Construtores >>>
        public udtOperacao()
        {

        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~udtOperacao()
        {
            this.Dispose();
        }

        #endregion


        #region <<< Parse >>>
        public bool Parse(string remessa)
        {
            bool retorno = false;
            StringBuilder sb = new StringBuilder();
            XmlWriter xw = new XmlTextWriter(new System.IO.StringWriter(sb));
            DataSet dsRetorno = new DataSet();
            //int headerLength = 0;

            try
            {
                //  verifica se está carregando o modelo de header700 ou headerMensagem
                //headerLength = Convert.ToInt16("0" + xmlUdtHeader.DocumentElement.SelectSingleNode("@Length").Value);

                remessaOperacao = remessa;
                //headerMensagem = remessaOperacao;

                string xmlAux = remessa;

                xmlOperacao = new XmlDocument();

                try
                {
                    xmlOperacao.LoadXml(xmlAux);

                    //xmlOperacao.LoadXml(xmlOperacao.SelectSingleNode("//SISMSG").OuterXml);

                    xmlOperacao.WriteTo(xw);
                    xw.Close();

                    dsRetorno.ReadXml(new System.IO.StringReader(sb.ToString()));
                    this._RowOperacao = dsRetorno.Tables[0].Rows[0];
                    retorno = true;
                }
                catch (Exception)
                {
                    //XML mal formado
                    codigoErro = 204;
                    retorno = false;
                }

            }
            catch (Exception exParse)
            {
                throw exParse;
            }

            return retorno;
            
        }
        #endregion

    }
}