using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtSTR0008R2 : IDisposable
    {
        #region <<< Fields >>>
        private string _NU_RECB_PGTO = null;            // String * 15
        private string _CO_EMPR = null;                 // String * 4
        private string _DT_RECB_PGTO = null;            // String * 8
        private string _HO_RECB_MESG = null;            // String * 6
        private string _VA_RECB_PGTO = null;            // String * 18
        private string _CO_MESG = null;                 // String * 9
        private string _CO_ISPB_REMT = null;            // String * 8
        private string _CO_BANC_REMT = null;            // String * 4
        private string _NO_REMT = null;                 // String * 80
        private string _NU_CNPJ_CPF_REMT = null;        // String * 15
        private string _TP_CNTA_REMT = null;            // String * 2
        private string _CO_AGEN_REMT = null;            // String * 9
        private string _DG_AGEN_REMT = null;            // String * 1
        private string _NU_CNTA_REMT = null;            // String * 12
        private string _DG_CNTA_REMT = null;            // String * 1
        private string _NO_DEST = null;                 // String * 80
        private string _NU_CNPJ_CPF_DEST = null;        // String * 15
        private string _IN_TIPO_PESS_DEST = null;       // String * 1
        private string _TP_CNTA_DEST = null;            // String * 2
        private string _CO_AGEN_DEST = null;            // String * 9
        private string _NU_CNTA_DEST = null;            // String * 13
        private string _SQ_TIPO_TAG_FIND = null;        // String * 4
        private string _CO_DOMI_FIND = null;            // String * 20
        private string _TX_HIST_COMP_RECB = null;       // String * 200
        private string _CO_SITU_RECB_PGTO = null;       // String * 4
        private string _NU_CTRL_REME = null;            // String * 20
        private string _NU_CTRL_EXTE_RECB = null;       // String * 20
        private string _NU_DOCT_CRED = null;            // String * 9

        private string _RemessaSTR0008R2 = null;
        private XmlDocument _XmlUdtSTR0008R2;
        #endregion

        #region <<< Construtores >>>
        public udtSTR0008R2()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtSTR0008R2.xml";
            _XmlUdtSTR0008R2 = new XmlDocument();
            _XmlUdtSTR0008R2.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtSTR0008R2.DocumentElement.SelectNodes("*");
            string ValorAtributo = null;

            foreach (XmlNode XmlAtributo in XmlNodeList)
            {
                PropertyInfo PropertyInfo = this.GetType().GetProperty(XmlAtributo.Name);
                ValorAtributo = FormatarValor("", XmlAtributo.Name, false);
                PropertyInfo.SetValue(this, ValorAtributo, null);
            }

        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~udtSTR0008R2()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string NU_RECB_PGTO
        {
            get { return _NU_RECB_PGTO; }
            set { _NU_RECB_PGTO = value; }
        }

        public string CO_EMPR
        {
            get { return _CO_EMPR; }
            set { _CO_EMPR = value; }
        }

        public string DT_RECB_PGTO
        {
            get { return _DT_RECB_PGTO; }
            set { _DT_RECB_PGTO = value; }
        }

        public string HO_RECB_MESG
        {
            get { return _HO_RECB_MESG; }
            set { _HO_RECB_MESG = value; }
        }

        public string VA_RECB_PGTO
        {
            get { return _VA_RECB_PGTO; }
            set { _VA_RECB_PGTO = value; }
        }

        public string CO_MESG
        {
            get { return _CO_MESG; }
            set { _CO_MESG = value; }
        }

        public string CO_ISPB_REMT
        {
            get { return _CO_ISPB_REMT; }
            set { _CO_ISPB_REMT = value; }
        }

        public string CO_BANC_REMT
        {
            get { return _CO_BANC_REMT; }
            set { _CO_BANC_REMT = value; }
        }

        public string NO_REMT
        {
            get { return _NO_REMT; }
            set { _NO_REMT = value; }
        }

        public string NU_CNPJ_CPF_REMT
        {
            get { return _NU_CNPJ_CPF_REMT; }
            set { _NU_CNPJ_CPF_REMT = value; }
        }
        
        public string TP_CNTA_REMT
        {
            get { return _TP_CNTA_REMT; }
            set { _TP_CNTA_REMT = value; }
        }
        
        public string CO_AGEN_REMT
        {
            get { return _CO_AGEN_REMT; }
            set { _CO_AGEN_REMT = value; }
        }
        
        public string DG_AGEN_REMT
        {
            get { return _DG_AGEN_REMT; }
            set { _DG_AGEN_REMT = value; }
        }
        
        public string NU_CNTA_REMT
        {
            get { return _NU_CNTA_REMT; }
            set { _NU_CNTA_REMT = value; }
        }
        
        public string DG_CNTA_REMT
        {
            get { return _DG_CNTA_REMT; }
            set { _DG_CNTA_REMT = value; }
        }
                
        public string NO_DEST
        {
            get { return _NO_DEST; }
            set { _NO_DEST = value; }
        }
                
        public string NU_CNPJ_CPF_DEST
        {
            get { return _NU_CNPJ_CPF_DEST; }
            set { _NU_CNPJ_CPF_DEST = value; }
        }
                
        public string IN_TIPO_PESS_DEST
        {
            get { return _IN_TIPO_PESS_DEST; }
            set { _IN_TIPO_PESS_DEST = value; }
        }
                
        public string TP_CNTA_DEST
        {
            get { return _TP_CNTA_DEST; }
            set { _TP_CNTA_DEST = value; }
        }
                
        public string CO_AGEN_DEST
        {
            get { return _CO_AGEN_DEST; }
            set { _CO_AGEN_DEST = value; }
        }

        public string NU_CNTA_DEST
        {
            get { return _NU_CNTA_DEST; }
            set { _NU_CNTA_DEST = value; }
        }

        public string SQ_TIPO_TAG_FIND
        {
            get { return _SQ_TIPO_TAG_FIND; }
            set { _SQ_TIPO_TAG_FIND = value; }
        }

        public string CO_DOMI_FIND
        {
            get { return _CO_DOMI_FIND; }
            set { _CO_DOMI_FIND = value; }
        }

        public string TX_HIST_COMP_RECB
        {
            get { return _TX_HIST_COMP_RECB; }
            set { _TX_HIST_COMP_RECB = value; }
        }

        public string CO_SITU_RECB_PGTO
        {
            get { return _CO_SITU_RECB_PGTO; }
            set { _CO_SITU_RECB_PGTO = value; }
        }

        public string NU_CTRL_REME
        {
            get { return _NU_CTRL_REME; }
            set { _NU_CTRL_REME = value; }
        }

        public string NU_CTRL_EXTE_RECB
        {
            get { return _NU_CTRL_EXTE_RECB; }
            set { _NU_CTRL_EXTE_RECB = value; }
        }

        public string NU_DOCT_CRED
        {
            get { return _NU_DOCT_CRED; }
            set { _NU_DOCT_CRED = value; }
        }
        #endregion

        #region <<< Parse >>>
        public void Parse(string remessaProtocoloMensagem)
        {
            try
            {
                _RemessaSTR0008R2 = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtSTR0008R2.DocumentElement.SelectNodes("*");
                string ValorAtributo = null;

                foreach (XmlNode XmlAtributo in XmlNodeList)
                {
                    PropertyInfo PropertyInfo = this.GetType().GetProperty(XmlAtributo.Name);
                    ValorAtributo = FormatarValor("", XmlAtributo.Name, true);
                    PropertyInfo.SetValue(this, ValorAtributo, null);
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region<<< Formatar Valor >>>
        private string FormatarValor(string valor, string nomeAtributo, bool ehParse)
        {

            string Tipo = _XmlUdtSTR0008R2.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtSTR0008R2.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtSTR0008R2.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtSTR0008R2.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaSTR0008R2.Substring(Inicio, Tamanho);
            }

            if (Tipo == "N")
            {
                if (Decimais == 0)
                {
                    Retorno = valor.PadLeft(Tamanho, '0').Replace(' ', '0');
                }
                else
                {
                    Retorno = valor.PadLeft(Tamanho, '0').Replace(' ', '0');
                    StrInteiro = Retorno.Substring(0, Tamanho - Decimais);
                    IntAux = Tamanho - Decimais;
                    StrDecimais = Retorno.Substring(IntAux, Decimais);
                    Retorno = StrInteiro + "," + StrDecimais;
                }
            }
            else
            {
                Retorno = valor.PadRight(Tamanho, ' ');
            }

            return Retorno;

        }
        #endregion

    }
}
