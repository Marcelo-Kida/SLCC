using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtPZO00140 : IDisposable
    {
        #region <<< Fields >>>
        private string _HeaderNZ_PZ = null;        // String * 200
        private string _CodMsg = null;             // String * 9
        private string _NumCtrlIF = null;          // String * 20
        private string _ISPBIFDebtd = null;        // String * 8
        private string _NumCtrlSTR = null;         // String * 20
        private string _SitLancSTR = null;         // String * 3
        private string _DtHrSit = null;            // String * 14
        private string _DtMovto = null;            // String * 8
        private string _Filler = null;             // String * 316

        private string _RemessaPZO00140 = null;
        private XmlDocument _XmlUdtPZO00140;
        #endregion

        #region <<< Construtores >>>
        public udtPZO00140()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtPZO00140.xml";
            _XmlUdtPZO00140 = new XmlDocument();
            _XmlUdtPZO00140.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtPZO00140.DocumentElement.SelectNodes("*");
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

        ~udtPZO00140()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string HeaderNZ_PZ
        {
            get { return _HeaderNZ_PZ; }
            set { _HeaderNZ_PZ = value; }
        }

        public string CodMsg
        {
            get { return _CodMsg; }
            set { _CodMsg = value; }
        }

        public string NumCtrlIF
        {
            get { return _NumCtrlIF; }
            set { _NumCtrlIF = value; }
        }

        public string ISPBIFDebtd
        {
            get { return _ISPBIFDebtd; }
            set { _ISPBIFDebtd = value; }
        }

        public string NumCtrlSTR
        {
            get { return _NumCtrlSTR; }
            set { _NumCtrlSTR = value; }
        }

        public string SitLancSTR
        {
            get { return _SitLancSTR; }
            set { _SitLancSTR = value; }
        }

        public string DtHrSit
        {
            get { return _DtHrSit; }
            set { _DtHrSit = value; }
        }

        public string DtMovto
        {
            get { return _DtMovto; }
            set { _DtMovto = value; }
        }

        public string Filler
        {
            get { return _Filler; }
            set { _Filler = value; }
        }
        #endregion

        #region <<< Parse >>>
        public void Parse(string remessaProtocoloMensagem)
        {
            try
            {
                _RemessaPZO00140 = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtPZO00140.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtPZO00140.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtPZO00140.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtPZO00140.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtPZO00140.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaPZO00140.Substring(Inicio, Tamanho);
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
