using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtProtocoloMensagem : IDisposable 
    {
        #region <<< Fields >>>
        private string _TipoMensagem = null;                 //String * 9
        private string _SiglaSistemaOrigem = null;           //String * 3
        private string _SiglaSistemaDestino = null;          //String * 3
        private string _CodigoEmpresa = null;                //String * 5

        private string _RemessaProtocoloMensagem = null;
        private XmlDocument _XmlUdtProtocoloMensagem;
        #endregion

        #region <<< Construtores >>>
        public udtProtocoloMensagem()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtProtocoloMensagem.xml";
            _XmlUdtProtocoloMensagem = new XmlDocument();
            _XmlUdtProtocoloMensagem.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtProtocoloMensagem.DocumentElement.SelectNodes("*");
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

        ~udtProtocoloMensagem()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string TipoMensagem
        {
            get { return _TipoMensagem; }
            set { _TipoMensagem = value; }

        }

        public string SiglaSistemaOrigem
        {
            get { return _SiglaSistemaOrigem; }
            set { _SiglaSistemaOrigem = value; }

        }

        public string SiglaSistemaDestino
        {
            get { return _SiglaSistemaDestino; }
            set { _SiglaSistemaDestino = value; }

        }

        public string CodigoEmpresa
        {
            get { return _CodigoEmpresa; }
            set { _CodigoEmpresa = value; }

        }
        #endregion

        #region <<< Parse >>>
        public void Parse(string remessaProtocoloMensagem)
        {
            try
            {
                _RemessaProtocoloMensagem = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtProtocoloMensagem.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtProtocoloMensagem.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtProtocoloMensagem.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtProtocoloMensagem.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtProtocoloMensagem.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaProtocoloMensagem.Substring(Inicio, Tamanho);
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
