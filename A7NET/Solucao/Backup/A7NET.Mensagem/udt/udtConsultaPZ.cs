using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtConsultaPZ : IDisposable
    {
        #region <<< Fields >>>
        private string _Banco = null;                  // String * 3
        private string _Agencia = null;                // String * 5
        private string _Conta = null;                  // String * 13
        private string _DataMensagem = null;           // String * 8
        private string _HoraMensagem = null;           // String * 8
        private string _CodigoMensagem = null;         // String * 9
        private string _SiglaSistema = null;           // String * 3
        private string _Filler = null;                 // String * 20
        private string _QuantidadeItem = null;         // String * 5
        private string _RCRotina = null;               // String * 5
        private string _MensagemRCRotina = null;       // String * 80

        private string _RemessaConsultaPZ = null;
        private XmlDocument _XmlUdtConsultaPZ;
        #endregion

        #region <<< Construtores >>>
        public udtConsultaPZ()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtConsultaPZ.xml";
            _XmlUdtConsultaPZ = new XmlDocument();
            _XmlUdtConsultaPZ.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtConsultaPZ.DocumentElement.SelectNodes("*");
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

        ~udtConsultaPZ()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string Banco
        {
            get { return _Banco; }
            set { _Banco = value; }
        }

        public string Agencia
        {
            get { return _Agencia; }
            set { _Agencia = value; }
        }

        public string Conta
        {
            get { return _Conta; }
            set { _Conta = value; }
        }

        public string DataMensagem
        {
            get { return _DataMensagem; }
            set { _DataMensagem = value; }
        }

        public string HoraMensagem
        {
            get { return _HoraMensagem; }
            set { _HoraMensagem = value; }
        }

        public string CodigoMensagem
        {
            get { return _CodigoMensagem; }
            set { _CodigoMensagem = value; }
        }

        public string SiglaSistema
        {
            get { return _SiglaSistema; }
            set { _SiglaSistema = value; }
        }

        public string Filler
        {
            get { return _Filler; }
            set { _Filler = value; }
        }

        public string QuantidadeItem
        {
            get { return _QuantidadeItem; }
            set { _QuantidadeItem = value; }
        }

        public string RCRotina
        {
            get { return _RCRotina; }
            set { _RCRotina = value; }
        }

        public string MensagemRCRotina
        {
            get { return _MensagemRCRotina; }
            set { _MensagemRCRotina = value; }
        }
        #endregion

        #region <<< Parse >>>
        public void Parse(string remessaProtocoloMensagem)
        {
            try
            {
                _RemessaConsultaPZ = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtConsultaPZ.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtConsultaPZ.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtConsultaPZ.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtConsultaPZ.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtConsultaPZ.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaConsultaPZ.Substring(Inicio, Tamanho);
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
