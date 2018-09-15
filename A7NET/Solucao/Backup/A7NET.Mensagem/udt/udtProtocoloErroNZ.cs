using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtProtocoloErroNZ : IDisposable
    {
        #region <<< Fields >>>
        private string _CodigoMensagem = null;            //String * 9
        private string _ControleRemessaNZ = null;         //String * 20
        private string _CodigoEmpresa = null;             //String * 5
        private string _OrigemErro = null;                //String * 1
        private string _DataRemessa = null;               //String * 8
        private string _NomeDoCampo1 = null;              //String * 80
        private string _CodigoErro1 = null;               //String * 9
        private string _ConteúdoCampoErro1 = null;        //String * 30
        private string _NomeDoCampo2 = null;              //String * 80
        private string _CodigoErro2 = null;               //String * 9
        private string _ConteúdoCampoErro2 = null;        //String * 30
        private string _NomeDoCampo3 = null;              //String * 80
        private string _CodigoErro3 = null;               //String * 9
        private string _ConteúdoCampoErro3 = null;        //String * 30
        private string _NomeDoCampo4 = null;              //String * 80
        private string _CodigoErro4 = null;               //String * 9
        private string _ConteúdoCampoErro4 = null;        //String * 30
        private string _NomeDoCampo5 = null;              //String * 80
        private string _CodigoErro5 = null;               //String * 9
        private string _ConteúdoCampoErro5 = null;        //String * 30

        private string _RemessaProtocoloErroNZ = null;
        private XmlDocument _XmlUdtProtocoloErroNZ;
        #endregion

        #region <<< Construtores >>>
        public udtProtocoloErroNZ()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtProtocoloErroNZ.xml";
            _XmlUdtProtocoloErroNZ = new XmlDocument();
            _XmlUdtProtocoloErroNZ.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtProtocoloErroNZ.DocumentElement.SelectNodes("*");
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

        ~udtProtocoloErroNZ()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string CodigoMensagem
        {
            get { return _CodigoMensagem; }
            set { _CodigoMensagem = value; }
        }

        public string ControleRemessaNZ
        {
            get { return _ControleRemessaNZ; }
            set { _ControleRemessaNZ = value; }
        }

        public string CodigoEmpresa
        {
            get { return _CodigoEmpresa; }
            set { _CodigoEmpresa = value; }
        }

        public string OrigemErro
        {
            get { return _OrigemErro; }
            set { _OrigemErro = value; }
        }

        public string DataRemessa
        {
            get { return _DataRemessa; }
            set { _DataRemessa = value; }
        }

        public string NomeDoCampo1
        {
            get { return _NomeDoCampo1; }
            set { _NomeDoCampo1 = value; }
        }

        public string CodigoErro1
        {
            get { return _CodigoErro1; }
            set { _CodigoErro1 = value; }
        }

        public string ConteúdoCampoErro1
        {
            get { return _ConteúdoCampoErro1; }
            set { _ConteúdoCampoErro1 = value; }
        }

        public string NomeDoCampo2
        {
            get { return _NomeDoCampo2; }
            set { _NomeDoCampo2 = value; }
        }

        public string CodigoErro2
        {
            get { return _CodigoErro2; }
            set { _CodigoErro2 = value; }
        }

        public string ConteúdoCampoErro2
        {
            get { return _ConteúdoCampoErro2; }
            set { _ConteúdoCampoErro2 = value; }
        }

        public string NomeDoCampo3
        {
            get { return _NomeDoCampo3; }
            set { _NomeDoCampo3 = value; }
        }

        public string CodigoErro3
        {
            get { return _CodigoErro3; }
            set { _CodigoErro3 = value; }
        }

        public string ConteúdoCampoErro3
        {
            get { return _ConteúdoCampoErro3; }
            set { _ConteúdoCampoErro3 = value; }
        }

        public string NomeDoCampo4
        {
            get { return _NomeDoCampo4; }
            set { _NomeDoCampo4 = value; }
        }

        public string CodigoErro4
        {
            get { return _CodigoErro4; }
            set { _CodigoErro4 = value; }
        }

        public string ConteúdoCampoErro4
        {
            get { return _ConteúdoCampoErro4; }
            set { _ConteúdoCampoErro4 = value; }
        }

        public string NomeDoCampo5
        {
            get { return _NomeDoCampo5; }
            set { _NomeDoCampo5 = value; }
        }

        public string CodigoErro5
        {
            get { return _CodigoErro5; }
            set { _CodigoErro5 = value; }
        }

        public string ConteúdoCampoErro5
        {
            get { return _ConteúdoCampoErro5; }
            set { _ConteúdoCampoErro5 = value; }
        }
        #endregion

        #region <<< Parse >>>
        public void Parse(string remessaProtocoloMensagem)
        {
            try
            {
                _RemessaProtocoloErroNZ = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtProtocoloErroNZ.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtProtocoloErroNZ.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtProtocoloErroNZ.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtProtocoloErroNZ.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtProtocoloErroNZ.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaProtocoloErroNZ.Substring(Inicio, Tamanho);
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
