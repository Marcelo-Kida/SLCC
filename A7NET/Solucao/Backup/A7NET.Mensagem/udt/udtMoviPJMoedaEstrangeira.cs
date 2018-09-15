using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtMoviPJMoedaEstrangeira : IDisposable
    {
        #region <<< Fields >>>
        private string _TipoRemessa = null;                  // String * 3
        private string _CodigoEmpresa = null;                // String * 5
        private string _SiglaSistema = null;                 // String * 3
        private string _IdentificadorMovimento = null;       // String * 25
        private string _CodigoMoeda = null;                  // String * 4
        private string _CodigoBanqueiroSwift = null;         // String * 30
        private string _CodigoProduto = null;                // String * 4
        private string _DataMovimento = null;                // String * 8
        private string _CodigoReferenciaSwift = null;        // String * 16
        private string _TipoEntradaSaida = null;             // String * 1
        private string _ValorMovimento = null;               // String * 19
        private string _NomeCliente = null;                  // String * 50
        private string _TipoMovimento = null;                // String * 3
        private string _TipoProcessamento = null;            // String * 1
        private string _ContaBanqueiro = null;               // String * 35
        private string _Filler = null;                       // String * 93

        private string _RemessaMoviPJMoedaEstrangeira = null;
        private XmlDocument _XmlUdtMoviPJMoedaEstrangeira;
        #endregion

        #region <<< Construtores >>>
        public udtMoviPJMoedaEstrangeira()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtMoviPJMoedaEstrangeira.xml";
            _XmlUdtMoviPJMoedaEstrangeira = new XmlDocument();
            _XmlUdtMoviPJMoedaEstrangeira.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtMoviPJMoedaEstrangeira.DocumentElement.SelectNodes("*");
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

        ~udtMoviPJMoedaEstrangeira()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string TipoRemessa
        {
            get { return _TipoRemessa; }
            set { _TipoRemessa = value; }
        }

        public string CodigoEmpresa
        {
            get { return _CodigoEmpresa; }
            set { _CodigoEmpresa = value; }
        }

        public string SiglaSistema
        {
            get { return _SiglaSistema; }
            set { _SiglaSistema = value; }
        }

        public string IdentificadorMovimento
        {
            get { return _IdentificadorMovimento; }
            set { _IdentificadorMovimento = value; }
        }

        public string CodigoMoeda
        {
            get { return _CodigoMoeda; }
            set { _CodigoMoeda = value; }
        }

        public string CodigoBanqueiroSwift
        {
            get { return _CodigoBanqueiroSwift; }
            set { _CodigoBanqueiroSwift = value; }
        }

        public string CodigoProduto
        {
            get { return _CodigoProduto; }
            set { _CodigoProduto = value; }
        }

        public string DataMovimento
        {
            get { return _DataMovimento; }
            set { _DataMovimento = value; }
        }

        public string CodigoReferenciaSwift
        {
            get { return _CodigoReferenciaSwift; }
            set { _CodigoReferenciaSwift = value; }
        }

        public string TipoEntradaSaida
        {
            get { return _TipoEntradaSaida; }
            set { _TipoEntradaSaida = value; }
        }

        public string ValorMovimento
        {
            get { return _ValorMovimento; }
            set { _ValorMovimento = value; }
        }

        public string NomeCliente
        {
            get { return _NomeCliente; }
            set { _NomeCliente = value; }
        }

        public string TipoMovimento
        {
            get { return _TipoMovimento; }
            set { _TipoMovimento = value; }
        }

        public string TipoProcessamento
        {
            get { return _TipoProcessamento; }
            set { _TipoProcessamento = value; }
        }

        public string ContaBanqueiro
        {
            get { return _ContaBanqueiro; }
            set { _ContaBanqueiro = value; }
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
                _RemessaMoviPJMoedaEstrangeira = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtMoviPJMoedaEstrangeira.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtMoviPJMoedaEstrangeira.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtMoviPJMoedaEstrangeira.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtMoviPJMoedaEstrangeira.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtMoviPJMoedaEstrangeira.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaMoviPJMoedaEstrangeira.Substring(Inicio, Tamanho);
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
