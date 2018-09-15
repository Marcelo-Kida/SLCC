using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtProtocoloNZ : IDisposable 
    {
        #region <<< Fields >>>
        private string _SiglaSistemaEnviouNZ = null;     //String * 3
        private string _CodigoMensagem = null;           //String * 9
        private string _ControleRemessaNZ = null;        //String * 20
        private string _DataRemessa = null;              //String * 8
        private string _CodigoEmpresa = null;            //String * 5
        private string _CodigoMoeda = null;              //String * 5
        private string _FormatoMensagem = null;          //String * 1
        private string _AssinaturaInterna = null;        //String * 50
        private string _SiglaSistemaLegadoOrigem = null; //String * 3
        private string _ReferenciaContabil = null;       //String * 8
        private string _BancoAgencia = null;             //String * 15
        private string _QuantidadeMensagem = null;       //String * 6
        private string _NuOP = null;                     //String * 23
        private string _Filler = null;                   //String * 44

        private string _RemessaProtocoloNZ = null;
        private XmlDocument _XmlUdtProtocoloNZ;
        #endregion

        #region <<< Construtores >>>
        public udtProtocoloNZ()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtProtocoloNZ.xml";
            _XmlUdtProtocoloNZ = new XmlDocument();
            _XmlUdtProtocoloNZ.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtProtocoloNZ.DocumentElement.SelectNodes("*");
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

        ~udtProtocoloNZ()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string SiglaSistemaEnviouNZ
        {
            get { return _SiglaSistemaEnviouNZ; }
            set { _SiglaSistemaEnviouNZ = value; }

        }

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

        public string DataRemessa
        {
            get { return _DataRemessa; }
            set { _DataRemessa = value; }

        }

        public string CodigoEmpresa
        {
            get { return _CodigoEmpresa; }
            set { _CodigoEmpresa = value; }

        }

        public string CodigoMoeda
        {
            get { return _CodigoMoeda; }
            set { _CodigoMoeda = value; }

        }

        public string FormatoMensagem
        {
            get { return _FormatoMensagem; }
            set { _FormatoMensagem = value; }

        }

        public string AssinaturaInterna
        {
            get { return _AssinaturaInterna; }
            set { _AssinaturaInterna = value; }

        }

        public string SiglaSistemaLegadoOrigem
        {
            get { return _SiglaSistemaLegadoOrigem; }
            set { _SiglaSistemaLegadoOrigem = value; }

        }

        public string ReferenciaContabil
        {
            get { return _ReferenciaContabil; }
            set { _ReferenciaContabil = value; }

        }

        public string BancoAgencia
        {
            get { return _BancoAgencia; }
            set { _BancoAgencia = value; }

        }

        public string QuantidadeMensagem
        {
            get { return _QuantidadeMensagem; }
            set { _QuantidadeMensagem = value; }

        }

        public string NuOP
        {
            get { return _NuOP; }
            set { _NuOP = value; }

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
                _RemessaProtocoloNZ = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtProtocoloNZ.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtProtocoloNZ.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtProtocoloNZ.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtProtocoloNZ.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtProtocoloNZ.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaProtocoloNZ.Substring(Inicio, Tamanho);
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
