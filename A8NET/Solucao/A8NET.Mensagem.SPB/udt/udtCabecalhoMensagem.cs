using System;
using System.Xml;
using System.Xml.XPath;
using System.Reflection;
using A8NET.Comum;

namespace A8NET.Mensagem.SPB.udt
{
    public class udtCabecalhoMensagem : IDisposable
    {
        #region >>> Variaveis Privadas>>>
        private XmlDocument _xmlUdtCabecalhoMensagem;
        private string _CabecalhoMensagem = null;

        public string _SiglaSistemaEnviouNZ = string.Empty;
        public string _CodigoMensagem = string.Empty;
        public string _ControleRemessaNZ = string.Empty;
        public DateTime _DataRemessa = new DateTime();
        public string _CodigoEmpresa = string.Empty;
        public string _CodigoMoeda = string.Empty;
        public int _FormatoMensagem = 0;
        public string _AssinaturaInterna = string.Empty;
        public string _SiglaSistemaLegadoOrigem = string.Empty;
        public string _ReferenciaContabil = string.Empty;
        public string _BancoAgencia = string.Empty;
        public int _QuantidadeMensagem = 0;
        public string _NuOP = string.Empty;
        public string _FILLER = string.Empty;
        public string _Dominio = string.Empty;
        public string _FILLER_1 = string.Empty;
        #endregion

        #region >>> Construtores >>>
        public udtCabecalhoMensagem()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtCabecalhoMensagem.xml";
            _xmlUdtCabecalhoMensagem = new XmlDocument();
            _xmlUdtCabecalhoMensagem.Load(XmlAux);

            //XmlNodeList XmlNodeList = _xmlUdtCabecalhoMensagem.DocumentElement.SelectNodes("*");
            //string ValorAtributo = null;

            //foreach (XmlNode XmlAtributo in XmlNodeList)
            //{
            //    PropertyInfo PropertyInfo = this.GetType().GetProperty(XmlAtributo.Name);
            //    ValorAtributo = FormatarValor("", XmlAtributo.Name, false);
            //    PropertyInfo.SetValue(this, ValorAtributo, null);
            //}
        }
        #endregion

        #region >>> IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~udtCabecalhoMensagem()
        {
            this.Dispose();
        }

        #endregion

        #region >>> Propriedades >>>
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
        public DateTime DataRemessa
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
        public int FormatoMensagem
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
        public int QuantidadeMensagem
        {
            get { return _QuantidadeMensagem; }
            set { _QuantidadeMensagem = value; }
        }
        public string NuOP
        {
            get { return _NuOP; }
            set { _NuOP = value; }
        }
        public string FILLER
        {
            get { return _FILLER; }
            set { _FILLER = value; }
        }
        public string Dominio
        {
            get { return _Dominio; }
            set { _Dominio = value; }
        }
        public string FILLER_1
        {
            get { return _FILLER_1; }
            set { _FILLER_1 = value; }
        }
        #endregion

        #region >>> Parse >>>
        public void Parse(string remessaMensagemBatch)
        {
            try
            {
                _CabecalhoMensagem = remessaMensagemBatch;

                XmlNodeList XmlNodeList = _xmlUdtCabecalhoMensagem.DocumentElement.SelectNodes("*");
                object ValorAtributo = null;

                foreach (XmlNode XmlAtributo in XmlNodeList)
                {
                    PropertyInfo PropertyInfo = this.GetType().GetProperty(XmlAtributo.Name);
                    ValorAtributo = this.FormatarValorUdt("", XmlAtributo.Name, true);
                    PropertyInfo.SetValue(this, ValorAtributo, null);
                }
            }
            catch 
            {
                throw;
            }
        }
        #endregion

        #region>>> Formatar Valor >>>
        public object FormatarValorUdt(string valor, string nomeAtributo, bool ehParse)
        {

            string Tipo = _xmlUdtCabecalhoMensagem.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_xmlUdtCabecalhoMensagem.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_xmlUdtCabecalhoMensagem.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_xmlUdtCabecalhoMensagem.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            object Retorno;
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _CabecalhoMensagem.Substring(Inicio, Tamanho);
            }

            if (Tipo == "N")
            {
                if (Decimais == 0)
                {
                    Retorno = int.Parse(valor.PadLeft(Tamanho, '0').Replace(' ', '0'));
                }
                else
                {
                    Retorno = valor.PadLeft(Tamanho, '0').Replace(' ', '0');
                    StrInteiro = Retorno.ToString().Substring(0, Tamanho - Decimais);
                    IntAux = Tamanho - Decimais;
                    StrDecimais = Retorno.ToString().Substring(IntAux, Decimais);
                    Retorno = decimal.Parse(StrInteiro + "," + StrDecimais);
                }
            }
            else if (Tipo == "D")
            {
                Retorno = (object)Comum.Comum.ConvertDtToDateTime(valor);
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
