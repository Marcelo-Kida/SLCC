using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtMaioresValores : IDisposable
    {
        #region <<< Fields >>>
        private string _TipoRemessa = null;                            // String * 3
        private string _CodigoRemessa = null;                          // String * 23
        private string _DataRemessa = null;                            // String * 8
        private string _HoraRemessa = null;                            // String * 4
        private string _CodigoEmpresa = null;                          // String * 5
        private string _SiglaSistema = null;                           // String * 3
        private string _CodigoMoeda = null;                            // String * 4
        private string _CodigoBanqueiro = null;                        // String * 12
        private string _TipoCaixa = null;                              // String * 3
        private string _CodigoItemCaixa = null;                        // String * 9
        private string _CodigoProduto = null;                          // String * 4
        private string _TipoConta = null;                              // String * 3
        private string _CodigoSegmento = null;                         // String * 3
        private string _CodigoEventoFinanceiro = null;                 // String * 3
        private string _CodigoIndexador = null;                        // String * 3
        private string _CodigoLocalLiquidacao = null;                  // String * 4
        private string _TipoMovimento = null;                          // String * 3
        private string _DataMovimento = null;                          // String * 8
        private string _HoraMovimento = null;                          // String * 4
        private string _TipoEntradaSaida = null;                       // String * 1
        private string _ValorMovimento = null;                         // String * 17
        private string _CodigoBanco = null;                            // String * 3
        private string _CodigoAgencia = null;                          // String * 5
        private string _NumeroContaCorrente = null;                    // String * 13
        private string _TipoPessoa = null;                             // String * 1
        private string _CodigoCNPJ_CPF = null;                         // String * 15
        private string _NomeCliente = null;                            // String * 64
        private string _TipoProcessamento = null;                      // String * 1
        private string _TipoEnvio = null;                              // String * 1
        private string _Filler = null;                                 // String * 20

        private string _RemessaMaioresValores = null;
        private XmlDocument _XmlUdtMaioresValores;
        #endregion

        #region <<< Construtores >>>
        public udtMaioresValores()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtMaioresValores.xml";
            _XmlUdtMaioresValores = new XmlDocument();
            _XmlUdtMaioresValores.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtMaioresValores.DocumentElement.SelectNodes("*");
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

        ~udtMaioresValores()
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

        public string CodigoRemessa
        {
            get { return _CodigoRemessa; }
            set { _CodigoRemessa = value; }
        }

        public string DataRemessa
        {
            get { return _DataRemessa; }
            set { _DataRemessa = value; }
        }
        
        public string HoraRemessa
        {
            get { return _HoraRemessa; }
            set { _HoraRemessa = value; }
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

        public string CodigoMoeda
        {
            get { return _CodigoMoeda; }
            set { _CodigoMoeda = value; }
        }

        public string CodigoBanqueiro
        {
            get { return _CodigoBanqueiro; }
            set { _CodigoBanqueiro = value; }
        }

        public string TipoCaixa
        {
            get { return _TipoCaixa; }
            set { _TipoCaixa = value; }
        }

        public string CodigoItemCaixa
        {
            get { return _CodigoItemCaixa; }
            set { _CodigoItemCaixa = value; }
        }
        
        public string CodigoProduto
        {
            get { return _CodigoProduto; }
            set { _CodigoProduto = value; }
        }
        
        public string TipoConta
        {
            get { return _TipoConta; }
            set { _TipoConta = value; }
        }
        
        public string CodigoSegmento
        {
            get { return _CodigoSegmento; }
            set { _CodigoSegmento = value; }
        }
        
        public string CodigoEventoFinanceiro
        {
            get { return _CodigoEventoFinanceiro; }
            set { _CodigoEventoFinanceiro = value; }
        }
        
        public string CodigoIndexador
        {
            get { return _CodigoIndexador; }
            set { _CodigoIndexador = value; }
        }
                
        public string CodigoLocalLiquidacao
        {
            get { return _CodigoLocalLiquidacao; }
            set { _CodigoLocalLiquidacao = value; }
        }
                
        public string TipoMovimento
        {
            get { return _TipoMovimento; }
            set { _TipoMovimento = value; }
        }
                
        public string DataMovimento
        {
            get { return _DataMovimento; }
            set { _DataMovimento = value; }
        }
                
        public string HoraMovimento
        {
            get { return _HoraMovimento; }
            set { _HoraMovimento = value; }
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
                        
        public string CodigoBanco
        {
            get { return _CodigoBanco; }
            set { _CodigoBanco = value; }
        }
                        
        public string CodigoAgencia
        {
            get { return _CodigoAgencia; }
            set { _CodigoAgencia = value; }
        }
                        
        public string NumeroContaCorrente
        {
            get { return _NumeroContaCorrente; }
            set { _NumeroContaCorrente = value; }
        }
                        
        public string TipoPessoa
        {
            get { return _TipoPessoa; }
            set { _TipoPessoa = value; }
        }

        public string CodigoCNPJ_CPF
        {
            get { return _CodigoCNPJ_CPF; }
            set { _CodigoCNPJ_CPF = value; }
        }

        public string NomeCliente
        {
            get { return _NomeCliente; }
            set { _NomeCliente = value; }
        }

        public string TipoProcessamento
        {
            get { return _TipoProcessamento; }
            set { _TipoProcessamento = value; }
        }

        public string TipoEnvio
        {
            get { return _TipoEnvio; }
            set { _TipoEnvio = value; }
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
                _RemessaMaioresValores = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtMaioresValores.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtMaioresValores.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtMaioresValores.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtMaioresValores.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtMaioresValores.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaMaioresValores.Substring(Inicio, Tamanho);
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
