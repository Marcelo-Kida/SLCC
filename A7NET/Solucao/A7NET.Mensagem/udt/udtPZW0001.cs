using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace A7NET.Mensagem.udt
{
    public class udtPZW0001 : IDisposable
    {
        #region <<< Fields >>>
        private string _TipoRegistro = null;                           // String * 1
        private string _BancoOrigem = null;                            // String * 3
        private string _UnidadeOrigem = null;                          // String * 7
        private string _SiglaSistemaOrigem = null;                     // String * 3
        private string _CodigoSistemaOrigem = null;                    // String * 4
        private string _ControleLegado = null;                         // String * 23
        private string _NumeroDocumento = null;                        // String * 6
        private string _BancoDestino = null;                           // String * 3
        private string _ISPBDestino = null;                            // String * 8
        private string _DataContabil = null;                           // String * 8
        private string _HoraAgendamento = null;                        // String * 6
        private string _CodigoAcao = null;                             // String * 1
        private string _MeioTransferencia = null;                      // String * 1
        private string _IdentificadorAlteracao = null;                 // String * 1
        private string _NumeroVersao = null;                           // String * 2
        private string _LancamentoCC = null;                           // String * 1
        private string _HistoricoCC = null;                            // String * 5
        private string _ContaDebitada = null;                          // String * 13
        private string _AgenciaDebitada = null;                        // String * 5
        private string _OrigemContabilidade = null;                    // String * 2
        private string _Filler1 = null;                                // String * 39
        private string _Erro1 = null;                                  // String * 5
        private string _Erro2 = null;                                  // String * 5
        private string _Erro3 = null;                                  // String * 5
        private string _CodigoPZ = null;                               // String * 50
        private string _CodigoMensagem = null;                         // String * 9
        private string _ValorLancamento = null;                        // String * 18
        private string _DataMovimento = null;                          // String * 8
        private string _AgenciaCreditada = null;                       // String * 5
        private string _AgenciaRemetente = null;                       // String * 5
        private string _Filler2 = null;                                // String * 40

        private string _RemessaPZW0001 = null;
        private XmlDocument _XmlUdtPZW0001;
        #endregion

        #region <<< Construtores >>>
        public udtPZW0001()
        {

            string XmlAux = AppDomain.CurrentDomain.BaseDirectory + "\\xml\\udtPZW0001.xml";
            _XmlUdtPZW0001 = new XmlDocument();
            _XmlUdtPZW0001.Load(XmlAux);

            XmlNodeList XmlNodeList = _XmlUdtPZW0001.DocumentElement.SelectNodes("*");
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

        ~udtPZW0001()
        {
            this.Dispose();
        }

        #endregion

        #region <<< Propriedades >>>
        public string TipoRegistro
        {
            get { return _TipoRegistro; }
            set { _TipoRegistro = value; }
        }

        public string BancoOrigem
        {
            get { return _BancoOrigem; }
            set { _BancoOrigem = value; }
        }

        public string UnidadeOrigem
        {
            get { return _UnidadeOrigem; }
            set { _UnidadeOrigem = value; }
        }

        public string SiglaSistemaOrigem
        {
            get { return _SiglaSistemaOrigem; }
            set { _SiglaSistemaOrigem = value; }
        }

        public string CodigoSistemaOrigem
        {
            get { return _CodigoSistemaOrigem; }
            set { _CodigoSistemaOrigem = value; }
        }

        public string ControleLegado
        {
            get { return _ControleLegado; }
            set { _ControleLegado = value; }
        }

        public string NumeroDocumento
        {
            get { return _NumeroDocumento; }
            set { _NumeroDocumento = value; }
        }

        public string BancoDestino
        {
            get { return _BancoDestino; }
            set { _BancoDestino = value; }
        }

        public string ISPBDestino
        {
            get { return _ISPBDestino; }
            set { _ISPBDestino = value; }
        }

        public string DataContabil
        {
            get { return _DataContabil; }
            set { _DataContabil = value; }
        }

        public string HoraAgendamento
        {
            get { return _HoraAgendamento; }
            set { _HoraAgendamento = value; }
        }

        public string CodigoAcao
        {
            get { return _CodigoAcao; }
            set { _CodigoAcao = value; }
        }

        public string MeioTransferencia
        {
            get { return _MeioTransferencia; }
            set { _MeioTransferencia = value; }
        }

        public string IdentificadorAlteracao
        {
            get { return _IdentificadorAlteracao; }
            set { _IdentificadorAlteracao = value; }
        }

        public string NumeroVersao
        {
            get { return _NumeroVersao; }
            set { _NumeroVersao = value; }
        }

        public string LancamentoCC
        {
            get { return _LancamentoCC; }
            set { _LancamentoCC = value; }
        }

        public string HistoricoCC
        {
            get { return _HistoricoCC; }
            set { _HistoricoCC = value; }
        }

        public string ContaDebitada
        {
            get { return _ContaDebitada; }
            set { _ContaDebitada = value; }
        }

        public string AgenciaDebitada
        {
            get { return _AgenciaDebitada; }
            set { _AgenciaDebitada = value; }
        }

        public string OrigemContabilidade
        {
            get { return _OrigemContabilidade; }
            set { _OrigemContabilidade = value; }
        }

        public string Filler1
        {
            get { return _Filler1; }
            set { _Filler1 = value; }
        }

        public string Erro1
        {
            get { return _Erro1; }
            set { _Erro1 = value; }
        }

        public string Erro2
        {
            get { return _Erro2; }
            set { _Erro2 = value; }
        }

        public string Erro3
        {
            get { return _Erro3; }
            set { _Erro3 = value; }
        }

        public string CodigoPZ
        {
            get { return _CodigoPZ; }
            set { _CodigoPZ = value; }
        }

        public string CodigoMensagem
        {
            get { return _CodigoMensagem; }
            set { _CodigoMensagem = value; }
        }

        public string ValorLancamento
        {
            get { return _ValorLancamento; }
            set { _ValorLancamento = value; }
        }

        public string DataMovimento
        {
            get { return _DataMovimento; }
            set { _DataMovimento = value; }
        }

        public string AgenciaCreditada
        {
            get { return _AgenciaCreditada; }
            set { _AgenciaCreditada = value; }
        }

        public string AgenciaRemetente
        {
            get { return _AgenciaRemetente; }
            set { _AgenciaRemetente = value; }
        }
        
        public string Filler2
        {
            get { return _Filler2; }
            set { _Filler2 = value; }
        }
        #endregion

        #region <<< Parse >>>
        public void Parse(string remessaProtocoloMensagem)
        {
            try
            {
                _RemessaPZW0001 = remessaProtocoloMensagem;

                XmlNodeList XmlNodeList = _XmlUdtPZW0001.DocumentElement.SelectNodes("*");
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

            string Tipo = _XmlUdtPZW0001.SelectSingleNode("//" + nomeAtributo + "//@tipo").Value;
            int Tamanho = Convert.ToInt16(_XmlUdtPZW0001.SelectSingleNode("//" + nomeAtributo + "//@tamanho").Value);
            int Decimais = Convert.ToInt16(_XmlUdtPZW0001.SelectSingleNode("//" + nomeAtributo + "//@decimais").Value);
            int Inicio = Convert.ToInt16(_XmlUdtPZW0001.SelectSingleNode("//" + nomeAtributo + "//@inicio").Value);

            if (nomeAtributo.ToUpper() == "FILLER") return "";

            Tamanho = Tamanho + Decimais;

            string Retorno = "";
            int IntAux = 0;
            string StrInteiro = "";
            string StrDecimais = "";

            if (ehParse)
            {
                valor = _RemessaPZW0001.Substring(Inicio, Tamanho);
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
