using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Xml;
using A7NET.Data;

namespace A7NET.Mensagem
{
    public class MensagemPadrao : Mensagem, IDisposable
    {
        #region <<< Construtor >>>
        public MensagemPadrao(DsParametrizacoes DataSetCache)
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion

            base.DataSetCache = DataSetCache;

        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~MensagemPadrao()
        {
            this.Dispose();
        }
        #endregion

        #region <<< Variaveis >>>
        
        #endregion

        #region <<< ProcessaMensagem >>>
        public override void ProcessaMensagem(string nomeFila, string mensagemRecebida, string messageId)
        {
            A7NET.Data.MensagemDAO MensagemDAO = new A7NET.Data.MensagemDAO();
            string ProtocoloMensagem;
            bool TraduzMensagem;

            try
            {
                mensagemRecebida = mensagemRecebida.Replace("&", " ");

                ProtocoloMensagem = mensagemRecebida.Substring(0, 20).ToString().Trim();

                base._ProtocoloMensagem = new A7NET.Mensagem.udt.udtProtocoloMensagem();
                base._ProtocoloMensagem.Parse(ProtocoloMensagem);

                base._XmlMensagem.LoadXml(ObterPropriedades("Grupo_Mensagem"));

                base._XmlMensagem.DocumentElement.SelectSingleNode("TP_MESG").InnerText = base._ProtocoloMensagem.TipoMensagem;
                base._XmlMensagem.DocumentElement.SelectSingleNode("CO_EMPR_ORIG").InnerText = base._ProtocoloMensagem.CodigoEmpresa;
                base._XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerText = base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper();
                base._XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerText = base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper();

                base._XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_ENTR").InnerXml = mensagemRecebida;

                //Autenticar a Mensagem
                if (base.AutenticarMensagem())
                {
                    TraduzMensagem = true;
                    //Inclui Tag de Repeticao na mensagem recebida pelo GPC ou R2
                    if (base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "GPC"
                    ||  base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "R2"
                    || base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "BOL"
                    ||  base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "E2"
                    ||  base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "E2A")
                    {
                        TraduzMensagem = base.IncluiRepeticao("TX_CNTD_ENTR");
                    }

                    if (TraduzMensagem)
                    {
                        //Traduzir a mensagem recebida
                        if (base.TraduzirMensagem())
                        {
                            //Retira Repeticao da mensagem recebida
                            if (base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "GPC"
                            ||  base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "R2"
                            || base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "BOL"
                            ||  base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "E2"
                            ||  base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper() == "E2A")
                            {
                                base.RetiraRepeticao("TX_CNTD_ENTR");
                            }

                            if (base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() != "A7")
                            {
                                if (base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "GPC"
                                ||  base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "R2"
                                ||  base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "BOL"
                                ||  base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "E2"
                                ||  base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "E2A")
                                {
                                    base.RetiraRepeticao("TX_CNTD_SAID");
                                }

                                base.PostarMensagemTraduzida();
                            }

                            base.SalvarMensagem(A7NET.Comum.Comum.EnumOcorrencia.PostagemBemSucedida);
                        }
                        else
                        {
                            base.SalvarMensagem(A7NET.Comum.Comum.EnumOcorrencia.CanceladaErroTraducao);

                            base.EnviarMensagemRejeicaoLegado();
                        }
                    }
                }
                else
                {
                    SalvarMensagemRejeitada(mensagemRecebida, nomeFila, messageId);
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

    }
}
