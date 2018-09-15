using System;
using System.Collections.Generic;
using System.Text;
using A7NET.Data;

namespace A7NET.Mensagem
{
    public class MensagemA8NZ : Mensagem, IDisposable
    {
        #region <<< Construtor >>>
        public MensagemA8NZ(DsParametrizacoes DataSetCache)
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

        ~MensagemA8NZ()
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
                    //Traduzir a mensagem recebida
                    if (base.TraduzirMensagem())
                    {
                        if (base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() != "A7")
                        {
                            //Exclui Tag de Grupo da mensagem recebida pelo A8 e enviada para NZ - Somente CAM
                            if (base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() == "NZ")
                            {
                                switch (base._ProtocoloMensagem.TipoMensagem.Trim())
                                {
                                    case "CAM0021": case "CAM0022": case "CAM0023": case "CAM0024": case "CAM0025": case "CAM0026":
                                    case "CAM0028": case "CAM0030": case "CAM0031": case "CAM0032": case "CAM0033": case "CAM0039":
                                        base.RetiraGrupo("TX_CNTD_SAID", base._ProtocoloMensagem.TipoMensagem.Trim());
                                        break;

                                    default:
                                        break;

                                }
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
