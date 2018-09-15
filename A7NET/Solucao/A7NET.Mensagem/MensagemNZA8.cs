using System;
using System.Collections.Generic;
using System.Text;
using A7NET.Data;

namespace A7NET.Mensagem
{
    public class MensagemNZA8 : Mensagem, IDisposable
    {
        #region <<< Construtor >>>
        public MensagemNZA8(DsParametrizacoes DataSetCache)
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

        ~MensagemNZA8()
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
            A7NET.Mensagem.udt.udtProtocoloNZ ProtocoloNZ;
            A7NET.Comum.Comum.EnumFluxoMonitor TipoFluxoCarimbo;
            A7NET.Comum.Comum.EnumStatusMonitor StatusCarimbo;
            string ProtocoloMensagem;
            string ProtocoloMensagemNZ;

            try
            {
                mensagemRecebida = mensagemRecebida.Replace("&", " ");

                ProtocoloMensagem = mensagemRecebida.Substring(0, 20).ToString().Trim();
                base._ProtocoloMensagem = new A7NET.Mensagem.udt.udtProtocoloMensagem();
                base._ProtocoloMensagem.Parse(ProtocoloMensagem);

                ProtocoloMensagemNZ = mensagemRecebida.Substring(20, 200).ToString().Trim();
                ProtocoloNZ = new A7NET.Mensagem.udt.udtProtocoloNZ();
                ProtocoloNZ.Parse(ProtocoloMensagemNZ);

                //Seta Codigo Empresa a partir do codigo enviado pelo NZ
                base._ProtocoloMensagem.CodigoEmpresa = ProtocoloNZ.CodigoEmpresa;

                base._XmlMensagem.LoadXml(ObterPropriedades("Grupo_Mensagem"));
                base._XmlMensagem.DocumentElement.SelectSingleNode("TP_MESG").InnerText = base._ProtocoloMensagem.TipoMensagem;
                base._XmlMensagem.DocumentElement.SelectSingleNode("CO_EMPR_ORIG").InnerText = base._ProtocoloMensagem.CodigoEmpresa;
                base._XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_ORIG").InnerText = base._ProtocoloMensagem.SiglaSistemaOrigem.Trim().ToUpper();
                base._XmlMensagem.DocumentElement.SelectSingleNode("SG_SIST_DEST").InnerText = base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper();
                base._XmlMensagem.DocumentElement.SelectSingleNode("TX_CNTD_ENTR").InnerXml = mensagemRecebida.Substring(20).ToString();

                if (ProtocoloNZ.CodigoMensagem.Substring(7, 2).ToUpper() == "R1")
                {
                    TipoFluxoCarimbo = A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_NORMAL;
                    StatusCarimbo = A7NET.Comum.Comum.EnumStatusMonitor.MONITOR_RECEB_R1;
                }
                else //R2, Avisos e Informacoes
                {
                    TipoFluxoCarimbo = A7NET.Comum.Comum.EnumFluxoMonitor.FLUXO_MONITOR_MSG_EXTERNA;
                    StatusCarimbo = A7NET.Comum.Comum.EnumStatusMonitor.MONITOR_RECEB_MENS_EXTERNA;
                }

                //Inclui Aviso no Monitor
                base.IncluirAvisoMonitor(ProtocoloNZ.CodigoEmpresa, ProtocoloNZ.ControleRemessaNZ, ProtocoloNZ.CodigoMensagem,
                                        TipoFluxoCarimbo, StatusCarimbo, "A8", "A8", nomeFila, DateTime.Today, "", "", "", null);

                //Autenticar a Mensagem
                if (base.AutenticarMensagem())
                {
                    //Inclui Tag de Grupo na mensagem recebida pelo NZ e enviada para A8 - Somente CAM
                    switch (ProtocoloNZ.CodigoMensagem.Trim().ToUpper())
                    {
                        case "CAM0021R2": case "CAM0022R2": case "CAM0023R2": case "CAM0024R2": case "CAM0025R2":
                        case "CAM0026R2": case "CAM0030R2": case "CAM0031R2": case "CAM0032R2": case "CAM0039R2":
                            base.IncluiGrupo("TX_CNTD_ENTR", ProtocoloNZ.CodigoMensagem.Trim().ToUpper());
                            break;

                        default:
                            break;

                    }

                    //Traduzir a mensagem recebida
                    if (base.TraduzirMensagem())
                    {
                        //Exclui Tag de Grupo da mensagem recebida pelo NZ, para manter a mensagem original
                        switch (ProtocoloNZ.CodigoMensagem.Trim().ToUpper())
                        {
                            case "CAM0021R2": case "CAM0022R2": case "CAM0023R2": case "CAM0024R2": case "CAM0025R2":
                            case "CAM0026R2": case "CAM0030R2": case "CAM0031R2": case "CAM0032R2": case "CAM0039R2":
                                base.RetiraGrupo("TX_CNTD_ENTR", ProtocoloNZ.CodigoMensagem.Trim().ToUpper());
                                break;

                            default:
                                break;

                        }

                        if (base._ProtocoloMensagem.SiglaSistemaDestino.Trim().ToUpper() != "A7")
                        {
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
