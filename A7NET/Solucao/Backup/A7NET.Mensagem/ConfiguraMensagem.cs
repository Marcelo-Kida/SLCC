using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace A7NET.Mensagem
{
    public class ConfiguraMensagem
    {
        #region <<< Variaveis >>>
        
        #endregion

        #region >>> Construtor >>>
        public ConfiguraMensagem()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion
        }
        #endregion

        #region <<< PrepararXML >>>
        public void PrepararXML(ref XmlDocument xmlMensagem)
        {
            int Count;

            try
            {
                //Verifica se a formatacao de saida contem tags a partir do 1o. nivel.
                //Caso sim, indica que a mensagem pode ter repetições
                if (xmlMensagem.SelectSingleNode("//Documento/*/Formato/*") != null)
                {
                    Count = 0;

                    //Adiciona indice nas tags
                    foreach (XmlElement Elemento in xmlMensagem.SelectNodes("//Documento/Mensagem//*"))
                    {
                        Count++;
                        Elemento.SetAttribute("Posicao", Convert.ToString(Count));
                    }

                    //Adiciona as tags de controle de repeticao caso elas nao existam
                    foreach (XmlElement Elemento in xmlMensagem.SelectNodes("//Documento/*/Formato/*/*"))
                    {
                        if (Elemento.SelectSingleNode("@RepetTag") == null)
                        {
                            Elemento.SetAttribute("RepetTag", "1");
                        }

                        if (Elemento.SelectSingleNode("@UltimaPosicao") == null)
                        {
                            Elemento.SetAttribute("UltimaPosicao", "0");
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< MontaXmlRegraTransporte >>>
        public string MontaXmlRegraTransporte(int codigoTextoXmlRegra, int codigoTextoXmlMensagem)
        {
            StringBuilder Regra = new StringBuilder();
            string FormatoInicio = "<{0}>";
            string FormatoFim = "</{0}>";
            string FormatoXml = "<{0}>{1}</{0}>";
            A7NET.Data.MensagemDAO MensagemDAO;

            try
            {
                Regra.AppendFormat(FormatoInicio, "RegraTransporte");
                MensagemDAO = new A7NET.Data.MensagemDAO();
                Regra.AppendFormat(FormatoXml, "TX_REGR_TRNF_MESG", MensagemDAO.SelecionaMensagemBase64(codigoTextoXmlRegra));
                Regra.AppendFormat(FormatoXml, "TX_VALID_SAID_MESG", MensagemDAO.SelecionaMensagemBase64(codigoTextoXmlMensagem));
                Regra.AppendFormat(FormatoFim, "RegraTransporte");

                return Regra.ToString();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< MontaMensagemErro >>>
        public string MontaMensagemErro(string mensagemErro, string nomeFila, string detalheOcorrencia)
        {
            XmlDocument XmlMensagemErro;

            try
            {

                XmlMensagemErro = new XmlDocument();
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemErro, "", "FILA_ERRO", "");
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemErro, "FILA_ERRO", "NO_FILA_ORIG_MQSE", nomeFila);
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemErro, "FILA_ERRO", "DH_MESG_ERRO", String.Format("{0:yyyyMMddHHmmss}", DateTime.Now));
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemErro, "FILA_ERRO", "TX_MESG_ORIG", mensagemErro);
                A7NET.Comum.Comum.AppendNode(ref XmlMensagemErro, "FILA_ERRO", "TX_MESG_ERRO", detalheOcorrencia);

                return XmlMensagemErro.OuterXml;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< ObterCodigoMensagemNZ >>>
        public string ObterCodigoMensagemNZ(string headerNZ, string tipoMensagem)
        {
            A7NET.Mensagem.udt.udtProtocoloNZ ProtocoloNZ;
            A7NET.Mensagem.udt.udtProtocoloErroNZ ProtocoloErroNZ;
            A7NET.Mensagem.udt.udtRemessaMovimento RemessaMovimento;
            A7NET.Mensagem.udt.udtMaioresValores MaioresValores;
            A7NET.Mensagem.udt.udtPZW0001 PZW0001;
            A7NET.Mensagem.udt.udtPZW0916 PZW0916;
            A7NET.Mensagem.udt.udtPZO00140 PZO00140;
            A7NET.Mensagem.udt.udtConsultaPZ ConsultaPZ;
            A7NET.Mensagem.udt.udtSTR0008R2 STR0008R2;
            A7NET.Mensagem.udt.udtMoviPJMoedaEstrangeira MoviPJMoedaEstrangeira;
            string CodigoMensagemNZ = "";

            try
            {
                switch (int.Parse(tipoMensagem))
                {
                    #region MensagemNZA8 / A8NZ
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemNZA8:
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemA8NZ:

                        ProtocoloNZ = new A7NET.Mensagem.udt.udtProtocoloNZ();
                        ProtocoloNZ.Parse(headerNZ);

                        CodigoMensagemNZ = ProtocoloNZ.CodigoMensagem.ToString();

                        break;
                    #endregion

                    #region MensagemErroNZA8
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemErroNZA8:

                        ProtocoloErroNZ = new A7NET.Mensagem.udt.udtProtocoloErroNZ();
                        ProtocoloErroNZ.Parse(headerNZ);

                        CodigoMensagemNZ = ProtocoloErroNZ.CodigoMensagem.ToString();

                        break;
                    #endregion

                    #region MensagemA8PJPrevisto
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemA8PJPrevisto:

                        RemessaMovimento = new A7NET.Mensagem.udt.udtRemessaMovimento();
                        RemessaMovimento.Parse(headerNZ.Substring(20));

                        if (int.Parse(RemessaMovimento.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.Previsto)
                        {
                            CodigoMensagemNZ = "Previsto";
                        }
                        else if (int.Parse(RemessaMovimento.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.Realizado)
                        {
                            CodigoMensagemNZ = "Realizado";
                        }
                        else if (int.Parse(RemessaMovimento.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto)
                        {
                            CodigoMensagemNZ = "Estorno Previsto";
                        }
                        else if (int.Parse(RemessaMovimento.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoRealizado)
                        {
                            CodigoMensagemNZ = "Estorno Realizado";
                        }

                        break;
                    #endregion

                    #region MensagemA8PJRealizado
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemA8PJRealizado:

                        MaioresValores = new A7NET.Mensagem.udt.udtMaioresValores();
                        MaioresValores.Parse(headerNZ.Substring(20));

                        if (int.Parse(MaioresValores.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.Previsto)
                        {
                            CodigoMensagemNZ = "Previsto";
                        }
                        else if (int.Parse(MaioresValores.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.Realizado)
                        {
                            CodigoMensagemNZ = "Realizado";
                        }
                        else if (int.Parse(MaioresValores.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto)
                        {
                            CodigoMensagemNZ = "Estorno Previsto";
                        }
                        else if (int.Parse(MaioresValores.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoRealizado)
                        {
                            CodigoMensagemNZ = "Estorno Realizado";
                        }

                        break;
                    #endregion

                    #region MensagemA8PJMoedaEstrangeira
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemA8PJMoedaEstrangeira:

                        MoviPJMoedaEstrangeira = new A7NET.Mensagem.udt.udtMoviPJMoedaEstrangeira();
                        MoviPJMoedaEstrangeira.Parse(headerNZ.Substring(20));

                        if (int.Parse(MoviPJMoedaEstrangeira.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.Previsto)
                        {
                            CodigoMensagemNZ = "ME - Previsto";
                        }
                        else if (int.Parse(MoviPJMoedaEstrangeira.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.Realizado)
                        {
                            CodigoMensagemNZ = "ME - Realizado";
                        }
                        else if (int.Parse(MoviPJMoedaEstrangeira.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoPrevisto)
                        {
                            CodigoMensagemNZ = "ME - Estorno Previsto";
                        }
                        else if (int.Parse(MoviPJMoedaEstrangeira.TipoMovimento.Trim()) == (int)A7NET.Comum.Comum.EnumTipoMovimentoPJ.EstornoRealizado)
                        {
                            CodigoMensagemNZ = "ME - Estorno Realizado";
                        }

                        break;
                    #endregion

                    #region MensagemPZErro
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZErro:

                        PZW0001 = new A7NET.Mensagem.udt.udtPZW0001();
                        PZW0001.Parse(headerNZ.Substring(20));

                        CodigoMensagemNZ = PZW0001.ControleLegado.ToString();

                        break;
                    #endregion

                    #region MensagemPZR1
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZR1:

                        PZO00140 = new A7NET.Mensagem.udt.udtPZO00140();
                        PZO00140.Parse(headerNZ);

                        PZW0916 = new A7NET.Mensagem.udt.udtPZW0916();
                        PZW0916.Parse(PZO00140.HeaderNZ_PZ);

                        CodigoMensagemNZ = PZW0916.NumeroControleLegado.ToString();

                        break;
                    #endregion

                    #region MensagemPZR2
                    case (int)A7NET.Comum.Comum.EnumTipoMensagemEntrada.MensagemPZR2:

                        if (headerNZ.Substring(20).Substring(37, 9) == "STR0008R2")
                        {
                            ConsultaPZ = new A7NET.Mensagem.udt.udtConsultaPZ();
                            ConsultaPZ.Parse(headerNZ.Substring(20));

                            CodigoMensagemNZ = string.Concat(ConsultaPZ.CodigoMensagem, " - ", ConsultaPZ.RCRotina);
                        }
                        else
                        {
                            STR0008R2 = new A7NET.Mensagem.udt.udtSTR0008R2();
                            STR0008R2.Parse(headerNZ.Substring(20));

                            CodigoMensagemNZ = STR0008R2.NU_CTRL_EXTE_RECB;
                        }

                        break;
                    #endregion

                    default:
                        break;

                }

                return CodigoMensagemNZ;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

    }
}
