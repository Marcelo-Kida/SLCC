using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace A7NET.Mensagem
{
    public class Repeticao
    {
        #region <<< Variaveis >>>
        
        #endregion

        #region >>> Construtor >>>
        public Repeticao()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion
        }
        #endregion

        #region <<< IncluiRepeticaRecursivo >>>
        public void IncluiRepeticaRecursivo(ref XmlDocument xmlRepeticaoRecursivo)
        {
            XmlDocument XmlRepeticao;
            StringBuilder MesgSaida = new StringBuilder();
            StringBuilder MesgAux = new StringBuilder();
            StringBuilder Repeticao = new StringBuilder();
            string FormatoInicio = "<{0}>";
            string FormatoFim = "</{0}>";

            try
            {
                MesgSaida.AppendFormat(FormatoInicio, "REPE" + xmlRepeticaoRecursivo.DocumentElement.Name.Substring(2));
                MesgSaida.AppendFormat(FormatoInicio, xmlRepeticaoRecursivo.DocumentElement.Name);

                foreach (XmlNode Node in xmlRepeticaoRecursivo.DocumentElement.ChildNodes)
                {
                    if (Node.Name.Substring(0, 3).ToUpper() != "GR_")
                    {
                        MesgSaida.Append(Node.OuterXml);
                    }
                    else
                    {
                        if (MesgSaida.ToString().IndexOf("</REPE" + Node.Name.Substring(2)) != -1)
                        {
                            Repeticao.Remove(0, Repeticao.Length);
                            Repeticao.Append(Node.OuterXml);
                            if (Repeticao.ToString().IndexOf("<GR_", 1) != -1)
                            {
                                XmlRepeticao = new XmlDocument();
                                xmlRepeticaoRecursivo.LoadXml(Repeticao.ToString());
                                IncluiRepeticaRecursivo(ref XmlRepeticao);
                                Repeticao.Remove(0, Repeticao.Length);
                                Repeticao.Append(XmlRepeticao.OuterXml);
                                MesgAux.Remove(0, MesgAux.Length);
                                MesgAux.Append(MesgSaida.ToString().Substring(0, (MesgSaida.ToString().IndexOf("</REPE" + Node.Name.Substring(2)) - 1)));
                                MesgAux.Append(Repeticao.ToString().Substring(0, Repeticao.ToString().IndexOf(Node.Name) - 1));
                                MesgSaida.Remove(0, MesgSaida.Length);
                                MesgSaida.Append(MesgAux.ToString());
                                XmlRepeticao = null;
                            }
                            else
                            {
                                MesgAux.Remove(0, MesgAux.Length);
                                MesgAux.Append(MesgSaida.ToString().Substring(0, (MesgSaida.ToString().IndexOf("</REPE" + Node.Name.Substring(2)) - 1)));
                                MesgAux.Append(Repeticao.ToString());
                                MesgAux.Append(MesgSaida.ToString().Substring(MesgSaida.ToString().IndexOf("</REPE" + Node.Name.Substring(2))));
                                MesgSaida.Remove(0, MesgSaida.Length);
                                MesgSaida.Append(MesgAux.ToString());
                            }
                        }
                        else
                        {
                            XmlRepeticao = new XmlDocument();
                            Repeticao.Remove(0, Repeticao.Length);
                            Repeticao.Append(Node.OuterXml);
                            XmlRepeticao.LoadXml(Repeticao.ToString());
                            IncluiRepeticaRecursivo(ref XmlRepeticao);
                            MesgSaida.Append(XmlRepeticao.OuterXml);
                            XmlRepeticao = null;
                        }
                    }
                }

                MesgSaida.AppendFormat(FormatoFim, xmlRepeticaoRecursivo.DocumentElement.Name);
                MesgSaida.AppendFormat(FormatoFim, "REPE" + xmlRepeticaoRecursivo.DocumentElement.Name.Substring(2));

                xmlRepeticaoRecursivo.LoadXml(MesgSaida.ToString());

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< RetiraRepeticaRecursivo >>>
        public void RetiraRepeticaRecursivo(ref XmlDocument xmlRepeticaoRecursivo)
        {
            XmlDocument XmlRepeticao;
            StringBuilder MesgSaida = new StringBuilder();
            StringBuilder MesgAux = new StringBuilder();
            StringBuilder Repeticao = new StringBuilder();
            string FormatoInicio = "<{0}>";
            string FormatoFim = "</{0}>";

            try
            {
                MesgSaida.AppendFormat("<REPET>");

                foreach (XmlNode Node in xmlRepeticaoRecursivo.DocumentElement.ChildNodes)
                {
                    MesgAux.Append(Node.OuterXml);

                    if (MesgAux.ToString().ToUpper().IndexOf("REPE_" + Node.Name.Substring(2)) != -1)
                    {
                        MesgSaida.Append(Node.OuterXml);
                    }
                    else
                    {
                        MesgSaida.AppendFormat(FormatoInicio, Node.Name);
                        foreach (XmlNode NodeGrupo in Node.ChildNodes)
                        {
                            if (NodeGrupo.Name.Substring(0, 5).ToUpper() != "REPE_")
                            {
                                MesgSaida.Append(NodeGrupo.OuterXml);
                            }
                            else
                            {
                                XmlRepeticao = new XmlDocument();
                                Repeticao.Remove(0, Repeticao.Length);
                                Repeticao.Append(NodeGrupo.OuterXml);
                                XmlRepeticao.LoadXml(Repeticao.ToString());
                                RetiraRepeticaRecursivo(ref XmlRepeticao);
                                MesgSaida.Append(XmlRepeticao.OuterXml);
                                XmlRepeticao = null;
                            }
                        }
                        MesgSaida.AppendFormat(FormatoFim, Node.Name);
                    }
                }

                MesgSaida.AppendFormat("</REPET>");

                xmlRepeticaoRecursivo.LoadXml(MesgSaida.ToString());

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #region <<< IncluiGrupoRecursivo >>>
        public void IncluiGrupoRecursivo(ref XmlDocument xmlGrupoRecursivo)
        {
            StringBuilder MesgSaida = new StringBuilder();
            string FormatoInicio = "<{0}>";
            string FormatoFim = "</{0}>";

            try
            {
                MesgSaida.AppendFormat(FormatoInicio, xmlGrupoRecursivo.DocumentElement.Name);

                foreach (XmlNode Node in xmlGrupoRecursivo.DocumentElement.ChildNodes)
                {
                    MesgSaida.AppendFormat(FormatoInicio, "Grupo" + xmlGrupoRecursivo.DocumentElement.Name.Substring(5));
                    MesgSaida.Append(Node.OuterXml);
                    MesgSaida.AppendFormat(FormatoFim, "Grupo" + xmlGrupoRecursivo.DocumentElement.Name.Substring(5));
                }

                MesgSaida.AppendFormat(FormatoFim, xmlGrupoRecursivo.DocumentElement.Name);

                xmlGrupoRecursivo.LoadXml(MesgSaida.ToString());

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

    }
}
