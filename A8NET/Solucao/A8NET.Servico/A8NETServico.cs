using System;
using System.Collections.Generic;
using System.Data;
using System.ServiceProcess;
using System.Threading;
using System.Xml;
using System.IO;
using System.Text;
using System.Diagnostics;
using A8NET.GerenciadorFila;
using System.Globalization;


namespace A8NET.Servico
{
    public partial class A8NETServico : ServiceBase
    {
        A8NET.GerenciadorFila.VerificaFila verificaFilas = null;

        // Flag para indicar o status do serviço.
        private bool serviceStarted = false;
        private int TempoEspera = 2; // em segundos
        private Thread workerThread = null;

        public A8NETServico()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                #if DEBUG
                    System.Diagnostics.Debugger.Break();
                #endif

                //Mensagem.RecebimentoMensagemSPB aa = new RecebimentoMensagemSPB();
                //aa.ProcessaMensagem("aaa", "<MESG><TP_MESG>000001000</TP_MESG><SG_SIST_ORIG>NZ</SG_SIST_ORIG><SG_SIST_DEST>A8</SG_SIST_DEST><CO_EMPR>558</CO_EMPR><TX_MESG>A8 SEL1052R120110302A8 20409926420100820005580079029040088800038121CON12345                          A8 A8558188000000000000000000001904008882010082004124060000000000000000000000000000000000000000000 <SISMSG><SEL1052R1><CodMsg>SEL1052R1</CodMsg><NumCtrlIF>20110302A8 204099253</NumCtrlIF><ISPBIF>90400888</ISPBIF><NumOpSEL>12345</NumOpSEL><SitOpSEL>CON</SitOpSEL><DtHrSit>20100820122843</DtHrSit><DtMovto>20100820</DtMovto></SEL1052R1></SISMSG></TX_MESG></MESG>");

                #region >>> Setar a Cultura >>>
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
                #endregion

                #region >>> Tratamentos das Regras de Mensageria >>>>
                // Carrega Configuracao de Processamento das Filas    
                string xmlFilas = AppDomain.CurrentDomain.BaseDirectory + "\\A8NETFilas.xml";
                XmlDocument xmlDocConfig = new XmlDocument();
                xmlDocConfig.Load(xmlFilas);
                if (xmlDocConfig.SelectSingleNode("//TempoEspera") != null)
                {
                    TempoEspera = Convert.ToInt16(xmlDocConfig.SelectSingleNode("//TempoEspera").InnerText);
                }
                xmlDocConfig = null;
                xmlFilas = null;

                verificaFilas = new A8NET.GerenciadorFila.VerificaFila();

                verificaFilas.Carregar();

                ThreadStart st = new ThreadStart(WorkerFunction);
                workerThread = new Thread(st);

                // Indica que o processo foi inicializado
                serviceStarted = true;

                // Inicia a thread.
                workerThread.Start();

                #endregion

            }
            catch (Exception)
            {
                //System.Diagnostics.EventLog.WriteEntry("A8NETServico:OnStart()",
                //                                       "ERRO: " + ex.Message,
                //                                       EventLogEntryType.Error);

            }
        }

        protected override void OnStop()
        {
            try
            {
                if (verificaFilas != null)
                {
                    // Indica que o processo foi finalizado
                    serviceStarted = false;

                    // Espera 5 segundos antes da finalizacao

                    int esperaFinalizao = 0;
                    bool finalizar = false;

                    while (finalizar)
                    {
                        finalizar = verificaFilas.AllThreadsFinalized();
                        Thread.Sleep(new TimeSpan(0, 0, 1)); // Espera 1 segundo
                        esperaFinalizao++;

                        if (esperaFinalizao > 10) // Espera 10 segundos antes de terminar servico
                        {
                            if (!finalizar) finalizar = true;
                        }
                        else
                        {
                            finalizar = false;
                        }
                    }

                    verificaFilas.Dispose();
                    verificaFilas = null;

                }

            }
            catch (Exception)
            {
                //System.Diagnostics.EventLog.WriteEntry("A8NETServico:OnStop()",
                //                                       "ERRO: " + ex.Message,
                //                                       EventLogEntryType.Error);
            }
            finally
            {
                workerThread.Abort();
            }
        }

        private void WorkerFunction()
        {
            try
            {
                while (serviceStarted)
                {
                    if (serviceStarted)
                    {
                        verificaFilas.MainTask();
                        Thread.Sleep(new TimeSpan(0, 0, TempoEspera));
                    }
                }
            }
            catch (Exception)
            {
                //System.Diagnostics.EventLog.WriteEntry("A8NETServico:WorkerFunction",
                //                                       "ERRO: " + ex.Message,
                //                                       EventLogEntryType.Error);

            }
        }

        //#region LogErro Members
        //public void GravaLogErro(string strErro, Exception ex, string remessaRecebida)
        //{

        //    if (strErro == null || strErro.Trim() == "")
        //    {
        //        return;
        //    }

        //    string pathRel = "";
        //    string nomeArquivo = "Erro_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + ".txt";

        //    pathRel = AppDomain.CurrentDomain.BaseDirectory + "LogErro\\";

        //    StringBuilder stringBuilder1 = new StringBuilder();
        //    stringBuilder1.AppendFormat(strErro);
        //    stringBuilder1.Append(Environment.NewLine);
        //    stringBuilder1.Append("MESSAGEM ERRO: ");
        //    stringBuilder1.Append(Environment.NewLine);
        //    stringBuilder1.Append(ex.ToString().ToUpper());
        //    stringBuilder1.Append(Environment.NewLine);
        //    stringBuilder1.Append(ex.Message.ToUpper());
        //    stringBuilder1.Append(Environment.NewLine);
        //    stringBuilder1.Append("MESSAGEM RECEBIDA: ");
        //    stringBuilder1.Append(Environment.NewLine);
        //    stringBuilder1.Append(remessaRecebida.ToString());
        //    stringBuilder1.Append(Environment.NewLine);


        //    try
        //    {

        //        if (strErro.Length > 0)
        //        {
        //            #region Tratamento dos arquivos no diretório
        //            // Verifica se existe o Path.
        //            if (Directory.Exists(pathRel))
        //            {
        //                // Máximo de arquivos permitido. Expurgo parametrizável.
        //                if (Directory.GetFiles(pathRel).Length >= 2)
        //                {
        //                    foreach (string sArquivo in Directory.GetFiles(pathRel))
        //                    {
        //                        TimeSpan dataDiff = DateTime.Now.Subtract(File.GetLastWriteTime(sArquivo).Date);

        //                        int diasDiff = dataDiff.Days;

        //                        if (diasDiff >= 1)
        //                        {
        //                            File.Delete(sArquivo);
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                Directory.CreateDirectory(@pathRel);
        //            }
        //            #endregion

        //            #region Grava ou adiciona registro no log

        //            StringBuilder sb = new StringBuilder();

        //            sb.Append("####################################");
        //            sb.Append(DateTime.Now.ToLongDateString());
        //            sb.Append(" - ");
        //            sb.Append(DateTime.Now.ToLongTimeString());
        //            sb.Append("####################################");
        //            sb.Append(Environment.NewLine);
        //            sb.Append(stringBuilder1.ToString());
        //            sb.Append(Environment.NewLine);
        //            sb.Append("______________________________________");


        //            if (File.Exists(pathRel + nomeArquivo))
        //            {
        //                using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Append, FileAccess.Write, FileShare.Write))
        //                {
        //                    using (StreamWriter sw = new StreamWriter(fs))
        //                    {

        //                        sw.Write(sb.ToString());
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Create, FileAccess.Write, FileShare.Write))
        //                {
        //                    using (StreamWriter sw = new StreamWriter(fs))
        //                    {
        //                        sw.Write(sb.ToString());
        //                    }
        //                }
        //            }
        //            #endregion
        //        }


        //    }
        //    catch (Exception)
        //    {

        //    }

        //}
        //#endregion

    }
}
