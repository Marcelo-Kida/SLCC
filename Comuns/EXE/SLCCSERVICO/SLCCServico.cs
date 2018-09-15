using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using br.santander.SLCC;

namespace SLCCServico
{
    public partial class SLCCServico : ServiceBase
    {
        SLCCAsyncMain slccAsyncMain = null;

        // Flag para indicar o status do serviço.
        private bool serviceStarted = false;
        private int TempoEspera = 10; // em segundos

        private Thread workerThread = null;

        private XmlNode _xmlNodeAgendamento = null;

        public SLCCServico()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                //System.Diagnostics.Debugger.Break();

                AtivaValidaRemessa(true);

                //CARREGAR CONFIG DE PROCESSAMENTO DAS FILAS    
                string xmlFilas = AppDomain.CurrentDomain.BaseDirectory + "\\FilasNet.xml";
                XmlDocument xmlDocConfig = new XmlDocument();
                xmlDocConfig.Load(xmlFilas);

                TempoEspera = Convert.ToInt16(xmlDocConfig.SelectSingleNode("//TempoEspera").InnerText.Trim()) ;

                slccAsyncMain = new SLCCAsyncMain();

                slccAsyncMain.Carregar();

                ThreadStart st = new ThreadStart(WorkerFunction);
                workerThread = new Thread(st);

                // Indica que o processo foi inicializado.
                serviceStarted = true;

                // Inicia a thread.
                workerThread.Start();

                //Inicializa a execução das threads de Agendamento
                ExecutaProcessosAgendamento(xmlDocConfig);
            }
            catch (Exception ex)
            {
                // Grava log de erro.
                GravaLogErro(ex.Message);

                // Retira objetos da memória.
                slccAsyncMain.Dispose();
                slccAsyncMain = null;
            }
        }

        protected override void OnStop()
        {
            //System.Diagnostics.Debugger.Break();

            AtivaValidaRemessa(false);

            try
            {

                if (slccAsyncMain != null)
                {
                    // Indica que o processo foi finalizado.
                    serviceStarted = false;

                    // Espera 5 segundos antes da finalização

                    int esperaFinalizao = 0;
                    bool finalidar = false;

                    while(finalidar) 
                    {
                        finalidar = slccAsyncMain.AllThreadsFinalized();
                        Thread.Sleep(new TimeSpan(0, 0, 1)); // Espera 1 segundo
                        esperaFinalizao++;

                        if (esperaFinalizao > 10) // Espera 10 segundos antes de terminar serviço
                        {
                            if (!finalidar) finalidar = true;
                        }
                        else 
                        {
                            finalidar = false;
                        }
                    }
                    
                    workerThread.Suspend();

                    slccAsyncMain.Dispose();

                    slccAsyncMain = null;

                    
                }

            }
            catch (Exception ex)
            {
                // Grava log de erro.
                GravaLogErro(ex.Message);
            }

        }

        private void ExecutaProcessosAgendamento(XmlDocument xmlDocConfig)
        {
            try 
	        {
                foreach (XmlNode xmlNode in xmlDocConfig.SelectNodes("//Grupo_Parametros_Agendamento"))
                {
                    _xmlNodeAgendamento = xmlNode;

                    ThreadStart st = new ThreadStart(WorkerFunctionAgendamento);
                    workerThread = new Thread(st);

                    // Indica que o processo foi inicializado.
                    serviceStarted = true;

                    // Inicia a thread.
                    workerThread.Start();
                }

	        }
            catch (Exception ex)
            {
                // Grava log de erro.
                GravaLogErro(ex.Message);
            }
        }

        private void WorkerFunctionAgendamento()
        {
            int intervaloExecucao;
            try
            {
                if (this._xmlNodeAgendamento != null)
                {
                    intervaloExecucao = Int32.Parse(_xmlNodeAgendamento.SelectSingleNode("IntervaloExecucaoMinutos").InnerText.ToString());
                    while (serviceStarted)
                    {
                        if (serviceStarted)
                        {
                            slccAsyncMain.MainTaskAgendamento(_xmlNodeAgendamento.SelectSingleNode("ObjetoApoio").InnerText.ToString(), _xmlNodeAgendamento.SelectSingleNode("Metodo").InnerText.ToString(), _xmlNodeAgendamento.OuterXml);

                            Thread.Sleep(new TimeSpan(0, intervaloExecucao, 0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Grava log de erro.
                GravaLogErro(ex.Message);
            }
            finally
            {
                // Finaliza a thread. 
                Thread.CurrentThread.Abort();

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
                        slccAsyncMain.MainTask();

                        if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday ||
                            DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Thread.Sleep(new TimeSpan(0,  10 , 0));
                        }
                        else
                        {
                            Thread.Sleep(new TimeSpan(0, 0, TempoEspera));
                        }

                        
                    }
                }
            }
            catch (Exception ex)
            {
                // Grava log de erro.
                GravaLogErro(ex.Message);
            }
            finally
            {
                // Finaliza a thread. 
                Thread.CurrentThread.Abort();

            }
            

        }

        private static void AtivaValidaRemessa(bool _start)
        {

            Process ativaValidaRemessa = null;
            string cmd = Environment.GetFolderPath(Environment.SpecialFolder.System).ToUpper() + "\\Net.exe";

            try
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = cmd;
                processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                if (_start)
                {
                    processStartInfo.Arguments = "START A6A8AtivaValidaRemessa";
                }
                else
                {
                    processStartInfo.Arguments = "STOP A6A8AtivaValidaRemessa";
                }

                ativaValidaRemessa = Process.Start(processStartInfo);
                ativaValidaRemessa.WaitForExit();
                ativaValidaRemessa.Close();

                ativaValidaRemessa.Dispose();
                ativaValidaRemessa = null;

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRO->>>>>>IniciaAtivaValidaRemessa()" + ex.Message);
                //throw ex;
            }
            finally
            {

                ativaValidaRemessa = null;
            }
        }


        private void GravaLogErro(string strErro)
        {
            string pathRel = "";
            string nomeArquivo = "LogErros" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + ".txt";

            pathRel = AppDomain.CurrentDomain.BaseDirectory + "LogErro\\";

            if (strErro.Length > 0)
            {
                #region Tratamento dos arquivos no diretório
                // Verifica se existe o Path.
                if (Directory.Exists(pathRel))
                {
                    // Máximo de arquivos permitido. Expurgo parametrizável.
                    if (Directory.GetFiles(pathRel).Length >= 2)
                    {
                        foreach (string sArquivo in Directory.GetFiles(pathRel))
                        {
                            TimeSpan dataDiff = DateTime.Now.Subtract(File.GetLastWriteTime(sArquivo).Date);

                            int diasDiff = dataDiff.Days;

                            if (diasDiff >= 2)
                            {
                                File.Delete(sArquivo);
                            }
                        }
                    }
                }
                else
                {
                    Directory.CreateDirectory(@pathRel);
                }
                #endregion

                #region Grava ou adiciona registro no log
                if (File.Exists(pathRel + nomeArquivo))
                {
                    using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Append, FileAccess.Write, FileShare.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            sw.Write(strErro);
                        }
                    }
                }
                else
                {
                    using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Create, FileAccess.Write, FileShare.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            sw.Write(strErro);
                        }
                    }
                }
                #endregion
            }
        }
    }
}
