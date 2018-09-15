using System;
using System.Collections.Generic;
using System.Collections;
using System.Xml;
using System.Xml.XPath;
using System.Threading;
using System.Diagnostics;
using System.Text;
using IBM.WMQ;

namespace br.santander.SLCC
{
    public class SLCCAsyncMain
    {
        private static ArrayList alFilas = null;
        private static ArrayList alFilasSynchronized = null;
        private static SLCCFilas colFilas = null;
        private static SLCCExecutor slccTasxkExecutor = null;
        private static AsyncDelegate asyncDelegate = null;
        private static MQQueueManager queueManager = null;

        private static string _DataHoraInicio = null;
        private static string _DataHoraFim = null;

        private static string _DataHoraInicioBackup = null;
        private static string _DataHoraFimBackup = null;

        private static string _QueueManagerName = null;

        #region Consttutores Members
            
            public SLCCAsyncMain()
            {

            }

        #endregion

        
        #region IDisposable Members

            public void Dispose()
            {
                GC.SuppressFinalize(this);
            }

            ~SLCCAsyncMain()
            {
                slccTasxkExecutor.Dispose();
                slccTasxkExecutor = null;

                colFilas.Dispose();
                colFilas = null;

                asyncDelegate = null;
                alFilas = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                this.Dispose();
            }

        #endregion
        


        public void Carregar() 
        {
            try
            {
                int maxThreads = 0;
                int outComp = 0;

                //CARREGAR CONFIG DE PROCESSAMENTO DAS FILAS    
                string xmlFilas = AppDomain.CurrentDomain.BaseDirectory + "\\FilasNet.xml";
                XmlDocument xmlDocConfig = new XmlDocument();
                xmlDocConfig.Load(xmlFilas);

                _QueueManagerName = xmlDocConfig.SelectSingleNode("//NomeQueueManager").InnerText.Trim();

                _DataHoraInicio = xmlDocConfig.SelectSingleNode("//Janela_ParadaVerificacao/@HoraInicio").InnerText.Trim();
                _DataHoraFim = xmlDocConfig.SelectSingleNode("//Janela_ParadaVerificacao/@HoraFim").InnerText.Trim();

                if (xmlDocConfig.SelectSingleNode("//Janela_Backup") != null)
                {
                    _DataHoraInicioBackup = xmlDocConfig.SelectSingleNode("//Janela_Backup/@HoraInicio").InnerText.Trim();
                    _DataHoraFimBackup = xmlDocConfig.SelectSingleNode("//Janela_Backup/@HoraFim").InnerText.Trim();
                }
                else
                {
                    _DataHoraInicioBackup = "2130";
                    _DataHoraFimBackup = "2355";
                }

                ThreadPool.GetAvailableThreads(out maxThreads, out  outComp);
                maxThreads = Convert.ToInt16(xmlDocConfig.SelectSingleNode("//QuantidadeTotalThreads").InnerText.ToString()) + 2;
                ThreadPool.SetMaxThreads(maxThreads, outComp);

                alFilas = new ArrayList();

                //Carregar Lista de Filas para processamento
                foreach (XmlNode xmlNode in xmlDocConfig.SelectNodes("//Grupo_Parametros_Entrada"))
                {
                    if (Convert.ToInt16(xmlNode.SelectSingleNode("QtdeMaxThread").InnerText.ToString()) > 0)
                    {
                        SLCCFila itemFila = new SLCCFila(
                                                        xmlNode.SelectSingleNode("Fila").InnerText.ToString(),
                                                        xmlNode.SelectSingleNode("ObjetoApoio").InnerText.ToString(),
                                                        xmlNode.SelectSingleNode("Metodo").InnerText.ToString(),
                                                        Convert.ToInt16(xmlNode.SelectSingleNode("QtdeMaxThread").InnerText.ToString()),
                                                        0,
                                                        xmlNode.OuterXml,
                                                        xmlNode.SelectSingleNode("Priority").InnerText.ToString()
                                                        );
                        alFilas.Add(itemFila);
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRO->>>>>>Inicializar()" + ex.Message);
                throw ex;
            }
            finally
            {
                //Synchronized = ArrayLista Thread-Safe
                alFilasSynchronized = ArrayList.Synchronized(alFilas);

                colFilas = new SLCCFilas();
                colFilas.colFilas = alFilasSynchronized;
            }

        }


        [MTAThread]
        public void MainTask()
        {
            try
            {
                if (!VerificaJanelaProcessamento())
                {
                    return;
                }
               
                // Para cada fila na coleção  - Invoke COM+ Component
                foreach (SLCCFila itemFila in alFilasSynchronized)
                {
                    if (!ExisteMensagemFila(itemFila.nomeFila))
                    {
                        continue;
                    }

                    //Dados da fila
                    SLCCFila auxFila = itemFila;
                    
                    //Invoke COM+ component
                    slccTasxkExecutor = new SLCCExecutor();

                    //Async call  - trata retorno de cada thread
                    asyncDelegate = new AsyncDelegate(slccTasxkExecutor.COMTaskExecutor);
                    
                    //Adiciona 1 a quantidade de threads executada
                    colFilas.AddQt(itemFila.nomeFila);

                    //Thread Invoke COM+ Component - 
                    IAsyncResult asyncResult = asyncDelegate.BeginInvoke(itemFila,
                                                                         out auxFila,
                                                                         new AsyncCallback(CallbackMethod),
                                                                         asyncDelegate);

                }


                Thread.Sleep(3000);
                
                
                foreach (SLCCFila itemFila in alFilasSynchronized)
                {
                    if (!ExisteMensagemFila(itemFila.nomeFila))
                    {
                        continue;
                    }

                    //Dados da fila
                    SLCCFila auxFila = itemFila;

                    if (!itemFila.nomeFila.Contains("A8B."))
                    {
                        //Invoke COM+ component
                        slccTasxkExecutor = new SLCCExecutor();

                        //Async call  - trata retorno de cada thread
                        asyncDelegate = new AsyncDelegate(slccTasxkExecutor.COMTaskExecutor);

                        //Adiciona 1 a quantidade de threads executada
                        colFilas.AddQt(itemFila.nomeFila);

                        //Thread Invoke COM+ Component - 
                        IAsyncResult asyncResult = asyncDelegate.BeginInvoke(itemFila,
                                                                             out auxFila,
                                                                             new AsyncCallback(CallbackMethod),
                                                                             asyncDelegate);
                    }
                }



            }catch (Exception ex){

                slccTasxkExecutor = null;
                asyncDelegate = null;

                Console.WriteLine("ERRO->>>>>>MainTask()" + ex.Message);
                throw ex;
                
            }
        }

        // Trata retorno de cada Thread - Invoke COM+ component
        [MTAThread]
        private static void CallbackMethod(IAsyncResult asyncResult)
        {

            SLCCFila itemFila = null;

            try
            {
                AsyncDelegate asyncDelegate = (AsyncDelegate)asyncResult.AsyncState;

                bool ret = asyncDelegate.EndInvoke(out itemFila, asyncResult);
                
                colFilas.RemoveQt(itemFila.nomeFila);

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRO->>>>>>CallbackMethod()" + ex.Message);
                throw ex;

            }

        }

        public bool AllThreadsFinalized()
        {
            bool allThreadsFinalized = true;

            try
            {
                foreach (SLCCFila item in colFilas.colFilas)
                {
                    if (item.quantidadeAtualThread != 0)
                    {
                        allThreadsFinalized = false;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRO->>>>>>AllThreadsFinalized()" + ex.Message);
                throw ex;

            }
            return allThreadsFinalized;
        }

        private static bool VerificaJanelaProcessamento()
        {

            bool Processar = false;
            DateTime DataHoraAtual = DateTime.Now;
            DateTime DataHoraInicio = DateTime.Now ;
            DateTime DataHoraFim = DateTime.Now;
            
            string Inicio = null;
            string Fim = null;

            
            

            try
            {

                if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                {
                    StringBuilder sbInicio = new StringBuilder();

                    sbInicio.Append(DataHoraAtual.ToShortDateString());
                    sbInicio.Append(" ");
                    sbInicio.Append(_DataHoraInicioBackup.Substring(0,2));
                    sbInicio.Append(":");
                    sbInicio.Append(_DataHoraInicioBackup.Substring(2,2));
                    sbInicio.Append(":");
                    sbInicio.Append("00");

                    Inicio = sbInicio.ToString();
                    
                    StringBuilder sbFim = new StringBuilder();

                    sbFim.Append(DataHoraAtual.ToShortDateString());
                    sbFim.Append(" ");
                    sbFim.Append(_DataHoraFimBackup.Substring(0,2));
                    sbFim.Append(":");
                    sbFim.Append(_DataHoraFimBackup.Substring(2,2));
                    sbFim.Append(":");
                    sbFim.Append("00");

                    Fim = sbFim.ToString();

                    DataHoraInicio = DateTime.Parse(Inicio);
                    DataHoraFim = DateTime.Parse(Fim);

                    if ((DataHoraAtual >= DataHoraInicio) && (DataHoraAtual <= DataHoraFim))
                    {
                        Processar = false;
                    }
                    else
                    {
                        Processar = true;
                    }

                }
                else
                {

                    StringBuilder sbInicio = new StringBuilder();

                    sbInicio.Append(DataHoraAtual.ToShortDateString());
                    sbInicio.Append(" ");
                    sbInicio.Append(_DataHoraInicio.Substring(0,2));
                    sbInicio.Append(":");
                    sbInicio.Append(_DataHoraInicio.Substring(2,2));
                    sbInicio.Append(":");
                    sbInicio.Append("00");

                    Inicio = sbInicio.ToString();
                    
                    StringBuilder sbFim = new StringBuilder();

                    sbFim.Append(DataHoraAtual.ToShortDateString());
                    sbFim.Append(" ");
                    sbFim.Append(_DataHoraFim.Substring(0,2));
                    sbFim.Append(":");
                    sbFim.Append(_DataHoraFim.Substring(2,2));
                    sbFim.Append(":");
                    sbFim.Append("00");

                    Fim = sbFim.ToString();
                    
                    DataHoraInicio = DateTime.Parse(Inicio);
                    DataHoraFim = DateTime.Parse(Fim);

                    if ((DataHoraAtual >= DataHoraInicio) && (DataHoraAtual <= DataHoraFim))
                    {
                        Processar = false;
                    }
                    else
                    {
                        Processar = true;
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRO->>>>>>VerificaJanelaProcessamento()" + ex.Message);
                Processar = true;
            }
            finally
            {
                Inicio = null;
                Fim = null;
            }

            return Processar;
            
        }
        
        private static bool ExisteMensagemFila(string nomeFila)
        {

            bool ret = false;
            
            MQQueue queue = null;
            
            try
            {
                if (nomeFila.Trim().ToUpper().Contains("A8B.")){
                    
                    ret = true;
                
                }else{
                    if (queueManager == null){
                        queueManager = new MQQueueManager(_QueueManagerName);
                    }

                    if (!queueManager.IsConnected){
                        queueManager = new MQQueueManager(_QueueManagerName);
                    }

                    queue = queueManager.AccessQueue(nomeFila, MQC.MQOO_INQUIRE + MQC.MQQT_LOCAL);

                    if (queue.CurrentDepth > 0){
                        ret = true;
                    }else{
                        ret = false;
                    }
                }
            
            }catch (Exception ex){

                Console.WriteLine("ERRO->>>>>>ExisteMensagemFila()" + DateTime.Now + "-"+ ex);
                ret = false;
            }

            if (queue!=null){
                queue.Close();
                queue = null;
            }

            return ret;
        }
    }

}
