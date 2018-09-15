using System;
using System.Collections;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Xml;
using IBM.WMQ;
using A7NET.ConfiguracaoMQ;

namespace A7NET.GerenciadorFila
{
    public class VerificaFila : IDisposable
    {
        #region <<< Variables >>>
        private static ArrayList _Filas = null;
        private static ArrayList _FilasSynchronized = null;
        private static Filas _ColFilas = null;
        private static Executor _TaskExecutor = null;
        private static AsyncDelegate _AsyncDelegate = null;
        private static MQQueueManager _QueueManager = null;
        private static string _QueueManagerName = null;
        private static int _QuantidadePorThread = 10;
        #endregion

        #region <<< Constructors Members >>>
        public VerificaFila()
        {

        }
        #endregion

        #region <<< IDisposable Members >>>

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~VerificaFila()
        {
            this.Dispose();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region <<< Carregar >>>
        public void Carregar()
        {
            try
            {
                int MaxThreads = 0;
                int OutComp = 0;

                // Carregar config de processamento das filas
                string XmlFilas = AppDomain.CurrentDomain.BaseDirectory + "\\A7NETFilas.xml";
                XmlDocument xmlDocConfig = new XmlDocument();
                xmlDocConfig.Load(XmlFilas);

                _QueueManagerName = xmlDocConfig.SelectSingleNode("//NomeQueueManager").InnerText.Trim();

                ThreadPool.GetAvailableThreads(out MaxThreads, out OutComp);
                MaxThreads = Convert.ToInt16(xmlDocConfig.SelectSingleNode("//QuantidadeTotalThreads").InnerText.ToString()) + 2;
                ThreadPool.SetMaxThreads(MaxThreads, OutComp);

                if (xmlDocConfig.SelectSingleNode("//QuantidadePorThread") != null)
                {
                    _QuantidadePorThread = Convert.ToInt16(xmlDocConfig.SelectSingleNode("//QuantidadePorThread").InnerText.ToString());
                }

                _Filas = new ArrayList();

                // Carregar Lista de Filas para processamento
                foreach (XmlNode xmlNode in xmlDocConfig.SelectNodes("//Grupo_Parametros_Entrada"))
                {
                    if (Convert.ToInt16(xmlNode.SelectSingleNode("QtdeMaxThread").InnerText.ToString()) > 0)
                    {
                        Fila ItemFila = new Fila(
                                                 xmlNode.SelectSingleNode("Fila").InnerText.ToString(),
                                                 xmlNode.SelectSingleNode("NomeObjeto").InnerText.ToString(),
                                                 xmlNode.SelectSingleNode("Metodo").InnerText.ToString(),
                                                 xmlNode.SelectSingleNode("Tipo").InnerText.ToString(),
                                                 Convert.ToInt16(xmlNode.SelectSingleNode("QtdeMaxThread").InnerText.ToString()),
                                                 0,
                                                 _QuantidadePorThread
                                                 );
                        _Filas.Add(ItemFila);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                _FilasSynchronized = ArrayList.Synchronized(_Filas);

                _ColFilas = new Filas();
                _ColFilas.ColFilas = _FilasSynchronized;
            }

        }
        #endregion

        #region <<< MainTask >>>
        [MTAThread]
        public void MainTask()
        {
            try
            {
                // Para cada fila na coleção  - Invoke COM+ Component
                foreach (Fila ItemFila in _FilasSynchronized)
                {
                    if (!ExisteMensagemFila(ItemFila.NomeFila)) { continue; }

                    //Dados da fila
                    Fila AuxFila = ItemFila;

                    //Invoke component
                    _TaskExecutor = new Executor();

                    //Async call  - trata retorno de cada thread
                    if (ItemFila.TipoObjeto.Equals("NET"))
                    {
                        _AsyncDelegate = new AsyncDelegate(_TaskExecutor.NETTaskExecutor);
                    }

                    //Adiciona 1 a quantidade de threads executada
                    _ColFilas.AddQt(ItemFila.NomeFila);

                    //Thread Invoke Component - 
                    IAsyncResult asyncResult = _AsyncDelegate.BeginInvoke(ItemFila,
                                                                          out AuxFila,
                                                                          new AsyncCallback(CallbackMethod),
                                                                          _AsyncDelegate);

                }

            }
            catch (Exception ex)
            {
                _TaskExecutor = null;
                _AsyncDelegate = null;
                throw ex;

            }
        }
        #endregion

        #region <<< CallBack >>>
        [MTAThread]
        private static void CallbackMethod(IAsyncResult asyncResult)
        {
            Fila ItemFila = null;
            try
            {
                AsyncDelegate asyncDelegate = (AsyncDelegate)asyncResult.AsyncState;
                bool retorno = asyncDelegate.EndInvoke(out ItemFila, asyncResult);
                _ColFilas.RemoveQt(ItemFila.NomeFila);

            }
            catch (Exception ex)
            {
                throw ex;

            }
        }
        #endregion

        #region <<< AllThreadsFinalized >>>
        public bool AllThreadsFinalized()
        {
            bool AllThreadsFinalized = true;

            try
            {
                foreach (Fila Item in _ColFilas.ColFilas)
                {
                    if (Item.QuantidadeAtualThread != 0)
                    {
                        AllThreadsFinalized = false;
                        break;
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;

            }

            return AllThreadsFinalized;

        }
        #endregion

        #region <<< ExisteMensagemFila >>>
        private static bool ExisteMensagemFila(string nomeFila)
        {

            bool Retorno = false;

            MQQueue Queue = null;

            try
            {

                if (_QueueManager == null)
                {
                    _QueueManager = new MQQueueManager(_QueueManagerName);
                }

                if (!_QueueManager.IsConnected)
                {
                    _QueueManager = new MQQueueManager(_QueueManagerName);
                }


                Queue = _QueueManager.AccessQueue(nomeFila, MQC.MQOO_INQUIRE + MQC.MQQT_LOCAL);

                if (Queue.CurrentDepth > 0)
                {
                    Retorno = true;
                }
                else
                {
                    Retorno = false;
                }

            }
            catch (Exception)
            {

                Retorno = false;
            }

            if (Queue != null)
            {
                Queue.Close();
                Queue = null;
            }

            return Retorno;
        }
        #endregion
    }
}
