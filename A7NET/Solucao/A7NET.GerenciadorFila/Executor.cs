using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text;
using System.Threading;
using A7NET.ConfiguracaoMQ;

namespace A7NET.GerenciadorFila
{
    public delegate bool AsyncDelegate(Fila fila, out Fila fila1);

    public class Executor : IDisposable
    {

        #region <<< Constructors Members >>>
        public Executor()
        {

        }
        #endregion

        #region <<< IDisposable Members >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~Executor()
        {
            this.Dispose();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region <<< NETTaskExecutor Members >>>
        // Executa Processos .Net
        public bool NETTaskExecutor(Fila fila, out Fila fila1)
        {
            try
            {

                int AvailableThreads = 0;
                int OutComp = 0;
                bool Processar = true;

                ThreadPool.GetAvailableThreads(out AvailableThreads, out OutComp);
                Thread.CurrentThread.Name = fila.NomeFila;

                if (AvailableThreads == 0)
                {
                    Processar = false;

                }
                else
                {

                    if (fila.QuantidadeAtualThread <= fila.QuantidadeMaxThreads)
                    {
                        Processar = true;
                    }
                    else
                    {
                        Processar = false;
                    }
                }

                if (Processar)
                {
                    using (A7NET.GerenciadorFila.GerenciadorRecebimento Gerenciador = new A7NET.GerenciadorFila.GerenciadorRecebimento())
                    {
                        Gerenciador.ProcessaMensagemMQ(fila.NomeFila, fila.QuantidadePorThread.ToString());
                    }

                }

            }
            catch (Exception ex)
            {

                System.Diagnostics.EventLog.WriteEntry("NETTaskExecutor", "ERR-NETTaskExecutor:" + ex, EventLogEntryType.Error);
            }


            fila1 = fila;

            return true;
        }
        #endregion

    }
}
