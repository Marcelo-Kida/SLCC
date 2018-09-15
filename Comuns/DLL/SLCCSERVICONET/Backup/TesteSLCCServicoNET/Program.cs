using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using br.santander.SLCC;

namespace TesteSLCCServicoNET
{
    class Program
    {

        private static bool serviceStarted = true;
        private static SLCCAsyncMain slccAsyncMain = null;
        private static Thread workerThread = null;

        static void Main(string[] args)
        {




            slccAsyncMain = new SLCCAsyncMain();

            slccAsyncMain.Carregar();

            ThreadStart st = new ThreadStart(WorkerFunction);
            workerThread = new Thread(st);

            // Inicia a thread.
            workerThread.Start();

        }

        private static void WorkerFunction()
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
                            Thread.Sleep(new TimeSpan(0, 20 , 0));
                        }
                        else
                        {
                            Thread.Sleep(new TimeSpan(0 , 0 , 10));
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(">>>>>>>> {0}", ex);
                
            }
            finally
            {
                // Finaliza a thread. 
                Thread.CurrentThread.Abort();

            }


        }



    }
}
