using System;
using System.Diagnostics;

namespace br.santander.SLCC 
{
    public class SLCCFila:IDisposable
    {
        
        private string _nomeFila=null;
        private string _nomeObjeto = null;
        private string _Metodo = null;

        private long   _quantidadeMaxThreads=0;
        private long   _quantidadeAtualThread=0;
        private string _xmlNodeFila = null;
        private ThreadPriorityLevel _Priority = ThreadPriorityLevel.Normal;
        private DateTime _DataHoraUltiExec = DateTime.Now;

        #region Consttutores Members
            public SLCCFila(){

            }
            
            public SLCCFila(string nomeFila,
                        string nomeObjeto,
                        string Metodo,
                        long quantidadeMaxThreads,
                        long quantidadeAtualThread,
                        string xmlNodeFila,
                        string Priority)
            {
                this._nomeFila = nomeFila.Trim();
                this._nomeObjeto = nomeObjeto.Trim();
                this._Metodo = Metodo;
                this._quantidadeAtualThread = quantidadeAtualThread;
                this._quantidadeMaxThreads = quantidadeMaxThreads;
                this._xmlNodeFila = xmlNodeFila.Trim();
                this._DataHoraUltiExec = DateTime.Parse("01/01/1900 00:00:00");
                
                switch (Priority.ToString())
                {
                    case "AboveNormal":
                        _Priority = ThreadPriorityLevel.AboveNormal;
                        break;
                    case "BelowNormal":
                        _Priority = ThreadPriorityLevel.BelowNormal;
                        break;
                    case "Highest":
                        _Priority = ThreadPriorityLevel.Highest;
                        break;
                    case "Idle":
                        _Priority = ThreadPriorityLevel.Idle;
                        break;
                    case "Lowest":
                        _Priority = ThreadPriorityLevel.Lowest;
                        break;
                    case "Normal":
                        _Priority = ThreadPriorityLevel.Normal;
                        break;
                    case "TimeCritical":
                        _Priority = ThreadPriorityLevel.TimeCritical;
                        break;
                }

            }

        #endregion
        
        #region IDisposable Members
            

            public void Dispose()
            {
                GC.SuppressFinalize(this);
            }

            ~SLCCFila()
            {
                _nomeFila=null;
                _nomeObjeto = null;

                this.Dispose();
            }

        #endregion

        #region Getter / Setter Members

            public DateTime DataHoraUltiExec
            {
                get { return _DataHoraUltiExec; }
                set { _DataHoraUltiExec = value; }
            }

            public ThreadPriorityLevel Priority
            {
                get { return _Priority; }
                set { _Priority = value; }
            }

            public string Metodo
            {
                get { return _Metodo; }
                set { _Metodo = value; }
            }

            public string nomeFila
            {
                get { return _nomeFila; }
                set { _nomeFila = value; }
            }

            public string nomeObjeto
            {
                get { return _nomeObjeto; }
                set { _nomeObjeto = value; }
            }

            public long quantidadeMaxThreads
            {
                get { return _quantidadeMaxThreads; }
                set { _quantidadeMaxThreads = value; }
            }

            public long quantidadeAtualThread
            {
                get { return _quantidadeAtualThread; }
                set { _quantidadeAtualThread = value; }
            }

            public string xmlNodeFila
            {
                get { return _xmlNodeFila; }
                set { _xmlNodeFila = value; }
            }


        #endregion
    }
}
