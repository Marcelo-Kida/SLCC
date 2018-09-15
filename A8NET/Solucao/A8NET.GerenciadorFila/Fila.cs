using System;
using System.Collections.Generic;
using System.Text;

namespace A8NET.GerenciadorFila
{
    public class Fila : IDisposable
    {

        #region <<< Variables >>>
        private string _NomeFila = null;
        private string _Metodo = null;
        private string _NomeObjeto = null;
        private string _TipoObjeto = null;
        private long _QuantidadeMaxThreads = 0;
        private long _QuantidadeAtualThread = 0;
        private int _QuantidadePorThread = 10;
        #endregion

        #region <<< Constructors Members >>>
        public Fila()
        {

        }

        public Fila(string nomeFila,
                    string nomeObjeto,
                    string metodo,
                    string tipoObjeto,
                    long quantidadeMaxThreads,
                    long quantidadeAtualThread,
                    int quantidadePorThread)
        {
            this._NomeFila = nomeFila.Trim();
            this._NomeObjeto = nomeObjeto.Trim();
            this._Metodo = metodo;
            this._TipoObjeto = tipoObjeto;
            this._QuantidadeAtualThread = quantidadeAtualThread;
            this._QuantidadeMaxThreads = quantidadeMaxThreads;
            this._QuantidadePorThread = quantidadePorThread;
        }
        #endregion

        #region <<< IDisposable Members >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~Fila()
        {
            _NomeFila = null;
            _NomeObjeto = null;

            this.Dispose();
        }
        #endregion

        #region <<< Get / Set Members >>>
        public int QuantidadePorThread
        {
            get { return _QuantidadePorThread; }
            set { _QuantidadePorThread = value; }
        }

        public string NomeFila
        {
            get { return _NomeFila; }
            set { _NomeFila = value; }
        }

        public string NomeObjeto
        {
            get { return _NomeObjeto; }
            set { _NomeObjeto = value; }
        }

        public string Metodo
        {
            get { return _Metodo; }
            set { _Metodo = value; }
        }

        public long QuantidadeMaxThreads
        {
            get { return _QuantidadeMaxThreads; }
            set { _QuantidadeMaxThreads = value; }
        }

        public long QuantidadeAtualThread
        {
            get { return _QuantidadeAtualThread; }
            set { _QuantidadeAtualThread = value; }
        }

        public string TipoObjeto
        {
            get { return _TipoObjeto; }
            set { _TipoObjeto = value; }
        }
        #endregion

    }
}
