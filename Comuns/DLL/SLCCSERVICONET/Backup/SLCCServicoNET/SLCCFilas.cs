using System;
using System.Collections;

namespace br.santander.SLCC
{
    public class SLCCFilas:IDisposable
    {
        private ArrayList _Filas;

        #region Consttutores Members
            public SLCCFilas(){

            }
        #endregion
        
        #region IDisposable Members
            
            public void Dispose()
            {
                GC.SuppressFinalize(this);
            }

            ~SLCCFilas()
            {
                _Filas = null;
                this.Dispose();
            }

        #endregion
        
        #region Getter / Setter Members
            public ArrayList colFilas
            {
                get { return _Filas; }
                set { _Filas = value; }
            }
        #endregion

        #region Methods Members
            public SLCCFila GetItem(string nomeFila)
            {
                foreach (SLCCFila item in _Filas)
                {
                    if (item.nomeFila == nomeFila)
                    {
                        return item;
                    }
                }

                return null;
            }
        
            public bool AddQt(string nomeFila)
            {
                SLCCFila item = GetItem(nomeFila);

                if (item == null)
                {
                    return false;
                }else{
                    item.quantidadeAtualThread = item.quantidadeAtualThread + 1;
                    return true;
                }
            }

            public bool RemoveQt(string nomeFila)
            {
                SLCCFila item = GetItem(nomeFila);

                if (item == null)
                {
                    return false;
                }
                else
                {
                    item.quantidadeAtualThread = item.quantidadeAtualThread - 1;
                    item.DataHoraUltiExec = DateTime.Now;
                    return true;
                }
            }

        #endregion

    }
}
