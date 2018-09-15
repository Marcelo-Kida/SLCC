using System;
using System.Collections;
using System.Text;

namespace A7NET.GerenciadorFila
{
    public class Filas : IDisposable
    {

        #region <<< Variables >>>
        private ArrayList _Filas;
        #endregion

        #region <<< Constructors Members >>>
        public Filas()
        {

        }
        #endregion

        #region <<< IDisposable Members >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~Filas()
        {
            _Filas = null;
            this.Dispose();
        }
        #endregion

        #region <<< Get / Set Members >>>
        public ArrayList ColFilas
        {
            get { return _Filas; }
            set { _Filas = value; }
        }
        #endregion

        #region <<< Methods Members >>>
        public Fila GetItem(string nomeFila)
        {
            foreach (Fila Item in _Filas)
            {
                if (Item.NomeFila == nomeFila)
                {
                    return Item;
                }
            }

            return null;
        }

        public bool AddQt(string nomeFila)
        {
            Fila Item = GetItem(nomeFila);

            if (Item == null)
            {
                return false;
            }
            else
            {
                Item.QuantidadeAtualThread = Item.QuantidadeAtualThread + 1;
                return true;
            }
        }

        public bool RemoveQt(string nomeFila)
        {
            Fila Item = GetItem(nomeFila);

            if (Item == null)
            {
                return false;
            }
            else
            {
                Item.QuantidadeAtualThread = Item.QuantidadeAtualThread - 1;
                return true;
            }
        }

        #endregion

    }
}
