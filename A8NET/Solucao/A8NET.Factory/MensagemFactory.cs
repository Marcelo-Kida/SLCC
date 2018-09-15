using System;
using System.Collections.Generic;
using System.Text;

namespace A8NET.Factory
{
    public class MensagemFactory : IDisposable
    {
        #region <<< Constructors >>>
        public MensagemFactory()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~MensagemFactory()
        {
            this.Dispose();
        }
        #endregion

        #region <<< CriaMensagem >>>
        //public static Mensagem CriaMensagem(string tipoMensagem, A6CaixaColigadas.DsCadastros DataSetCache)
        public static A8NET.Mensagem.Mensagem CriaMensagem(string sistemaOrigem, A8NET.Data.DsParametrizacoes DataSetCache)
        {
            switch (sistemaOrigem)
            {
                case "NZ":
                    return new A8NET.Mensagem.SPB.MensagemSPB(DataSetCache);
                default:
                    return new A8NET.Mensagem.Operacao.Operacao(DataSetCache);
            }
        }
        #endregion
    }
}
