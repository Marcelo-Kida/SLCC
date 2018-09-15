using System;
using System.Collections.Generic;
using System.Text;
using A7NET.Comum;
using A7NET.Data;
using A7NET.Mensagem;

namespace A7NET.Factory
{
    public class MensagemFactory : IDisposable
    {
        #region <<< Construtor >>>
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
        public static A7NET.Mensagem.Mensagem CriaMensagem(string protocoloMensagem, A7NET.Data.DsParametrizacoes DataSetCache)
        {
            int OutN;

            A7NET.Mensagem.udt.udtProtocoloMensagem ProtocoloMensagem = new A7NET.Mensagem.udt.udtProtocoloMensagem();
            ProtocoloMensagem.Parse(protocoloMensagem);

            switch (int.TryParse(ProtocoloMensagem.TipoMensagem, out OutN) ? Convert.ToString(OutN) : ProtocoloMensagem.TipoMensagem.Trim().Substring(0, 3))
            {
                case "CAM": //MensagemA8NZ
                    return new A7NET.Mensagem.MensagemA8NZ(DataSetCache);

                case "1000": //MensagemNZA8
                    return new A7NET.Mensagem.MensagemNZA8(DataSetCache);

                default: //MensagemPadrao
                    return new A7NET.Mensagem.MensagemPadrao(DataSetCache);

            }
        }
        #endregion
    }
}
