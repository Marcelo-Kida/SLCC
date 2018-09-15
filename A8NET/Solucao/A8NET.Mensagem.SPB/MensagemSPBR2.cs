using System;
using System.Collections.Generic;
using System.Text;
using A8NET.Data.DAO;
using A8NET.Data;
using System.Xml;
using System.Data;
using A8NET.Historico;
using A8NET.ConfiguracaoMQ;
using System.Configuration;

namespace A8NET.Mensagem.SPB
{
    public class MensagemSPBR2 : MensagemSPB
    { 
        #region <<< Construtor >>>
        public MensagemSPBR2(Data.DsParametrizacoes dsCache): base(dsCache)
        {
        }
        #endregion
        
        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~MensagemSPBR2()
        {
            this.Dispose();
        }
        #endregion

        #region >>> ObterEventoProcessamento >>>
        public override string ObterEventoProcessamento(string codigoMensagem)
        {
            return "RecebimentoR2";
        }
        #endregion

        public override void GerenciarChamada(udt.udtMensagem entidadeMensagem)
        {
            base.GerenciarChamada(entidadeMensagem);
        }

    }
}
