using System;
using System.Collections.Generic;
using System.Text;
using A8NET.Data.DAO;

namespace A8NET.Historico
{
    public class HistoricoOperacao
    {
        HistoricoSituacaoOperacaoDAO _HistoricoDAO;

        #region <<< Construtor >>>
        public HistoricoOperacao()
        {
            _HistoricoDAO = new HistoricoSituacaoOperacaoDAO();
        }
        #endregion

    }
}
