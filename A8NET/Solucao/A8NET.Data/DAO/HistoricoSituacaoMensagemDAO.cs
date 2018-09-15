using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;

namespace A8NET.Data.DAO
{
    public class HistoricoSituacaoMensagemDAO:BaseDAO
    {
        EstruturaHistoricoSituacaoMsg _HistoricoSituacaoMsg;

        #region >>> Contrutor >>>
        public HistoricoSituacaoMensagemDAO():base()
        {

        }
        #endregion

        #region <<< Estrutura >>>
        public partial struct EstruturaHistoricoSituacaoMsg
        {
            public object NU_SEQU_CNTR_REPE;
            public object NU_CTRL_IF;
            public object DH_REGT_MESG_SPB;
            public object DH_SITU_ACAO_MESG_SPB;
            public object CO_SITU_PROC;
            public object TP_ACAO_MESG_SPB;
            public object TX_CNTD_ANTE_ACAO;
            public object CO_USUA_ULTI_ATLZ;
            public object CO_ETCA_TRAB_ULTI_ATLZ;
        }
        #endregion

        #region <<< Propriedades >>>
        public EstruturaHistoricoSituacaoMsg[] Itens
        {
            get
            {
                //Popula
                EstruturaHistoricoSituacaoMsg[] lHistoricoSituacaoMsg = null;
                Int32 lI;

                lHistoricoSituacaoMsg = new EstruturaHistoricoSituacaoMsg[_Lista.Count];

                for (lI = 0; lI < _Lista.Count; lI++)
                {
                    lHistoricoSituacaoMsg[lI] = (EstruturaHistoricoSituacaoMsg)_Lista[lI];
                }
                return lHistoricoSituacaoMsg;
            }
        }
        public EstruturaHistoricoSituacaoMsg TB_HIST_SITU_ACAO_MESG_SPB
        {
            get { return _HistoricoSituacaoMsg; }
            set { _HistoricoSituacaoMsg = value; }
        }
        public DateTime DataGravacao;
        public bool ObteveDataGravacao = false;
        #endregion

        #region <<< ProcessarMensagem >>>
        private void ProcessarHistoricoSituacaoMsg(System.Data.DataView lDView)
        {
            Int32 lI;

            //Processa Lista
            _Lista.Clear();
            for (lI = 0; lI < lDView.Count; lI++)
            {
                EstruturaHistoricoSituacaoMsg lHistoricoSituacaoMsg = new EstruturaHistoricoSituacaoMsg();

                lHistoricoSituacaoMsg.NU_SEQU_CNTR_REPE = lDView[lI].Row["NU_SEQU_CNTR_REPE"];
                lHistoricoSituacaoMsg.NU_CTRL_IF = lDView[lI].Row["NU_CTRL_IF"];
                lHistoricoSituacaoMsg.DH_REGT_MESG_SPB = lDView[lI].Row["DH_REGT_MESG_SPB"];
                lHistoricoSituacaoMsg.DH_SITU_ACAO_MESG_SPB = lDView[lI].Row["DH_SITU_ACAO_MESG_SPB"];
                lHistoricoSituacaoMsg.CO_SITU_PROC = lDView[lI].Row["CO_SITU_PROC"];
                lHistoricoSituacaoMsg.TP_ACAO_MESG_SPB = lDView[lI].Row["TP_ACAO_MESG_SPB"];
                lHistoricoSituacaoMsg.TX_CNTD_ANTE_ACAO = lDView[lI].Row["TX_CNTD_ANTE_ACAO"];
                lHistoricoSituacaoMsg.CO_USUA_ULTI_ATLZ = lDView[lI].Row["CO_USUA_ULTI_ATLZ"];
                lHistoricoSituacaoMsg.CO_ETCA_TRAB_ULTI_ATLZ = lDView[lI].Row["CO_ETCA_TRAB_ULTI_ATLZ"];

                //Adiciona
                _Lista.Add(lHistoricoSituacaoMsg);
            }
            if (_Lista.Count > 0) TB_HIST_SITU_ACAO_MESG_SPB = (EstruturaHistoricoSituacaoMsg)_Lista[0];
            //Define DView Final
            _DView = lDView;
        }
        #endregion

        #region <<< ObterDataGravacao >>>
        public DateTime ObterDataGravacao(string controleIF)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(new BaseDAO().GetStringConnection()))
                {
                    OracleConn.Open();
                    OracleCommand OracleCommand = new OracleCommand();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_HIST_SITU_ACAO_MESG.SPS_TB_HIST_MAX";

                    OracleParameter ParametroOUT = A8NETOracleParameter.DH_SITU_ACAO_MESG_SPB(null, ParameterDirection.Output);

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.NU_CTRL_IF(controleIF, ParameterDirection.Input),
                       ParametroOUT}
                       );

                    _OracleCommand.ExecuteNonQuery();

                    if (ParametroOUT.Value == DBNull.Value) DataGravacao = DateTime.Now;
                    else DataGravacao = DateTime.Parse(ParametroOUT.Value.ToString());

                    if (DataGravacao.Equals(DateTime.Now)) DataGravacao = DataGravacao.AddSeconds(1);
                    else DataGravacao = DateTime.Now;

                    ObteveDataGravacao = true;

                    return DataGravacao;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("HistoricoSituacaoMensagemDAO.ObterDataGravacao() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< Inserir >>>
        public void Inserir(EstruturaHistoricoSituacaoMsg parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    //_OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_HIST_SITU_ACAO_MESG.SPI_TB_HIST_SITU_ACAO_MESG_SPB";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.NU_SEQU_CNTR_REPE(parametro.NU_SEQU_CNTR_REPE, ParameterDirection.Input),
                        A8NETOracleParameter.NU_CTRL_IF(parametro.NU_CTRL_IF, ParameterDirection.Input),
                        A8NETOracleParameter.DH_REGT_MESG_SPB(parametro.DH_REGT_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.DH_SITU_ACAO_MESG_SPB(parametro.DH_SITU_ACAO_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SITU_PROC(parametro.CO_SITU_PROC, ParameterDirection.Input),
                        A8NETOracleParameter.TP_ACAO_MESG_SPB(parametro.TP_ACAO_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.TX_CNTD_ANTE_ACAO(parametro.TX_CNTD_ANTE_ACAO, ParameterDirection.Input),
                        A8NETOracleParameter.CO_USUA_ULTI_ATLZ(parametro.CO_USUA_ULTI_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(parametro.CO_ETCA_TRAB_ULTI_ATLZ, ParameterDirection.Input)

                });
                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSpbDAO.Inseir() - " + ex.ToString());
            }
        }
        #endregion


    }
}
