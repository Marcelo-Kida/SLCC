using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;

namespace A8NET.Data.DAO
{
    public class HistoricoSituacaoOperacaoDAO : BaseDAO
    {
        EstruturaHistoricoSituacaoOperacao _EstruturaHistoricoSituacaoOperacao;

        #region <<< Estrutura >>>
        public partial struct EstruturaHistoricoSituacaoOperacao
        {
			public object NU_SEQU_OPER_ATIV;
			public object DH_SITU_ACAO_OPER_ATIV;
			public object CO_SITU_PROC;
			public object TP_ACAO_OPER_ATIV;
			public object TP_JUST_SITU_PROC;
			public object TX_CNTD_ANTE_ACAO;
			public object CO_USUA_ATLZ;
			public object CO_ETCA_USUA_ATLZ;
		}
        #endregion
        
        #region <<< Propriedades >>>
        public EstruturaHistoricoSituacaoOperacao[] Itens
        {
            get
            {
                //Popula
                EstruturaHistoricoSituacaoOperacao[] lHistoricoSituacaoOperacao = null;
                Int32 lI;

                lHistoricoSituacaoOperacao = new EstruturaHistoricoSituacaoOperacao[_Lista.Count];

                for (lI = 0; lI < _Lista.Count; lI++)
                {
                    lHistoricoSituacaoOperacao[lI] = (EstruturaHistoricoSituacaoOperacao)_Lista[lI];
                }
                return lHistoricoSituacaoOperacao;
            }
        }
        public EstruturaHistoricoSituacaoOperacao TB_HIST_SITU_ACAO_OPER_ATIV
        {
            get { return _EstruturaHistoricoSituacaoOperacao; }
            set { _EstruturaHistoricoSituacaoOperacao = value; }
        }
        public DateTime DataGravacao;
        public bool ObteveDataGravacao;
        #endregion

        #region <<< ProcessarMensagem >>>
        private void ProcessarHistoricoSituacaoOperacao(System.Data.DataView lDView)
        {
            Int32 lI;

            //Processa Lista
            _Lista.Clear();
            for (lI = 0; lI < lDView.Count; lI++)
            {
                EstruturaHistoricoSituacaoOperacao lHistoricoSituacaoOperacao = new EstruturaHistoricoSituacaoOperacao();

							lHistoricoSituacaoOperacao.NU_SEQU_OPER_ATIV = lDView[lI].Row["NU_SEQU_OPER_ATIV"];
							lHistoricoSituacaoOperacao.DH_SITU_ACAO_OPER_ATIV = lDView[lI].Row["DH_SITU_ACAO_OPER_ATIV"];
							lHistoricoSituacaoOperacao.CO_SITU_PROC = lDView[lI].Row["CO_SITU_PROC"];
							lHistoricoSituacaoOperacao.TP_ACAO_OPER_ATIV = lDView[lI].Row["TP_ACAO_OPER_ATIV"];
							lHistoricoSituacaoOperacao.TP_JUST_SITU_PROC = lDView[lI].Row["TP_JUST_SITU_PROC"];
							lHistoricoSituacaoOperacao.TX_CNTD_ANTE_ACAO = lDView[lI].Row["TX_CNTD_ANTE_ACAO"];
							lHistoricoSituacaoOperacao.CO_USUA_ATLZ = lDView[lI].Row["CO_USUA_ATLZ"];
							lHistoricoSituacaoOperacao.CO_ETCA_USUA_ATLZ = lDView[lI].Row["CO_ETCA_USUA_ATLZ"];

                //Adiciona
                _Lista.Add(lHistoricoSituacaoOperacao);
            }
            if (_Lista.Count > 0) TB_HIST_SITU_ACAO_OPER_ATIV = (EstruturaHistoricoSituacaoOperacao)_Lista[0];
            //Define DView Final
            _DView = lDView;
        }
        #endregion

        #region <<< Inserir >>>
        public void Inserir(EstruturaHistoricoSituacaoOperacao parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_HIST_SITU_ACAO_OPER.SPI_TB_HIST_SITU_ACAO_OPER";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.DH_SITU_ACAO_OPER_ATIV(parametro.DH_SITU_ACAO_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SITU_PROC(parametro.CO_SITU_PROC, ParameterDirection.Input),
                        A8NETOracleParameter.TP_ACAO_OPER_ATIV(parametro.TP_ACAO_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.TP_JUST_SITU_PROC(parametro.TP_JUST_SITU_PROC, ParameterDirection.Input),
                        A8NETOracleParameter.TX_CNTD_ANTE_ACAO(parametro.TX_CNTD_ANTE_ACAO, ParameterDirection.Input),
                        A8NETOracleParameter.CO_USUA_ATLZ(parametro.CO_USUA_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ETCA_USUA_ATLZ(parametro.CO_ETCA_USUA_ATLZ, ParameterDirection.Input)
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

        #region <<< AtualizarJustificativa >>>
        public void AtualizarJustificativa(EstruturaHistoricoSituacaoOperacao parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_HIST_SITU_ACAO_OPER.SPU_TB_HIST_SITU_ACAO_OPER";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
						A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.CO_SITU_PROC(parametro.CO_SITU_PROC, ParameterDirection.Input),
						A8NETOracleParameter.TP_JUST_SITU_PROC(parametro.TP_JUST_SITU_PROC, ParameterDirection.Input)}
                        );
                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("HistoricoSituacaoOperacaoDAO.AtualizarJustificativa() - " + ex.ToString());
            }
        }
        #endregion


        #region <<< ObterDataGravacao >>>
        public DateTime ObterDataGravacao(int sequOperAtivo, int situacaoProcesso)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_HIST_SITU_ACAO_OPER.SPS_TB_HIST_MAX";

                    OracleParameter ParametroOUT = A8NETOracleParameter.DH_SITU_ACAO_OPER_ATIV(null, ParameterDirection.Output);

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.NU_SEQU_OPER_ATIV(sequOperAtivo, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SITU_PROC(situacaoProcesso, ParameterDirection.Input),
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
                throw new Exception("HistoricoSituacaoOperacaoDAO.ObterDataGravacao() - " + ex.ToString());
            }
        }
        #endregion

    }
}
