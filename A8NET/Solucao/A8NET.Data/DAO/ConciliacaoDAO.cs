using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OracleClient;
using A8NET.Data.DAO;

namespace A8NET.Data.DAO
{
    public class ConciliacaoDAO : BaseDAO  // MensagemSpbDAO
    {
        #region <<< Variaveis >>>
        EstruturaConciliacao _Conciliacao;
        #endregion

        #region <<< Estrutura >>>
        public partial struct EstruturaConciliacao
        {
            public object NU_SEQU_CNCL_OPER_ATIV_MESG;
            public object NU_SEQU_OPER_ATIV;
            public object NU_CTRL_IF;
            public object DH_REGT_MESG_SPB;
            public object QT_ATIV_MERC_CNCL;
            public object NU_SEQU_CNTR_REPE;
        }
        #endregion

        #region <<< Propriedades >>>
        public EstruturaConciliacao[] Itens
        {
            get
            {
                //Popula
                EstruturaConciliacao[] lConciliacao = null;
                Int32 lI;

                lConciliacao = new EstruturaConciliacao[_Lista.Count];

                for (lI = 0; lI < _Lista.Count; lI++)
                {
                    lConciliacao[lI] = (EstruturaConciliacao)_Lista[lI];
                }
                return lConciliacao;
            }
        }

        public EstruturaConciliacao TB_CNCL_OPER_ATIV
        {
            get { return _Conciliacao; }
            set { _Conciliacao = value; }
        }
        #endregion

        #region <<< ConciliarOperacaoBMC0015 >>>
        public bool ConciliarOperacaoBMC0015(long numeroSequenciaOperacao,
                                             Comum.Comum.EnumStatusOperacao codigoUltimaSituacaoOperacao,
                                             Comum.Comum.EnumStatusMensagem codigoUltimaSituacaoMensagemSPB,
                                         ref long numeroSequenciaConciliacao,
                                         ref int retornoConciliacao)
        {
            OracleParameter ParametroRetorno;
            OracleParameter ParametroRetornoNumeroSequenciaConciliacao;

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_CONCILIACAO_OPER_MESG.CONCILIAR_OPERACAO_X_BMC0015";

                    ParametroRetorno = A8NETOracleParameter.RETORNO2();
                    ParametroRetornoNumeroSequenciaConciliacao = A8NETOracleParameter.RETORNO_NU_SEQU_CNCL_OPER_ATIV_MESG();

                    _OracleCommand.Parameters.Add(ParametroRetorno);
                    _OracleCommand.Parameters.Add(ParametroRetornoNumeroSequenciaConciliacao);
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_OPER_ATIV(numeroSequenciaOperacao, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_ULTI_SITU_PROC_OPERACAO((int)codigoUltimaSituacaoOperacao, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_ULTI_SITU_PROC_MSGSPB((int)codigoUltimaSituacaoMensagemSPB, ParameterDirection.Input));

                    _OracleCommand.ExecuteNonQuery();

                    if (ParametroRetorno.Value != null && ParametroRetornoNumeroSequenciaConciliacao.Value != null)
                    {
                        numeroSequenciaConciliacao = long.Parse(ParametroRetornoNumeroSequenciaConciliacao.Value.ToString());
                        retornoConciliacao = int.Parse(ParametroRetorno.Value.ToString());
                        if (retornoConciliacao == 1)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        numeroSequenciaConciliacao = 0;
                        retornoConciliacao = 0;
                        return false;
                    }

                }
            }
            catch (Exception ex)
            {

                throw new Exception("GestaoCaixaDAO.ConciliarOperacaoBMC0015() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ConciliarComOperacao >>>
        public bool ConciliarComOperacao(string numeroComando,
                                         long registroOperacaoCambial2,
                                         DateTime dataOperacao,
                                         Comum.Comum.EnumStatusOperacao[] listaCodigosUltimaSituacaoOperacao,
                                     ref long numeroSequenciaOperacao,
                                     ref int retornoConciliacao)
        {
            DsTB_OPER_ATIV DsMesg = new DsTB_OPER_ATIV();

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPS_TB_OPER_ATIV_05";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_COMD_OPER(numeroComando, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NR_OPER_CAMB_2(registroOperacaoCambial2, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.DT_OPER_ATIV(dataOperacao, ParameterDirection.Input));

                    _OraDA.Fill(DsMesg.TB_OPER_ATIV);

                    // Obtem qtde de Operações encontradas
                    retornoConciliacao = DsMesg.TB_OPER_ATIV.DefaultView.Count;

                    if (retornoConciliacao == 1) // Se encontrou apenas 1 Operação então Conciliação OK
                    {
                        // Verifica se a Operação encontrada tem algum dos Status esperados
                        foreach (Comum.Comum.EnumStatusOperacao statusOperacao in listaCodigosUltimaSituacaoOperacao)
                        {
                            if ((int)statusOperacao == int.Parse(DsMesg.TB_OPER_ATIV.DefaultView[0]["CO_ULTI_SITU_PROC"].ToString()))
                            {
                                numeroSequenciaOperacao = int.Parse(DsMesg.TB_OPER_ATIV.DefaultView[0]["NU_SEQU_OPER_ATIV"].ToString());
                                return true;
                            }
                        }
                        return false; // se chegar neste ponto é porque a Operação encontrada não tem nenhum dos status esperados
                    }
                    else if (retornoConciliacao == 0) // Se não encontrou Operação então Conciliação NOK
                    {
                        return false;
                    }
                    else // >=2, Se encontrou mais de uma Operação então Conciliação NOK
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {

                throw new Exception("ConciliacaoDAO.ConciliarComOperacao() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< ConciliarComMensagemSPB >>>
        public bool ConciliarComMensagemSPB(string numeroComando,
                                            long registroOperacaoCambial2,
                                            DateTime dataRegistroMensagem,
                                            Comum.Comum.EnumStatusMensagem codigoUltimaSituacaoMensagem,
                                            string codigoMensagemSPB,
                                        ref DsTB_MESG_RECB_ENVI_SPB dsMesg,
                                        ref int retornoConciliacao)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_MESG_RECB_ENVI_SPB.SPS_TB_MESG_RECB_ENVI_SPB3";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_COMD_OPER(numeroComando, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NR_OPER_CAMB_2(registroOperacaoCambial2, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.DT_OPER_ATIV(dataRegistroMensagem, ParameterDirection.Input));

                    _OraDA.Fill(dsMesg.TB_MESG_RECB_ENVI_SPB);

                    // Obtem qtde de Mensagens encontradas
                    dsMesg.TB_MESG_RECB_ENVI_SPB.DefaultView.RowFilter = string.Format("CO_MESG_SPB ='{0}'", codigoMensagemSPB);
                    retornoConciliacao = dsMesg.TB_MESG_RECB_ENVI_SPB.DefaultView.Count;

                    if (retornoConciliacao == 1) // Se encontrou apenas 1 Mensagem então Conciliação OK
                    {
                        // Verifica se a Mensagem encontrada está com o status esperado
                        if ((int)codigoUltimaSituacaoMensagem == int.Parse(dsMesg.TB_MESG_RECB_ENVI_SPB.DefaultView[0]["CO_ULTI_SITU_PROC"].ToString()))
                        {
                            return true;
                        }
                        return false; // se chegar neste ponto é porque a Mensagem encontrada não tem o status esperado
                    }
                    else if (retornoConciliacao == 0) // Se não encontrou Mensagem então Conciliação NOK
                    {
                        return false;
                    }
                    else // >=2, Se encontrou mais de uma Mensagem então Conciliação NOK
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {

                throw new Exception("ConciliacaoDAO.ConciliarComMensagemSPB() - " + ex.ToString());
            }

        }
        #endregion

        #region <<< SelecionarConciliacaoOperacao >>>
        public DsTB_CNCL_OPER_ATIV.TB_CNCL_OPER_ATIVDataTable SelecionarConciliacaoOperacao(long numeroSequenciaConciliacao)
        {
            DsTB_CNCL_OPER_ATIV DsMesg = new DsTB_CNCL_OPER_ATIV();
            
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_CONCILIACAO_OPER_MESG.SPS_TB_CNCL_OPER_ATIV";
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_CNCL_OPER_ATIV_MESG(numeroSequenciaConciliacao, ParameterDirection.Input));
                    _OraDA.Fill(DsMesg.TB_CNCL_OPER_ATIV);
                    ProcessarMensagem(DsMesg.TB_CNCL_OPER_ATIV.DefaultView);
                    return DsMesg.TB_CNCL_OPER_ATIV;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ConciliacaoDAO.SelecionarConciliacaoOperacao() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< SelecionarOperacaoConciliada >>>
        public DsTB_CNCL_OPER_ATIV.TB_CNCL_OPER_ATIVDataTable SelecionarOperacaoConciliada(long numeroSequenciaOperacao)
        {
            DsTB_CNCL_OPER_ATIV DsMesg = new DsTB_CNCL_OPER_ATIV();

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_CONCILIACAO_OPER_MESG.SPS_TB_CNCL_OPER_ATIV_02";
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_OPER_ATIV(numeroSequenciaOperacao, ParameterDirection.Input));
                    _OraDA.Fill(DsMesg.TB_CNCL_OPER_ATIV);
                    ProcessarMensagem(DsMesg.TB_CNCL_OPER_ATIV.DefaultView);
                    return DsMesg.TB_CNCL_OPER_ATIV;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ConciliacaoDAO.SelecionarOperacaoConciliada() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< ProcessarMensagem >>>
        protected void ProcessarMensagem(System.Data.DataView lDView)
        {
            Int32 lI;

            //Processa Lista
            _Lista.Clear();
            for (lI = 0; lI < lDView.Count; lI++)
            {
                EstruturaConciliacao lConciliacao = new EstruturaConciliacao();

                //Popula
                lConciliacao.NU_SEQU_CNCL_OPER_ATIV_MESG = lDView[lI].Row["NU_SEQU_CNCL_OPER_ATIV_MESG"];
                lConciliacao.NU_SEQU_OPER_ATIV = lDView[lI].Row["NU_SEQU_OPER_ATIV"];
                lConciliacao.NU_CTRL_IF = lDView[lI].Row["NU_CTRL_IF"];
                lConciliacao.DH_REGT_MESG_SPB = lDView[lI].Row["DH_REGT_MESG_SPB"];
                lConciliacao.QT_ATIV_MERC_CNCL = lDView[lI].Row["QT_ATIV_MERC_CNCL"];
                lConciliacao.NU_SEQU_CNTR_REPE = lDView[lI].Row["NU_SEQU_CNTR_REPE"];
                
                //Adiciona
                _Lista.Add(lConciliacao);
            }
            if (_Lista.Count == 1) TB_CNCL_OPER_ATIV = (EstruturaConciliacao)_Lista[0];
            //Define DView Final
            _DView = lDView;
        }
        #endregion

    }
}
