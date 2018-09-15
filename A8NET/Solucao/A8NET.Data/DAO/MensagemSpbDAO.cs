using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using A8NET.Comum;
using System.Data;

namespace A8NET.Data.DAO
{
    public class MensagemSpbDAO:BaseDAO
    {
        #region <<< Variaveis >>>
        EstruturaMensagemSPB _MensagemSPB;
        #endregion

        #region <<< Estrutura >>>
        public partial struct EstruturaMensagemSPB
        {
            public object NU_CTRL_IF;
            public object DH_REGT_MESG_SPB;
            public object NU_SEQU_OPER_ATIV;  
            public object NU_SEQU_CNTR_REPE;
            public object NU_SEQU_CNCL_OPER_ATIV_MESG ;
            public object TP_BKOF;
            public object CO_EMPR;
            public object CO_VEIC_LEGA ;
            public object CO_LOCA_LIQU;
            public object SG_SIST;
            public object DH_RECB_ENVI_MESG_SPB;
            public object CO_MESG_SPB;
            public object NU_COMD_OPER;
            public object CO_SITU_MESG_SPB;
            public object CO_TEXT_XML;
            public object HO_ENVI_MESG_SPB;
            public object CO_ULTI_SITU_PROC;
            public object CO_USUA_ULTI_ATLZ;
            public object CO_ETCA_TRAB_ULTI_ATLZ;
            public object DH_ULTI_ATLZ;
            public object IN_ENTR_MANU;
            public object NU_CTRL_CAMR;
            public object IN_CONF_MESG_LTR;
            public object NU_PRTC_MESG_LG;
            public object CO_PARP_CAMR;
            public object TP_ACAO_MESG_SPB_EXEC;
            public object CO_ISPB_PART_CAMR;
            public object TP_CNAL_VEND;
            public object DT_OPER_CAMB_SISBACEN;
            public object CD_CLIE_SISBACEN;
            public object NO_CLIE;
            public object NR_CNPJ_CPF;
            public object CD_MOED_ISO;
            public object VL_MOED_ESTR;
            public object NR_OPER_CAMB_2;
        }
        #endregion

        #region <<< Propriedades >>>
        public EstruturaMensagemSPB[] Itens
        {
            get
            {
                //Popula
                EstruturaMensagemSPB[] lMensagemSPB = null;
                Int32 lI;

                lMensagemSPB = new EstruturaMensagemSPB[_Lista.Count];

                for (lI = 0; lI < _Lista.Count; lI++)
                {
                    lMensagemSPB[lI] = (EstruturaMensagemSPB)_Lista[lI];
                }
                return lMensagemSPB;
            }
        }
        public EstruturaMensagemSPB TB_MESG_RECB_ENVI_SPB
        {
            get { return _MensagemSPB; }
            set { _MensagemSPB = value; }
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
                MensagemSpbDAO.EstruturaMensagemSPB lMensagemSPB = new EstruturaMensagemSPB();

                //Popula
                lMensagemSPB.CO_EMPR = lDView[lI].Row["CO_EMPR"];
                lMensagemSPB.CO_ETCA_TRAB_ULTI_ATLZ = lDView[lI].Row["CO_ETCA_TRAB_ULTI_ATLZ"];
                lMensagemSPB.CO_ISPB_PART_CAMR = lDView[lI].Row["CO_ISPB_PART_CAMR"];
                lMensagemSPB.CO_LOCA_LIQU = lDView[lI].Row["CO_LOCA_LIQU"];
                lMensagemSPB.CO_MESG_SPB = lDView[lI].Row["CO_MESG_SPB"];
                lMensagemSPB.CO_PARP_CAMR = lDView[lI].Row["CO_PARP_CAMR"];
                lMensagemSPB.CO_SITU_MESG_SPB = lDView[lI].Row["CO_SITU_MESG_SPB"];
                lMensagemSPB.CO_TEXT_XML = lDView[lI].Row["CO_TEXT_XML"];
                lMensagemSPB.CO_ULTI_SITU_PROC = lDView[lI].Row["CO_ULTI_SITU_PROC"];
                lMensagemSPB.CO_USUA_ULTI_ATLZ = lDView[lI].Row["CO_USUA_ULTI_ATLZ"];
                lMensagemSPB.CO_VEIC_LEGA = lDView[lI].Row["CO_VEIC_LEGA"];
                lMensagemSPB.DH_RECB_ENVI_MESG_SPB = lDView[lI].Row["DH_RECB_ENVI_MESG_SPB"];
                lMensagemSPB.DH_REGT_MESG_SPB = lDView[lI].Row["DH_REGT_MESG_SPB"];
                lMensagemSPB.DH_ULTI_ATLZ = lDView[lI].Row["DH_ULTI_ATLZ"];
                lMensagemSPB.HO_ENVI_MESG_SPB = lDView[lI].Row["HO_ENVI_MESG_SPB"];
                lMensagemSPB.IN_CONF_MESG_LTR = lDView[lI].Row["IN_CONF_MESG_LTR"];
                lMensagemSPB.IN_ENTR_MANU = lDView[lI].Row["IN_ENTR_MANU"];
                lMensagemSPB.NU_COMD_OPER = lDView[lI].Row["NU_COMD_OPER"];
                lMensagemSPB.NU_CTRL_CAMR = lDView[lI].Row["NU_CTRL_CAMR"];
                lMensagemSPB.NU_CTRL_IF = lDView[lI].Row["NU_CTRL_IF"];
                lMensagemSPB.NU_PRTC_MESG_LG = lDView[lI].Row["NU_PRTC_MESG_LG"];
                lMensagemSPB.NU_SEQU_CNCL_OPER_ATIV_MESG = lDView[lI].Row["NU_SEQU_CNCL_OPER_ATIV_MESG"];
                lMensagemSPB.NU_SEQU_CNTR_REPE = lDView[lI].Row["NU_SEQU_CNTR_REPE"];
                lMensagemSPB.NU_SEQU_OPER_ATIV = lDView[lI].Row["NU_SEQU_OPER_ATIV"];
                lMensagemSPB.SG_SIST = lDView[lI].Row["SG_SIST"];
                lMensagemSPB.TP_ACAO_MESG_SPB_EXEC = lDView[lI].Row["TP_ACAO_MESG_SPB_EXEC"];
                lMensagemSPB.TP_BKOF = lDView[lI].Row["TP_BKOF"];
                lMensagemSPB.TP_CNAL_VEND = lDView[lI].Row["TP_CNAL_VEND"];
                lMensagemSPB.NR_OPER_CAMB_2 = lDView[lI].Row["NR_OPER_CAMB_2"];
                
                //Adiciona
                _Lista.Add(lMensagemSPB);
            }
            if (_Lista.Count == 1) TB_MESG_RECB_ENVI_SPB = (EstruturaMensagemSPB)_Lista[0];
            //Define DView Final
            _DView = lDView;
        }
        #endregion

        #region <<< SelecionarMensagensPorControleIF >>>
        public DsTB_MESG_RECB_ENVI_SPB.TB_MESG_RECB_ENVI_SPBDataTable SelecionarMensagensPorControleIF(string numeroControleIF)
        {
            DsTB_MESG_RECB_ENVI_SPB DsMesg = new DsTB_MESG_RECB_ENVI_SPB();
            
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    
                    //Abrir conexao banco de dados
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_MESG_RECB_ENVI_SPB.SPS_TB_MESG_RECB_ENVI_SPB";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_CTRL_IF(numeroControleIF, ParameterDirection.Input));

                    _OraDA.Fill(DsMesg.TB_MESG_RECB_ENVI_SPB);

                    this.ProcessarMensagem(DsMesg.TB_MESG_RECB_ENVI_SPB.DefaultView);
                    
                    return DsMesg.TB_MESG_RECB_ENVI_SPB;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSpbDAO.SelecionarMensagensPorControleIF() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< SelecionarMensagensPorNumeroSequenciaOperacao >>>
        public void SelecionarMensagensPorNumeroSequenciaOperacao(long numeroSequenciaOperacao,
                                                              ref DsTB_MESG_RECB_ENVI_SPB dsMesg)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {

                    //Abrir conexao banco de dados
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_MESG_RECB_ENVI_SPB.SPS_TB_MESG_RECB_ENVI_SPB4";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_OPER_ATIV(numeroSequenciaOperacao, ParameterDirection.Input));

                    _OraDA.Fill(dsMesg.TB_MESG_RECB_ENVI_SPB);

                    //this.ProcessarMensagem(DsMesg.TB_MESG_RECB_ENVI_SPB.DefaultView);

                    return; // DsMesg.TB_MESG_RECB_ENVI_SPB;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSpbDAO.SelecionarMensagensPorNumeroSequenciaOperacao() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterMensagemLida >>>
        /// <summary>
        /// Método retorna a mensagem de R0 
        /// </summary>
        /// <returns>retorna a estrutura da mensagem</returns>
        public EstruturaMensagemSPB ObterMensagemLida()
        {
            try
            {
                _DView.RowFilter = "";
                if(_DView.Table.Rows.Count == 0) return new EstruturaMensagemSPB();
                _DView.RowFilter = "DH_REGT_MESG_SPB='" + DateTime.Parse(_DView.Table.Compute("MIN(DH_REGT_MESG_SPB)","").ToString()) + "'";

                ProcessarMensagem(_DView);
                
                return TB_MESG_RECB_ENVI_SPB;
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSpbDAO.ObterMensagemLida()" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterUltimaMensagemPorNumeroControleIF >>>
        /// <summary>
        /// Método retorna a mensagem de R0 
        /// </summary>
        /// <returns>retorna a estrutura da mensagem</returns>
        public EstruturaMensagemSPB ObterUltimaMensagemPorNumeroControleIF()
        {
            try
            {
                _DView.RowFilter = "";
                if (_DView.Table.Rows.Count == 0) return new EstruturaMensagemSPB();
                _DView.RowFilter = "DH_REGT_MESG_SPB='" + DateTime.Parse(_DView.Table.Compute("MAX(DH_REGT_MESG_SPB)", "").ToString()) + "'";

                ProcessarMensagem(_DView);

                return TB_MESG_RECB_ENVI_SPB;
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSpbDAO.ObterUltimaMensagemPorNumeroControleIF()" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterMensagemLidaUnica >>>
        /// <summary>
        /// Método retorna a mensagem de R0 
        /// </summary>
        /// <returns>retorna a estrutura da mensagem</returns>
        public EstruturaMensagemSPB ObterMensagemLidaUnica(DateTime datahoraRegistroMensagem, int numeroControleRepeticao)
        {
            try
            {
                _DView.RowFilter = "";
                if (_DView.Table.Rows.Count == 0) return new EstruturaMensagemSPB();
                _DView.RowFilter = "DH_REGT_MESG_SPB='" + datahoraRegistroMensagem + "' AND NU_SEQU_CNTR_REPE =" + numeroControleRepeticao;

                ProcessarMensagem(_DView);

                return TB_MESG_RECB_ENVI_SPB;
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSpbDAO.ObterMensagemLidaUnica()" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterDataGravacao >>>
        public DateTime ObterDataGravacao(string controleIF)
        {
            try
            {
                if (_DView == null) {
                    this.SelecionarMensagensPorControleIF(controleIF);
                }

                _DView.RowFilter = "";
                if (_DView.Table.Rows.Count == 0) return DateTime.Now;

                if(DateTime.Now.Equals(DateTime.Parse(_DView.Table.Compute("MAX(DH_REGT_MESG_SPB)","").ToString())))
                {
                    return DateTime.Now.AddSeconds(1);
                }
                else 
                {
                    return DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ObterDataGravacao()" + ex.ToString());
            }         
        }
        #endregion

        #region <<< Inserir >>>
        public void Inserir(EstruturaMensagemSPB parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_MESG_RECB_ENVI_SPB.SPI_TB_MESG_RECB_ENVI_SPB";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.NU_CTRL_IF(parametro.NU_CTRL_IF, ParameterDirection.Input),
                        A8NETOracleParameter.DH_REGT_MESG_SPB(parametro.DH_REGT_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.NU_SEQU_CNTR_REPE(parametro.NU_SEQU_CNTR_REPE, ParameterDirection.Input),
                        A8NETOracleParameter.CO_LOCA_LIQU(parametro.CO_LOCA_LIQU, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ULTI_SITU_PROC(parametro.CO_ULTI_SITU_PROC, ParameterDirection.Input),
                        A8NETOracleParameter.CO_VEIC_LEGA(parametro.CO_VEIC_LEGA, ParameterDirection.Input),
                        A8NETOracleParameter.SG_SIST(parametro.SG_SIST, ParameterDirection.Input),
                        A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.TP_BKOF(parametro.TP_BKOF, ParameterDirection.Input),
                        A8NETOracleParameter.CO_EMPR(parametro.CO_EMPR, ParameterDirection.Input),
                        A8NETOracleParameter.DH_RECB_ENVI_MESG_SPB(parametro.DH_RECB_ENVI_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.NU_COMD_OPER(parametro.NU_COMD_OPER, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SITU_MESG_SPB(parametro.CO_SITU_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.CO_TEXT_XML(parametro.CO_TEXT_XML, ParameterDirection.Input),
                        A8NETOracleParameter.HO_ENVI_MESG_SPB(parametro.HO_ENVI_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENTR_MANU(parametro.IN_ENTR_MANU, ParameterDirection.Input),
                        A8NETOracleParameter.NU_SEQU_CNCL_OPER_ATIV_MESG(parametro.NU_SEQU_CNCL_OPER_ATIV_MESG, ParameterDirection.Input),
                        A8NETOracleParameter.NU_CTRL_CAMR(parametro.NU_CTRL_CAMR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_MESG_SPB(parametro.CO_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.IN_CONF_MESG_LTR(parametro.IN_CONF_MESG_LTR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_USUA_ULTI_ATLZ(parametro.CO_USUA_ULTI_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(parametro.CO_ETCA_TRAB_ULTI_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.NU_PRTC_MESG_LG(parametro.NU_PRTC_MESG_LG, ParameterDirection.Input),
                        A8NETOracleParameter.CO_PARP_CAMR(parametro.CO_PARP_CAMR, ParameterDirection.Input),
                        A8NETOracleParameter.TP_ACAO_MESG_SPB_EXEC(parametro.TP_ACAO_MESG_SPB_EXEC, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ISPB_PART_CAMR(parametro.CO_ISPB_PART_CAMR, ParameterDirection.Input),
                        A8NETOracleParameter.TP_CNAL_VEND(parametro.TP_CNAL_VEND, ParameterDirection.Input),
                        A8NETOracleParameter.DT_OPER_CAMB_SISBACEN(parametro.DT_OPER_CAMB_SISBACEN, ParameterDirection.Input),
                        A8NETOracleParameter.CD_CLIE_SISBACEN(parametro.CD_CLIE_SISBACEN, ParameterDirection.Input),
                        A8NETOracleParameter.NR_OPER_CAMB_2(parametro.NR_OPER_CAMB_2, ParameterDirection.Input)
                        
                        // ESTÁ IMPLEMENTAÇÃO ESTÁ EM STAND-BY, AGUARDANDO PRIORIZAÇÃO PARA SER IMPLANTADA
                        //A8NETOracleParameter.NO_CLIE(parametro.NO_CLIE, ParameterDirection.Input),
                        //A8NETOracleParameter.NR_CNPJ_CPF(parametro.NR_CNPJ_CPF, ParameterDirection.Input),
                        //A8NETOracleParameter.CD_MOED_ISO(parametro.CD_MOED_ISO, ParameterDirection.Input),
                        //A8NETOracleParameter.VL_MOED_ESTR(parametro.VL_MOED_ESTR, ParameterDirection.Input),
                        //A8NETOracleParameter.NR_OPER_CAMB(parametro.NR_OPER_CAMB, ParameterDirection.Input)
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

        #region <<< Atualizar >>>
        public void Atualizar(EstruturaMensagemSPB parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_MESG_RECB_ENVI_SPB.SPU_TB_MESG_RECB_ENVI_SPB";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
						A8NETOracleParameter.NU_CTRL_IF(parametro.NU_CTRL_IF, ParameterDirection.Input),
						A8NETOracleParameter.DH_REGT_MESG_SPB(parametro.DH_REGT_MESG_SPB, ParameterDirection.Input),
						A8NETOracleParameter.NU_SEQU_CNTR_REPE(parametro.NU_SEQU_CNTR_REPE, ParameterDirection.Input),
						A8NETOracleParameter.CO_LOCA_LIQU(parametro.CO_LOCA_LIQU, ParameterDirection.Input),
						A8NETOracleParameter.CO_ULTI_SITU_PROC(parametro.CO_ULTI_SITU_PROC, ParameterDirection.Input),
						A8NETOracleParameter.CO_VEIC_LEGA(parametro.CO_VEIC_LEGA, ParameterDirection.Input),
						A8NETOracleParameter.SG_SIST(parametro.SG_SIST, ParameterDirection.Input),
						A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.TP_BKOF(parametro.TP_BKOF, ParameterDirection.Input),
						A8NETOracleParameter.CO_EMPR(parametro.CO_EMPR, ParameterDirection.Input),
						A8NETOracleParameter.DH_RECB_ENVI_MESG_SPB(parametro.DH_RECB_ENVI_MESG_SPB, ParameterDirection.Input),
						A8NETOracleParameter.NU_COMD_OPER(parametro.NU_COMD_OPER, ParameterDirection.Input),
						A8NETOracleParameter.CO_SITU_MESG_SPB(parametro.CO_SITU_MESG_SPB, ParameterDirection.Input),
						A8NETOracleParameter.CO_TEXT_XML(parametro.CO_TEXT_XML, ParameterDirection.Input),
						A8NETOracleParameter.HO_ENVI_MESG_SPB(parametro.HO_ENVI_MESG_SPB, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENTR_MANU(parametro.IN_ENTR_MANU, ParameterDirection.Input),
						A8NETOracleParameter.NU_SEQU_CNCL_OPER_ATIV_MESG(parametro.NU_SEQU_CNCL_OPER_ATIV_MESG, ParameterDirection.Input),
						A8NETOracleParameter.NU_CTRL_CAMR(parametro.NU_CTRL_CAMR, ParameterDirection.Input),
						A8NETOracleParameter.CO_MESG_SPB(parametro.CO_MESG_SPB, ParameterDirection.Input),
						A8NETOracleParameter.IN_CONF_MESG_LTR(parametro.IN_CONF_MESG_LTR, ParameterDirection.Input),
						A8NETOracleParameter.CO_USUA_ULTI_ATLZ(parametro.CO_USUA_ULTI_ATLZ, ParameterDirection.Input),
						A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(parametro.CO_ETCA_TRAB_ULTI_ATLZ, ParameterDirection.Input),
						A8NETOracleParameter.DH_ULTI_ATLZ(parametro.DH_ULTI_ATLZ, ParameterDirection.Input),
						A8NETOracleParameter.NU_PRTC_MESG_LG(parametro.NU_PRTC_MESG_LG, ParameterDirection.Input),
						A8NETOracleParameter.CO_PARP_CAMR(parametro.CO_PARP_CAMR, ParameterDirection.Input),
						A8NETOracleParameter.TP_ACAO_MESG_SPB_EXEC(parametro.TP_ACAO_MESG_SPB_EXEC, ParameterDirection.Input),
						A8NETOracleParameter.CO_ISPB_PART_CAMR(parametro.CO_ISPB_PART_CAMR, ParameterDirection.Input),
						A8NETOracleParameter.TP_CNAL_VEND(parametro.TP_CNAL_VEND, ParameterDirection.Input)}
                        );
                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("MensagemSPBDAO.Alterar() - " + ex.ToString());
            }
        }
		#endregion
    }
}
