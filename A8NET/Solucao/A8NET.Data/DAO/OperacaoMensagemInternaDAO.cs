using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OracleClient;

namespace A8NET.Data.DAO
{
    public class OperacaoMensagemInternaDAO : BaseDAO
    {
        EstruturaOperacaoMensagemInterna _OperacaoMensagemInterna;

        #region <<< Estrutura >>>
        public partial struct EstruturaOperacaoMensagemInterna
        {
            public object NU_SEQU_OPER_ATIV;
            public object DH_MESG_INTE;
            public object TP_MESG_INTE;
            public object TP_SOLI_MESG_INTE;
            public object CO_TEXT_XML;
            public object TP_FORM_MESG_SAID;
        }
        #endregion

        #region <<< Propriedades >>>
        public DateTime DataGravacao;
        public bool ObteveDataGravacao;
        public EstruturaOperacaoMensagemInterna[] Itens
        {
            get
            {
                //Popula
                EstruturaOperacaoMensagemInterna[] lOperacaoMensagemInterna = null;
                Int32 lI;

                lOperacaoMensagemInterna = new EstruturaOperacaoMensagemInterna[_Lista.Count];

                for (lI = 0; lI < _Lista.Count; lI++)
                {
                    lOperacaoMensagemInterna[lI] = (EstruturaOperacaoMensagemInterna)_Lista[lI];
                }
                return lOperacaoMensagemInterna;
            }
        }
        public EstruturaOperacaoMensagemInterna TB_OPER_ATIV_MESG_INTE
        {
            get { return _OperacaoMensagemInterna; }
            set { _OperacaoMensagemInterna = value; }
        }
        #endregion

        #region <<< ProcessarMensagem >>>
        private void ProcessarOperacaoMensagemInterna(System.Data.DataView lDView)
        {
            Int32 lI;

            //Processa Lista
            _Lista.Clear();
            for (lI = 0; lI < lDView.Count; lI++)
            {
                EstruturaOperacaoMensagemInterna lOperacaoMensagemInterna = new EstruturaOperacaoMensagemInterna();

                lOperacaoMensagemInterna.NU_SEQU_OPER_ATIV = lDView[lI].Row["NU_SEQU_OPER_ATIV"];
                lOperacaoMensagemInterna.DH_MESG_INTE = lDView[lI].Row["DH_MESG_INTE"];
                lOperacaoMensagemInterna.TP_MESG_INTE = lDView[lI].Row["TP_MESG_INTE"];
                lOperacaoMensagemInterna.TP_SOLI_MESG_INTE = lDView[lI].Row["TP_SOLI_MESG_INTE"];
                lOperacaoMensagemInterna.CO_TEXT_XML = lDView[lI].Row["CO_TEXT_XML"];
                lOperacaoMensagemInterna.TP_FORM_MESG_SAID = lDView[lI].Row["TP_FORM_MESG_SAID"];

                //Adiciona
                _Lista.Add(lOperacaoMensagemInterna);
            }
            if (_Lista.Count > 0) TB_OPER_ATIV_MESG_INTE = (EstruturaOperacaoMensagemInterna)_Lista[0];
            //Define DView Final
            _DView = lDView;
        }
        #endregion

        #region <<< ObterDataGravacao >>>
        public DateTime ObterDataGravacao(string seqOperacao)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    OracleConn.Open();
                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV_MESG_INTE.SPS_MAX";

                    OracleParameter ParametroOUT = A8NETOracleParameter.DH_MESG_INTE(null, ParameterDirection.Output);

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        A8NETOracleParameter.NU_SEQU_OPER_ATIV(seqOperacao, ParameterDirection.Input),
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

        #region >>> VerificaIdentificadorOperacao() >>>
        public bool VerificaIdentificadorOperacao(string identificadorOperacao, string siglaSistema)
        {
            OracleParameter ParametroRetorno;

            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPS_TB_OPER_ATIV_03";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_OPER_ATIV(identificadorOperacao.Trim(), ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.SG_SIST(siglaSistema, ParameterDirection.Input));

                    //Adicionar parametro retorno
                    ParametroRetorno = A8NETOracleParameter.QUANTIDADE();
                    _OracleCommand.Parameters.Add(ParametroRetorno);

                    _OracleCommand.ExecuteNonQuery();

                    if (Convert.ToInt16(ParametroRetorno.Value) > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        #endregion

        #region <<< Inserir Registro >>>>>>
		/// <summary>
		///	Método responsável em inserir um registro na tabela TB_OPER_ATIV_MESG_INTE.
		/// </summary>
		/// <param name="registro<<NomeClasseBO>>">Valores para inclusão</param>
		public void Inserir(EstruturaOperacaoMensagemInterna parametro)
		{
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV_MESG_INTE.SPI_TB_OPER_ATIV_MESG_INTE";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
				        A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
				        A8NETOracleParameter.DH_MESG_INTE(parametro.DH_MESG_INTE, ParameterDirection.Input),
				        A8NETOracleParameter.TP_MESG_INTE(parametro.TP_MESG_INTE, ParameterDirection.Input),
				        A8NETOracleParameter.TP_SOLI_MESG_INTE(parametro.TP_SOLI_MESG_INTE, ParameterDirection.Input),
				        A8NETOracleParameter.CO_TEXT_XML(parametro.CO_TEXT_XML, ParameterDirection.Input),
				        A8NETOracleParameter.TP_FORM_MESG_SAID(parametro.TP_FORM_MESG_SAID, ParameterDirection.Input)});

                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("OperacaoMensagemInternaDAO.Inserir() - " + ex.ToString());
            }
		}		
		#endregion

    }
}
