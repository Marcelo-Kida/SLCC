using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Text;

namespace A7NET.Data
{
    public static class A7NETOracleParameter
    {
        #region >>>>>> Parametros TB_TEXT_XML >>>>>>
        public static OracleParameter CO_TEXT_XML(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_CO_TEXT_XML", OracleType.Number);
            Parametro.SourceColumn = "CO_TEXT_XML";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_SEQU_TEXT_XML(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_NU_SEQU_TEXT_XML", OracleType.Number);
            Parametro.SourceColumn = "NU_SEQU_TEXT_XML";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TX_XML(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_vch_TX_XML", OracleType.VarChar);
            Parametro.SourceColumn = "TX_XML";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros TB_MESG >>>>>>
        public static OracleParameter SG_SIST_ORIG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_chr_SG_SIST_ORIG", OracleType.Char);
            Parametro.SourceColumn = "SG_SIST_ORIG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_MESG_MQSE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_vch_CO_MESG_MQSE", OracleType.VarChar);
            Parametro.SourceColumn = "CO_MESG_MQSE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_MESG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_vch_TP_MESG", OracleType.VarChar);
            Parametro.SourceColumn = "TP_MESG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_EMPR_ORIG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_CO_EMPR_ORIG", OracleType.Number);
            Parametro.SourceColumn = "CO_EMPR_ORIG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DH_INIC_VIGE_REGR_TRAP(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_dat_DH_INIC_VIGE_REGR_TRAP", OracleType.DateTime);
            Parametro.SourceColumn = "DH_INIC_VIGE_REGR_TRAP";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CMPO_ATRB_IDEF_MESG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_vch_CO_CMPO_ATRB_IDEF_MESG", OracleType.VarChar);
            Parametro.SourceColumn = "CO_CMPO_ATRB_IDEF_MESG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_TEXT_XML_ENTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_CO_TEXT_XML_ENTR", OracleType.Number);
            Parametro.SourceColumn = "CO_TEXT_XML_ENTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_TEXT_XML_SAID(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_CO_TEXT_XML_SAID", OracleType.Number);
            Parametro.SourceColumn = "CO_TEXT_XML_SAID";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_FORM_MESG_SAID(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_TP_FORM_MESG_SAID", OracleType.Number);
            Parametro.SourceColumn = "TP_FORM_MESG_SAID";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros TB_SITU_MESG >>>>>>
        public static OracleParameter CO_MESG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_CO_MESG", OracleType.Number);
            Parametro.SourceColumn = "CO_MESG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_OCOR_MESG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_num_CO_OCOR_MESG", OracleType.Number);
            Parametro.SourceColumn = "CO_OCOR_MESG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TX_DTLH_OCOR_ERRO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_vch_TX_DTLH_OCOR_ERRO", OracleType.VarChar);
            Parametro.SourceColumn = "TX_DTLH_OCOR_ERRO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros TB_MESG_REJE >>>>>>
        public static OracleParameter NO_ARQU_ENTR_FILA_MQSE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_vch_NO_ARQU_ENTR_FILA_MQSE", OracleType.VarChar);
            Parametro.SourceColumn = "NO_ARQU_ENTR_FILA_MQSE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros Genericos >>>>>>
        //public static OracleParameter CD_ACAO(object value, ParameterDirection direcao)
        //{
        //    OracleParameter Parametro = new OracleParameter("p_num_CD_ACAO", OracleType.Number);
        //    Parametro.Value = value;
        //    Parametro.Direction = direcao;
        //    return Parametro;
        //}
        #endregion

        #region >>>>>> Parametros de Saida >>>>>>
        public static OracleParameter CURSOR()
        {
            //Parametro cursor
            OracleParameter Parametro = new OracleParameter("p_cur_out_CURSOR", OracleType.Cursor);
            Parametro.Direction = ParameterDirection.Output;
            Parametro.OracleType = OracleType.Cursor;
            return Parametro;
        }
        public static OracleParameter RETORNO()
        {
            //Parametro Retorno Number
            OracleParameter Parametro = new OracleParameter("p_num_out_RETORNO", OracleType.Number);
            Parametro.Direction = ParameterDirection.Output;
            Parametro.OracleType = OracleType.Number;
            return Parametro;
        }
        public static OracleParameter CO_MESG_OUT()
        {
            OracleParameter Parametro = new OracleParameter("p_num_out_CO_MESG_OUT", OracleType.Number);
            Parametro.Direction = ParameterDirection.Output;
            Parametro.Value = null;
            return Parametro;
        }
        #endregion

    }
}
