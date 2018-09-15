using System;
using System.Data;
using System.Data.OracleClient;

namespace A8NET.Data
{
    public static class A8NETOracleParameter
    {

        #region >>>>>> Parametros das Stores Procedures TB_TEXT_XML >>>>>>
        public static OracleParameter CO_TEXT_XML(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_TEXT_XML", OracleType.Number);
            Parametro.SourceColumn = "CO_TEXT_XML";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_SEQU_TEXT_XML(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_SEQU_TEXT_XML", OracleType.Number);
            Parametro.SourceColumn = "NU_SEQU_TEXT_XML";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TX_XML(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TX_XML", OracleType.VarChar);
            Parametro.SourceColumn = "TX_XML";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_HIST_SITU_ACAO_OPER_ATIV >>>>>>
        public static OracleParameter NU_SEQU_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_SEQU_OPER_ATIV", OracleType.Number);
            Parametro.SourceColumn = "NU_SEQU_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DH_SITU_ACAO_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_SITU_ACAO_OPER_ATIV", OracleType.DateTime);
            Parametro.SourceColumn = "DH_SITU_ACAO_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_SITU_PROC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_SITU_PROC", OracleType.Number);
            Parametro.SourceColumn = "CO_SITU_PROC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_ACAO_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_ACAO_OPER_ATIV", OracleType.Number);
            Parametro.SourceColumn = "TP_ACAO_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_JUST_SITU_PROC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_JUST_SITU_PROC", OracleType.Number);
            Parametro.SourceColumn = "TP_JUST_SITU_PROC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TX_CNTD_ANTE_ACAO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TX_CNTD_ANTE_ACAO", OracleType.VarChar);
            Parametro.SourceColumn = "TX_CNTD_ANTE_ACAO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_USUA_ATLZ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_USUA_ATLZ", OracleType.Char);
            Parametro.SourceColumn = "CO_USUA_ATLZ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ETCA_USUA_ATLZ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ETCA_USUA_ATLZ", OracleType.VarChar);
            Parametro.SourceColumn = "CO_ETCA_USUA_ATLZ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_OPER_ATIV_MESG_INTE >>>>>>
       
        public static OracleParameter DH_MESG_INTE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_MESG_INTE", OracleType.DateTime);
            Parametro.SourceColumn = "DH_MESG_INTE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_MESG_INTE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_MESG_INTE", OracleType.VarChar);
            Parametro.SourceColumn = "TP_MESG_INTE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }

        public static OracleParameter TP_SOLI_MESG_INTE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_SOLI_MESG_INTE", OracleType.Number);
            Parametro.SourceColumn = "TP_SOLI_MESG_INTE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
       
        public static OracleParameter TP_FORM_MESG_SAID(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_FORM_MESG_SAID", OracleType.Number);
            Parametro.SourceColumn = "TP_FORM_MESG_SAID";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_OPER_ATIV >>>>>>
        
        public static OracleParameter TP_OPER(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_OPER", OracleType.Number);
            Parametro.SourceColumn = "TP_OPER";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_LOCA_LIQU(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_LOCA_LIQU", OracleType.Number);
            Parametro.SourceColumn = "CO_LOCA_LIQU";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_LIQU_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_LIQU_OPER_ATIV", OracleType.Number);
            Parametro.SourceColumn = "TP_LIQU_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_EMPR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_EMPR", OracleType.Number);
            Parametro.SourceColumn = "CO_EMPR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_USUA_CADR_OPER(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_USUA_CADR_OPER", OracleType.VarChar);
            Parametro.SourceColumn = "CO_USUA_CADR_OPER";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter HO_ENVI_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_HO_ENVI_MESG_SPB", OracleType.DateTime);
            Parametro.SourceColumn = "HO_ENVI_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_OPER_ATIV", OracleType.VarChar);
            Parametro.SourceColumn = "CO_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_COMD_OPER(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_COMD_OPER", OracleType.VarChar);
            Parametro.SourceColumn = "NU_COMD_OPER";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_COMD_OPER_RETN(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_COMD_OPER_RETN", OracleType.VarChar);
            Parametro.SourceColumn = "NU_COMD_OPER_RETN";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DT_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DT_OPER_ATIV", OracleType.DateTime);
            Parametro.SourceColumn = "DT_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DT_OPER_ATIV_RETN(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DT_OPER_ATIV_RETN", OracleType.DateTime);
            Parametro.SourceColumn = "DT_OPER_ATIV_RETN";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_VEIC_LEGA(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_VEIC_LEGA", OracleType.VarChar);
            Parametro.SourceColumn = "CO_VEIC_LEGA";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter SG_SIST(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_SG_SIST", OracleType.Char);
            Parametro.SourceColumn = "SG_SIST";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CNTA_CUTD_SELIC_VEIC_LEGA(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_CNTA_CUTD_SELIC_VEIC_LEGA", OracleType.Number);
            Parametro.SourceColumn = "CO_CNTA_CUTD_SELIC_VEIC_LEGA";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CNPJ_CNPT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_CNPJ_CNPT", OracleType.Number);
            Parametro.SourceColumn = "CO_CNPJ_CNPT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CNTA_CUTD_SELIC_CNPT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_CNTA_CUTD_SELIC_CNPT", OracleType.Number);
            Parametro.SourceColumn = "CO_CNTA_CUTD_SELIC_CNPT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NO_CNPT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NO_CNPT", OracleType.VarChar);
            Parametro.SourceColumn = "NO_CNPT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_OPER_DEBT_CRED(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_OPER_DEBT_CRED", OracleType.Number);
            Parametro.SourceColumn = "IN_OPER_DEBT_CRED";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_ATIV_MERC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_ATIV_MERC", OracleType.VarChar);
            Parametro.SourceColumn = "NU_ATIV_MERC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DE_ATIV_MERC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DE_ATIV_MERC", OracleType.VarChar);
            Parametro.SourceColumn = "DE_ATIV_MERC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter PU_ATIV_MERC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_PU_ATIV_MERC", OracleType.Number);
            Parametro.SourceColumn = "PU_ATIV_MERC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter QT_ATIV_MERC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_QT_ATIV_MERC", OracleType.Number);
            Parametro.SourceColumn = "QT_ATIV_MERC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_ENTR_SAID_RECU_FINC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_ENTR_SAID_RECU_FINC", OracleType.Number);
            Parametro.SourceColumn = "IN_ENTR_SAID_RECU_FINC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DT_VENC_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DT_VENC_ATIV", OracleType.DateTime);
            Parametro.SourceColumn = "DT_VENC_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter VA_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_VA_OPER_ATIV", OracleType.Number);
            Parametro.SourceColumn = "VA_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter VA_OPER_ATIV_REAJ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_VA_OPER_ATIV_REAJ", OracleType.Number);
            Parametro.SourceColumn = "VA_OPER_ATIV_REAJ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DT_LIQU_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DT_LIQU_OPER_ATIV", OracleType.DateTime);
            Parametro.SourceColumn = "DT_LIQU_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_CPRO_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_CPRO_OPER_ATIV", OracleType.Char);
            Parametro.SourceColumn = "TP_CPRO_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_CPRO_RETN_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_CPRO_RETN_OPER_ATIV", OracleType.Char);
            Parametro.SourceColumn = "TP_CPRO_RETN_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_DISP_CONS(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_DISP_CONS", OracleType.Number);
            Parametro.SourceColumn = "IN_DISP_CONS";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_ENVI_PREV_SIST_PJ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_ENVI_PREV_SIST_PJ", OracleType.Number);
            Parametro.SourceColumn = "IN_ENVI_PREV_SIST_PJ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_ENVI_RELZ_SIST_PJ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_ENVI_RELZ_SIST_PJ", OracleType.Number);
            Parametro.SourceColumn = "IN_ENVI_RELZ_SIST_PJ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_ENVI_PREV_SIST_A6(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_ENVI_PREV_SIST_A6", OracleType.Number);
            Parametro.SourceColumn = "IN_ENVI_PREV_SIST_A6";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_ENVI_RELZ_SIST_A6(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_ENVI_RELZ_SIST_A6", OracleType.Number);
            Parametro.SourceColumn = "IN_ENVI_RELZ_SIST_A6";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ULTI_SITU_PROC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ULTI_SITU_PROC", OracleType.Number);
            Parametro.SourceColumn = "CO_ULTI_SITU_PROC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_ACAO_OPER_ATIV_EXEC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_ACAO_OPER_ATIV_EXEC", OracleType.Number);
            Parametro.SourceColumn = "TP_ACAO_OPER_ATIV_EXEC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_COMD_ACAO_EXEC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_COMD_ACAO_EXEC", OracleType.Number);
            Parametro.SourceColumn = "NU_COMD_ACAO_EXEC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_ENTR_MANU(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_ENTR_MANU", OracleType.Number);
            Parametro.SourceColumn = "IN_ENTR_MANU";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_PRTC_OPER_LG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_PRTC_OPER_LG", OracleType.Number);
            Parametro.SourceColumn = "NU_PRTC_OPER_LG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_SEQU_CNCL_OPER_ATIV_MESG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_SEQU_CNCL_OPER_ATIV_MESG", OracleType.Number);
            Parametro.SourceColumn = "NU_SEQU_CNCL_OPER_ATIV_MESG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_CTRL_MESG_SPB_ORIG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_CTRL_MESG_SPB_ORIG", OracleType.VarChar);
            Parametro.SourceColumn = "NU_CTRL_MESG_SPB_ORIG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter PE_TAXA_NEGO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_PE_TAXA_NEGO", OracleType.Number);
            Parametro.SourceColumn = "PE_TAXA_NEGO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_TITL_CUTD(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_TITL_CUTD", OracleType.Number);
            Parametro.SourceColumn = "CO_TITL_CUTD";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_OPER_CETIP(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_OPER_CETIP", OracleType.Number);
            Parametro.SourceColumn = "CO_OPER_CETIP";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ISPB_BANC_LIQU_CNPT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ISPB_BANC_LIQU_CNPT", OracleType.VarChar);
            Parametro.SourceColumn = "CO_ISPB_BANC_LIQU_CNPT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DH_ULTI_ATLZ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_ULTI_ATLZ", OracleType.DateTime);
            Parametro.SourceColumn = "DH_ULTI_ATLZ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ETCA_TRAB_ULTI_ATLZ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ETCA_TRAB_ULTI_ATLZ", OracleType.VarChar);
            Parametro.SourceColumn = "CO_ETCA_TRAB_ULTI_ATLZ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_USUA_ULTI_ATLZ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_USUA_ULTI_ATLZ", OracleType.Char);
            Parametro.SourceColumn = "CO_USUA_ULTI_ATLZ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_IF_CRED_DEBT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_IF_CRED_DEBT", OracleType.Number);
            Parametro.SourceColumn = "TP_IF_CRED_DEBT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_AGEN_COTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_AGEN_COTR", OracleType.Number);
            Parametro.SourceColumn = "CO_AGEN_COTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_CC_COTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_CC_COTR", OracleType.Number);
            Parametro.SourceColumn = "NU_CC_COTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter PZ_DIAS_RETN_OPER_ATIV(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_PZ_DIAS_RETN_OPER_ATIV", OracleType.Number);
            Parametro.SourceColumn = "PZ_DIAS_RETN_OPER_ATIV";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter VA_OPER_ATIV_RETN(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_VA_OPER_ATIV_RETN", OracleType.Number);
            Parametro.SourceColumn = "VA_OPER_ATIV_RETN";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_CNPT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_CNPT", OracleType.Number);
            Parametro.SourceColumn = "TP_CNPT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CNPT_CAMR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_CNPT_CAMR", OracleType.VarChar);
            Parametro.SourceColumn = "CO_CNPT_CAMR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_IDEF_LAST(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_IDEF_LAST", OracleType.VarChar);
            Parametro.SourceColumn = "CO_IDEF_LAST";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_PARP_CAMR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_PARP_CAMR", OracleType.VarChar);
            Parametro.SourceColumn = "CO_PARP_CAMR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_PGTO_LDL(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_PGTO_LDL", OracleType.Number);
            Parametro.SourceColumn = "TP_PGTO_LDL";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_GRUP_LANC_FINC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_GRUP_LANC_FINC", OracleType.Number);
            Parametro.SourceColumn = "CO_GRUP_LANC_FINC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_MOED_ESTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_MOED_ESTR", OracleType.Number);
            Parametro.SourceColumn = "CO_MOED_ESTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CNTR_SISB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_CNTR_SISB", OracleType.Number);
            Parametro.SourceColumn = "CO_CNTR_SISB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ISPB_IF_CNPT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ISPB_IF_CNPT", OracleType.VarChar);
            Parametro.SourceColumn = "CO_ISPB_IF_CNPT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_PRAC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_PRAC", OracleType.Number);
            Parametro.SourceColumn = "CO_PRAC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter VA_MOED_ESTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_VA_MOED_ESTR", OracleType.Number);
            Parametro.SourceColumn = "VA_MOED_ESTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DT_LIQU_OPER_ATIV_MOED_ESTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DT_LIQU_OPER_ATIV_MOED_ESTR", OracleType.DateTime);
            Parametro.SourceColumn = "DT_LIQU_OPER_ATIV_MOED_ESTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_SISB_COTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_SISB_COTR", OracleType.Number);
            Parametro.SourceColumn = "CO_SISB_COTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_CNAL_OPER_INTE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_CNAL_OPER_INTE", OracleType.Char);
            Parametro.SourceColumn = "CO_CNAL_OPER_INTE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_SITU_PROC_MESG_SPB_RECB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_SITU_PROC_MESG_SPB_RECB", OracleType.VarChar);
            Parametro.SourceColumn = "CO_SITU_PROC_MESG_SPB_RECB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_CNAL_VEND(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_CNAL_VEND", OracleType.Number);
            Parametro.SourceColumn = "TP_CNAL_VEND";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CD_SUB_PROD(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CD_SUB_PROD", OracleType.VarChar);
            Parametro.SourceColumn = "CD_SUB_PROD";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NR_IDEF_NEGO_BMC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NR_IDEF_NEGO_BMC", OracleType.Number);
            Parametro.SourceColumn = "NR_IDEF_NEGO_BMC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_NEGO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_NEGO", OracleType.Number);
            Parametro.SourceColumn = "TP_NEGO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CD_ASSO_CAMB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CD_ASSO_CAMB", OracleType.VarChar);
            Parametro.SourceColumn = "CD_ASSO_CAMB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CD_OPER_ETRT(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CD_OPER_ETRT", OracleType.VarChar);
            Parametro.SourceColumn = "CD_OPER_ETRT";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NR_CNPJ_CPF(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NR_CNPJ_CPF", OracleType.Number);
            Parametro.SourceColumn = "NR_CNPJ_CPF";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NO_CLIE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NO_CLIE", OracleType.VarChar);
            Parametro.SourceColumn = "NO_CLIE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CD_MOED_ISO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CD_MOED_ISO", OracleType.VarChar);
            Parametro.SourceColumn = "CD_MOED_ISO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NR_PERC_TAXA_CAMB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NR_PERC_TAXA_CAMB", OracleType.Number);
            Parametro.SourceColumn = "NR_PERC_TAXA_CAMB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_OPER_CAMB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_OPER_CAMB", OracleType.Char);
            Parametro.SourceColumn = "TP_OPER_CAMB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NR_OPER_CAMB_2(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NR_OPER_CAMB_2", OracleType.Number);
            Parametro.SourceColumn = "NR_OPER_CAMB_2";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TP_NEGO_INTB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_NEGO_INTB", OracleType.Number);
            Parametro.SourceColumn = "TP_NEGO_INTB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_MESG_RECB_ENVI_SPB >>>>>>
        public static OracleParameter NU_CTRL_IF(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_CTRL_IF", OracleType.VarChar);
            Parametro.SourceColumn = "NU_CTRL_IF";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DH_REGT_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_REGT_MESG_SPB", OracleType.DateTime);
            Parametro.SourceColumn = "DH_REGT_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter NU_SEQU_CNTR_REPE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_SEQU_CNTR_REPE", OracleType.Number);
            Parametro.SourceColumn = "NU_SEQU_CNTR_REPE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
     
        public static OracleParameter TP_BKOF(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_BKOF", OracleType.Number);
            Parametro.SourceColumn = "TP_BKOF";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        
        public static OracleParameter DH_RECB_ENVI_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_RECB_ENVI_MESG_SPB", OracleType.DateTime);
            Parametro.SourceColumn = "DH_RECB_ENVI_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
       
        public static OracleParameter CO_SITU_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_SITU_MESG_SPB", OracleType.VarChar);
            Parametro.SourceColumn = "CO_SITU_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
      
        public static OracleParameter NU_CTRL_CAMR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_CTRL_CAMR", OracleType.VarChar);
            Parametro.SourceColumn = "NU_CTRL_CAMR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_MESG_SPB", OracleType.Char);
            Parametro.SourceColumn = "CO_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter IN_CONF_MESG_LTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_IN_CONF_MESG_LTR", OracleType.Number);
            Parametro.SourceColumn = "IN_CONF_MESG_LTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
      
        public static OracleParameter NU_PRTC_MESG_LG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_NU_PRTC_MESG_LG", OracleType.Number);
            Parametro.SourceColumn = "NU_PRTC_MESG_LG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
      
        public static OracleParameter TP_ACAO_MESG_SPB_EXEC(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_ACAO_MESG_SPB_EXEC", OracleType.Number);
            Parametro.SourceColumn = "TP_ACAO_MESG_SPB_EXEC";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ISPB_PART_CAMR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ISPB_PART_CAMR", OracleType.VarChar);
            Parametro.SourceColumn = "CO_ISPB_PART_CAMR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DT_OPER_CAMB_SISBACEN(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DT_OPER_CAMB_SISBACEN", OracleType.DateTime);
            Parametro.SourceColumn = "DT_OPER_CAMB_SISBACEN";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CD_CLIE_SISBACEN(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CD_CLIE_SISBACEN", OracleType.VarChar);
            Parametro.SourceColumn = "CD_CLIE_SISBACEN";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter VL_MOED_ESTR(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_VL_MOED_ESTR", OracleType.Number);
            Parametro.SourceColumn = "VL_MOED_ESTR";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_HIST_SITU_ACAO_MESG_SPB >>>>>>
        public static OracleParameter DH_SITU_ACAO_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_SITU_ACAO_MESG_SPB", OracleType.DateTime);
            Parametro.SourceColumn = "DH_SITU_ACAO_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
     
        public static OracleParameter TP_ACAO_MESG_SPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_ACAO_MESG_SPB", OracleType.Number);
            Parametro.SourceColumn = "TP_ACAO_MESG_SPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
       
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_HIST_ENVI_INFO_GEST_CAIX >>>>>>
        public static OracleParameter DH_ENVI_GEST_CAIX(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_ENVI_GEST_CAIX", OracleType.DateTime);
            Parametro.SourceColumn = "DH_ENVI_GEST_CAIX";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }

        public static OracleParameter CO_SITU_MOVI_GEST_CAIX(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_SITU_MOVI_GEST_CAIX", OracleType.Number);
            Parametro.SourceColumn = "CO_SITU_MOVI_GEST_CAIX";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_REME_REJE >>>>>>
        public static OracleParameter SG_SIST_ORIG_INFO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_SG_SIST_ORIG_INFO", OracleType.Char);
            Parametro.SourceColumn = "SG_SIST_ORIG_INFO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_TEXT_XML_REJE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_TEXT_XML_REJE", OracleType.Number);
            Parametro.SourceColumn = "CO_TEXT_XML_REJE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_TEXT_XML_RETN_SIST_ORIG(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_TEXT_XML_RETN_SIST_ORIG", OracleType.Number);
            Parametro.SourceColumn = "CO_TEXT_XML_RETN_SIST_ORIG";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TX_XML_ERRO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TX_XML_ERRO", OracleType.VarChar);
            Parametro.SourceColumn = "TX_XML_ERRO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter DH_REME_REJE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_DH_REME_REJE", OracleType.DateTime);
            Parametro.SourceColumn = "DH_REME_REJE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }

        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_TIPO_OPER >>>>>>
        public static OracleParameter TP_MESG_RECB_INTE(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_MESG_RECB_INTE", OracleType.VarChar);
            Parametro.SourceColumn = "TP_MESG_RECB_INTE";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_PRODUTO >>>>>>
        public static OracleParameter CO_PROD(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_PROD", OracleType.Number);
            Parametro.SourceColumn = "CO_PROD";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter VA_MINI_MAIR_VALO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_VA_MINI_MAIR_VALO", OracleType.Number);
            Parametro.SourceColumn = "VA_MINI_MAIR_VALO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros das Stores Procedures TB_JUST_CNCL_OPER_ATIV_MESG >>>>>>
        public static OracleParameter TP_JUST_CNCL(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TP_JUST_CNCL", OracleType.Number);
            Parametro.SourceColumn = "TP_JUST_CNCL";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter TX_JUST(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_TX_JUST", OracleType.VarChar);
            Parametro.SourceColumn = "TX_JUST";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ETCA_TRAB_ATLZ(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("P_CO_ETCA_TRAB_ATLZ", OracleType.VarChar);
            Parametro.SourceColumn = "CO_ETCA_TRAB_ATLZ";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ULTI_SITU_PROC_MSGSPB(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_CO_ULTI_SITU_PROC_MSGSPB", OracleType.Number);
            Parametro.SourceColumn = "CO_ULTI_SITU_PROC_MSGSPB";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        public static OracleParameter CO_ULTI_SITU_PROC_OPERACAO(object value, ParameterDirection direcao)
        {
            OracleParameter Parametro = new OracleParameter("p_CO_ULTI_SITU_PROC_OPERACAO", OracleType.Number);
            Parametro.SourceColumn = "CO_ULTI_SITU_PROC_OPERACAO";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }
        #endregion

        #region >>>>>> Parametros Genericos >>>>>>

        public static OracleParameter SISTEMA(object value, ParameterDirection direcao)
        {
            //Parametro Retorno Number
            OracleParameter Parametro = new OracleParameter("PSISTEMA", OracleType.VarChar);
            Parametro.SourceColumn = "PSISTEMA";
            Parametro.Value = value;
            Parametro.Direction = direcao;
            return Parametro;
        }

        //public static OracleParameter TIPO_MENSAGEM(object value, ParameterDirection direcao)
        //{
        //    OracleParameter Parametro = new OracleParameter("p_num_TIPO_MENSAGEM", OracleType.Number);
        //    Parametro.Value = value;
        //    Parametro.Direction = direcao;
        //    return Parametro;
        //}

        //public static OracleParameter CD_ACAO(object value, ParameterDirection direcao)
        //{
        //    OracleParameter Parametro = new OracleParameter("p_num_CD_ACAO", OracleType.Number);
        //    Parametro.Value = value;
        //    Parametro.Direction = direcao;
        //    return Parametro;
        //}

        //public static OracleParameter TP_HIST(object value, ParameterDirection direcao)
        //{
        //    OracleParameter Parametro = new OracleParameter("p_num_TP_HIST", OracleType.Number);
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

        public static OracleParameter QUANTIDADE()
        {
            //Parametro Retorno Number
            OracleParameter Parametro = new OracleParameter("p_num_out_QUANTIDADE", OracleType.Number);
            Parametro.Direction = ParameterDirection.Output;
            Parametro.OracleType = OracleType.Number;
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

        public static OracleParameter SEQUENCIA(object value, ParameterDirection direcao)
        {
            //Parametro Retorno Number
            OracleParameter Parametro = new OracleParameter("PSEQUENCIA", OracleType.VarChar, 20);
            Parametro.Direction = direcao;
            Parametro.Value = value;
            //Parametro.Value = value;
            //Parametro.OracleType = OracleType.VarChar;
            return Parametro;
        }
        public static OracleParameter RETORNO2()
        {
            //Parametro Retorno Number
            OracleParameter Parametro = new OracleParameter("p_RETORNO", OracleType.Number);
            Parametro.Direction = ParameterDirection.Output;
            Parametro.OracleType = OracleType.Number;
            return Parametro;
        }
        public static OracleParameter RETORNO_NU_SEQU_CNCL_OPER_ATIV_MESG()
        {
            //Parametro Retorno Number
            OracleParameter Parametro = new OracleParameter("p_NU_SEQU_CNCL_OPER_ATIV_MESG", OracleType.Number);
            Parametro.Direction = ParameterDirection.Output;
            Parametro.OracleType = OracleType.Number;
            return Parametro;
        }
        #endregion
    }
}
