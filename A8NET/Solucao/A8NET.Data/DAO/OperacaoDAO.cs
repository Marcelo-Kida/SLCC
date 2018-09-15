using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;
using System.ComponentModel;
using System.Collections;
using System.Xml;

namespace A8NET.Data.DAO
{
    public class OperacaoDAO : BaseDAO
    {
        
        #region <<< Estrutura >>>
        public partial class EstruturaOperacao
        {
            #region <<< Variáveis internas (Campos da tabela) >>>
            public System.Collections.Hashtable HashNovasTAGs = new Hashtable();
            private object _NU_SEQU_OPER_ATIV;
            private object _TP_OPER;
            private object _CO_LOCA_LIQU;
            private object _TP_LIQU_OPER_ATIV;
            private object _CO_EMPR;
            private object _CO_USUA_CADR_OPER;
            private object _HO_ENVI_MESG_SPB;
            private object _CO_OPER_ATIV;
            private object _NU_COMD_OPER;
            private object _NU_COMD_OPER_RETN;
            private object _DT_OPER_ATIV;
            private object _DT_OPER_ATIV_RETN;
            private object _CO_VEIC_LEGA;
            private object _SG_SIST;
            private object _CO_CNTA_CUTD_SELIC_VEIC_LEGA;
            private object _CO_CNPJ_CNPT;
            private object _CO_CNTA_CUTD_SELIC_CNPT;
            private object _NO_CNPT;
            private object _IN_OPER_DEBT_CRED;
            private object _NU_ATIV_MERC;
            private object _DE_ATIV_MERC;
            private object _PU_ATIV_MERC;
            private object _QT_ATIV_MERC;
            private object _IN_ENTR_SAID_RECU_FINC;
            private object _DT_VENC_ATIV;
            private object _VA_OPER_ATIV;
            private object _VA_OPER_ATIV_REAJ;
            private object _DT_LIQU_OPER_ATIV;
            private object _TP_CPRO_OPER_ATIV;
            private object _TP_CPRO_RETN_OPER_ATIV;
            private object _IN_DISP_CONS;
            private object _IN_ENVI_PREV_SIST_PJ;
            private object _IN_ENVI_RELZ_SIST_PJ;
            private object _IN_ENVI_PREV_SIST_A6;
            private object _IN_ENVI_RELZ_SIST_A6;
            private object _CO_ULTI_SITU_PROC;
            private object _TP_ACAO_OPER_ATIV_EXEC;
            private object _NU_COMD_ACAO_EXEC;
            private object _IN_ENTR_MANU;
            private object _NU_PRTC_OPER_LG;
            private object _NU_SEQU_CNCL_OPER_ATIV_MESG;
            private object _NU_CTRL_MESG_SPB_ORIG;
            private object _PE_TAXA_NEGO;
            private object _CO_TITL_CUTD;
            private object _CO_OPER_CETIP;
            private object _CO_ISPB_BANC_LIQU_CNPT;
            private object _DH_ULTI_ATLZ;
            private object _CO_ETCA_TRAB_ULTI_ATLZ;
            private object _CO_USUA_ULTI_ATLZ;
            private object _TP_IF_CRED_DEBT;
            private object _CO_AGEN_COTR;
            private object _NU_CC_COTR;
            private object _PZ_DIAS_RETN_OPER_ATIV;
            private object _VA_OPER_ATIV_RETN;
            private object _TP_CNPT;
            private object _CO_CNPT_CAMR;
            private object _CO_IDEF_LAST;
            private object _CO_PARP_CAMR;
            private object _TP_PGTO_LDL;
            private object _CO_GRUP_LANC_FINC;
            private object _CO_MOED_ESTR;
            private object _CO_CNTR_SISB;
            private object _CO_ISPB_IF_CNPT;
            private object _CO_PRAC;
            private object _VA_MOED_ESTR;
            private object _DT_LIQU_OPER_ATIV_MOED_ESTR;
            private object _CO_SISB_COTR;
            private object _CO_CNAL_OPER_INTE;
            private object _CO_SITU_PROC_MESG_SPB_RECB;
            private object _TP_CNAL_VEND;
            private object _CD_SUB_PROD;
            private object _NR_IDEF_NEGO_BMC;
            private object _TP_NEGO;
            private object _CD_ASSO_CAMB;
            private object _CD_OPER_ETRT;
            private object _NR_CNPJ_CPF;
            private object _TP_NEGO_INTB;
            private object _NR_OPER_CAMB_2;
            #endregion

            #region <<<  Propriedades (Campos da Tabela) >>>
            // Propriedade que representa o campo: NU_SEQU_OPER_ATIV
            public object NU_SEQU_OPER_ATIV
            {
                get
                {
                    return _NU_SEQU_OPER_ATIV;
                }
                set
                {
                    _NU_SEQU_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: TP_OPER
            public object TP_OPER
            {
                get
                {
                    return _TP_OPER;
                }
                set
                {
                    _TP_OPER = value;
                }
            }
            // Propriedade que representa o campo: CO_LOCA_LIQU
            public object CO_LOCA_LIQU
            {
                get
                {
                    return _CO_LOCA_LIQU;
                }
                set
                {
                    _CO_LOCA_LIQU = value;
                }
            }
            // Propriedade que representa o campo: TP_LIQU_OPER_ATIV
            public object TP_LIQU_OPER_ATIV
            {
                get
                {
                    return _TP_LIQU_OPER_ATIV;
                }
                set
                {
                    _TP_LIQU_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: CO_EMPR
            public object CO_EMPR
            {
                get
                {
                    return _CO_EMPR;
                }
                set
                {
                    _CO_EMPR = value;
                }
            }
            // Propriedade que representa o campo: CO_USUA_CADR_OPER
            public object CO_USUA_CADR_OPER
            {
                get
                {
                    return _CO_USUA_CADR_OPER;
                }
                set
                {
                    _CO_USUA_CADR_OPER = value;
                }
            }
            // Propriedade que representa o campo: HO_ENVI_MESG_SPB
            public object HO_ENVI_MESG_SPB
            {
                get
                {
                    return _HO_ENVI_MESG_SPB;
                }
                set
                {
                    _HO_ENVI_MESG_SPB = value;
                }
            }
            // Propriedade que representa o campo: CO_OPER_ATIV
            public object CO_OPER_ATIV
            {
                get
                {
                    return _CO_OPER_ATIV;
                }
                set
                {
                    _CO_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: NU_COMD_OPER
            public object NU_COMD_OPER
            {
                get
                {
                    return _NU_COMD_OPER;
                }
                set
                {
                    _NU_COMD_OPER = value;
                }
            }
            // Propriedade que representa o campo: NU_COMD_OPER_RETN
            public object NU_COMD_OPER_RETN
            {
                get
                {
                    return _NU_COMD_OPER_RETN;
                }
                set
                {
                    _NU_COMD_OPER_RETN = value;
                }
            }
            // Propriedade que representa o campo: DT_OPER_ATIV
            public object DT_OPER_ATIV
            {
                get
                {
                    return _DT_OPER_ATIV;
                }
                set
                {
                    _DT_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: DT_OPER_ATIV_RETN
            public object DT_OPER_ATIV_RETN
            {
                get
                {
                    return _DT_OPER_ATIV_RETN;
                }
                set
                {
                    _DT_OPER_ATIV_RETN = value;
                }
            }
            // Propriedade que representa o campo: CO_VEIC_LEGA
            public object CO_VEIC_LEGA
            {
                get
                {
                    return _CO_VEIC_LEGA;
                }
                set
                {
                    _CO_VEIC_LEGA = value;
                }
            }
            // Propriedade que representa o campo: SG_SIST
            public object SG_SIST
            {
                get
                {
                    return _SG_SIST;
                }
                set
                {
                    _SG_SIST = value;
                }
            }
            // Propriedade que representa o campo: CO_CNTA_CUTD_SELIC_VEIC_LEGA
            public object CO_CNTA_CUTD_SELIC_VEIC_LEGA
            {
                get
                {
                    return _CO_CNTA_CUTD_SELIC_VEIC_LEGA;
                }
                set
                {
                    _CO_CNTA_CUTD_SELIC_VEIC_LEGA = value;
                }
            }
            // Propriedade que representa o campo: CO_CNPJ_CNPT
            public object CO_CNPJ_CNPT
            {
                get
                {
                    return _CO_CNPJ_CNPT;
                }
                set
                {
                    _CO_CNPJ_CNPT = value;
                }
            }
            // Propriedade que representa o campo: CO_CNTA_CUTD_SELIC_CNPT
            public object CO_CNTA_CUTD_SELIC_CNPT
            {
                get
                {
                    return _CO_CNTA_CUTD_SELIC_CNPT;
                }
                set
                {
                    _CO_CNTA_CUTD_SELIC_CNPT = value;
                }
            }
            // Propriedade que representa o campo: NO_CNPT
            public object NO_CNPT
            {
                get
                {
                    return _NO_CNPT;
                }
                set
                {
                    _NO_CNPT = value;
                }
            }
            // Propriedade que representa o campo: IN_OPER_DEBT_CRED
            public object IN_OPER_DEBT_CRED
            {
                get
                {
                    return _IN_OPER_DEBT_CRED;
                }
                set
                {
                    _IN_OPER_DEBT_CRED = value;
                }
            }
            // Propriedade que representa o campo: NU_ATIV_MERC
            public object NU_ATIV_MERC
            {
                get
                {
                    return _NU_ATIV_MERC;
                }
                set
                {
                    _NU_ATIV_MERC = value;
                }
            }
            // Propriedade que representa o campo: DE_ATIV_MERC
            public object DE_ATIV_MERC
            {
                get
                {
                    return _DE_ATIV_MERC;
                }
                set
                {
                    _DE_ATIV_MERC = value;
                }
            }
            // Propriedade que representa o campo: PU_ATIV_MERC
            public object PU_ATIV_MERC
            {
                get
                {
                    return _PU_ATIV_MERC;
                }
                set
                {
                    _PU_ATIV_MERC = value;
                }
            }
            // Propriedade que representa o campo: QT_ATIV_MERC
            public object QT_ATIV_MERC
            {
                get
                {
                    return _QT_ATIV_MERC;
                }
                set
                {
                    _QT_ATIV_MERC = value;
                }
            }
            // Propriedade que representa o campo: IN_ENTR_SAID_RECU_FINC
            public object IN_ENTR_SAID_RECU_FINC
            {
                get
                {
                    return _IN_ENTR_SAID_RECU_FINC;
                }
                set
                {
                    _IN_ENTR_SAID_RECU_FINC = value;
                }
            }
            // Propriedade que representa o campo: DT_VENC_ATIV
            public object DT_VENC_ATIV
            {
                get
                {
                    return _DT_VENC_ATIV;
                }
                set
                {
                    _DT_VENC_ATIV = value;
                }
            }
            // Propriedade que representa o campo: VA_OPER_ATIV
            public object VA_OPER_ATIV
            {
                get
                {
                    return _VA_OPER_ATIV;
                }
                set
                {
                    _VA_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: VA_OPER_ATIV_REAJ
            public object VA_OPER_ATIV_REAJ
            {
                get
                {
                    return _VA_OPER_ATIV_REAJ;
                }
                set
                {
                    _VA_OPER_ATIV_REAJ = value;
                }
            }
            // Propriedade que representa o campo: DT_LIQU_OPER_ATIV
            public object DT_LIQU_OPER_ATIV
            {
                get
                {
                    return _DT_LIQU_OPER_ATIV;
                }
                set
                {
                    _DT_LIQU_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: TP_CPRO_OPER_ATIV
            public object TP_CPRO_OPER_ATIV
            {
                get
                {
                    return _TP_CPRO_OPER_ATIV;
                }
                set
                {
                    _TP_CPRO_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: TP_CPRO_RETN_OPER_ATIV
            public object TP_CPRO_RETN_OPER_ATIV
            {
                get
                {
                    return _TP_CPRO_RETN_OPER_ATIV;
                }
                set
                {
                    _TP_CPRO_RETN_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: IN_DISP_CONS
            public object IN_DISP_CONS
            {
                get
                {
                    return _IN_DISP_CONS;
                }
                set
                {
                    _IN_DISP_CONS = value;
                }
            }
            // Propriedade que representa o campo: IN_ENVI_PREV_SIST_PJ
            public object IN_ENVI_PREV_SIST_PJ
            {
                get
                {
                    return _IN_ENVI_PREV_SIST_PJ;
                }
                set
                {
                    _IN_ENVI_PREV_SIST_PJ = value;
                }
            }
            // Propriedade que representa o campo: IN_ENVI_RELZ_SIST_PJ
            public object IN_ENVI_RELZ_SIST_PJ
            {
                get
                {
                    return _IN_ENVI_RELZ_SIST_PJ;
                }
                set
                {
                    _IN_ENVI_RELZ_SIST_PJ = value;
                }
            }
            // Propriedade que representa o campo: IN_ENVI_PREV_SIST_A6
            public object IN_ENVI_PREV_SIST_A6
            {
                get
                {
                    return _IN_ENVI_PREV_SIST_A6;
                }
                set
                {
                    _IN_ENVI_PREV_SIST_A6 = value;
                }
            }
            // Propriedade que representa o campo: IN_ENVI_RELZ_SIST_A6
            public object IN_ENVI_RELZ_SIST_A6
            {
                get
                {
                    return _IN_ENVI_RELZ_SIST_A6;
                }
                set
                {
                    _IN_ENVI_RELZ_SIST_A6 = value;
                }
            }
            // Propriedade que representa o campo: CO_ULTI_SITU_PROC
            public object CO_ULTI_SITU_PROC
            {
                get
                {
                    return _CO_ULTI_SITU_PROC;
                }
                set
                {
                    _CO_ULTI_SITU_PROC = value;
                }
            }
            // Propriedade que representa o campo: TP_ACAO_OPER_ATIV_EXEC
            public object TP_ACAO_OPER_ATIV_EXEC
            {
                get
                {
                    return _TP_ACAO_OPER_ATIV_EXEC;
                }
                set
                {
                    _TP_ACAO_OPER_ATIV_EXEC = value;
                }
            }
            // Propriedade que representa o campo: NU_COMD_ACAO_EXEC
            public object NU_COMD_ACAO_EXEC
            {
                get
                {
                    return _NU_COMD_ACAO_EXEC;
                }
                set
                {
                    _NU_COMD_ACAO_EXEC = value;
                }
            }
            // Propriedade que representa o campo: IN_ENTR_MANU
            public object IN_ENTR_MANU
            {
                get
                {
                    return _IN_ENTR_MANU;
                }
                set
                {
                    _IN_ENTR_MANU = value;
                }
            }
            // Propriedade que representa o campo: NU_PRTC_OPER_LG
            public object NU_PRTC_OPER_LG
            {
                get
                {
                    return _NU_PRTC_OPER_LG;
                }
                set
                {
                    _NU_PRTC_OPER_LG = value;
                }
            }
            // Propriedade que representa o campo: NU_SEQU_CNCL_OPER_ATIV_MESG
            public object NU_SEQU_CNCL_OPER_ATIV_MESG
            {
                get
                {
                    return _NU_SEQU_CNCL_OPER_ATIV_MESG;
                }
                set
                {
                    _NU_SEQU_CNCL_OPER_ATIV_MESG = value;
                }
            }
            // Propriedade que representa o campo: NU_CTRL_MESG_SPB_ORIG
            public object NU_CTRL_MESG_SPB_ORIG
            {
                get
                {
                    return _NU_CTRL_MESG_SPB_ORIG;
                }
                set
                {
                    _NU_CTRL_MESG_SPB_ORIG = value;
                }
            }
            // Propriedade que representa o campo: PE_TAXA_NEGO
            public object PE_TAXA_NEGO
            {
                get
                {
                    return _PE_TAXA_NEGO;
                }
                set
                {
                    _PE_TAXA_NEGO = value;
                }
            }
            // Propriedade que representa o campo: CO_TITL_CUTD
            public object CO_TITL_CUTD
            {
                get
                {
                    return _CO_TITL_CUTD;
                }
                set
                {
                    _CO_TITL_CUTD = value;
                }
            }
            // Propriedade que representa o campo: CO_OPER_CETIP
            public object CO_OPER_CETIP
            {
                get
                {
                    return _CO_OPER_CETIP;
                }
                set
                {
                    _CO_OPER_CETIP = value;
                }
            }
            // Propriedade que representa o campo: CO_ISPB_BANC_LIQU_CNPT
            public object CO_ISPB_BANC_LIQU_CNPT
            {
                get
                {
                    return _CO_ISPB_BANC_LIQU_CNPT;
                }
                set
                {
                    _CO_ISPB_BANC_LIQU_CNPT = value;
                }
            }
            // Propriedade que representa o campo: DH_ULTI_ATLZ
            public object DH_ULTI_ATLZ
            {
                get
                {
                    return _DH_ULTI_ATLZ;
                }
                set
                {
                    _DH_ULTI_ATLZ = value;
                }
            }
            // Propriedade que representa o campo: CO_ETCA_TRAB_ULTI_ATLZ
            public object CO_ETCA_TRAB_ULTI_ATLZ
            {
                get
                {
                    return _CO_ETCA_TRAB_ULTI_ATLZ;
                }
                set
                {
                    _CO_ETCA_TRAB_ULTI_ATLZ = value;
                }
            }
            // Propriedade que representa o campo: CO_USUA_ULTI_ATLZ
            public object CO_USUA_ULTI_ATLZ
            {
                get
                {
                    return _CO_USUA_ULTI_ATLZ;
                }
                set
                {
                    _CO_USUA_ULTI_ATLZ = value;
                }
            }
            // Propriedade que representa o campo: TP_IF_CRED_DEBT
            public object TP_IF_CRED_DEBT
            {
                get
                {
                    return _TP_IF_CRED_DEBT;
                }
                set
                {
                    _TP_IF_CRED_DEBT = value;
                }
            }
            // Propriedade que representa o campo: CO_AGEN_COTR
            public object CO_AGEN_COTR
            {
                get
                {
                    return _CO_AGEN_COTR;
                }
                set
                {
                    _CO_AGEN_COTR = value;
                }
            }
            // Propriedade que representa o campo: NU_CC_COTR
            public object NU_CC_COTR
            {
                get
                {
                    return _NU_CC_COTR;
                }
                set
                {
                    _NU_CC_COTR = value;
                }
            }
            // Propriedade que representa o campo: PZ_DIAS_RETN_OPER_ATIV
            public object PZ_DIAS_RETN_OPER_ATIV
            {
                get
                {
                    return _PZ_DIAS_RETN_OPER_ATIV;
                }
                set
                {
                    _PZ_DIAS_RETN_OPER_ATIV = value;
                }
            }
            // Propriedade que representa o campo: VA_OPER_ATIV_RETN
            public object VA_OPER_ATIV_RETN
            {
                get
                {
                    return _VA_OPER_ATIV_RETN;
                }
                set
                {
                    _VA_OPER_ATIV_RETN = value;
                }
            }
            // Propriedade que representa o campo: TP_CNPT
            public object TP_CNPT
            {
                get
                {
                    return _TP_CNPT;
                }
                set
                {
                    _TP_CNPT = value;
                }
            }
            // Propriedade que representa o campo: CO_CNPT_CAMR
            public object CO_CNPT_CAMR
            {
                get
                {
                    return _CO_CNPT_CAMR;
                }
                set
                {
                    _CO_CNPT_CAMR = value;
                }
            }
            // Propriedade que representa o campo: CO_IDEF_LAST
            public object CO_IDEF_LAST
            {
                get
                {
                    return _CO_IDEF_LAST;
                }
                set
                {
                    _CO_IDEF_LAST = value;
                }
            }
            // Propriedade que representa o campo: CO_PARP_CAMR
            public object CO_PARP_CAMR
            {
                get
                {
                    return _CO_PARP_CAMR;
                }
                set
                {
                    _CO_PARP_CAMR = value;
                }
            }
            // Propriedade que representa o campo: TP_PGTO_LDL
            public object TP_PGTO_LDL
            {
                get
                {
                    return _TP_PGTO_LDL;
                }
                set
                {
                    _TP_PGTO_LDL = value;
                }
            }
            // Propriedade que representa o campo: CO_GRUP_LANC_FINC
            public object CO_GRUP_LANC_FINC
            {
                get
                {
                    return _CO_GRUP_LANC_FINC;
                }
                set
                {
                    _CO_GRUP_LANC_FINC = value;
                }
            }
            // Propriedade que representa o campo: CO_MOED_ESTR
            public object CO_MOED_ESTR
            {
                get
                {
                    return _CO_MOED_ESTR;
                }
                set
                {
                    _CO_MOED_ESTR = value;
                }
            }
            // Propriedade que representa o campo: CO_CNTR_SISB
            public object CO_CNTR_SISB
            {
                get
                {
                    return _CO_CNTR_SISB;
                }
                set
                {
                    _CO_CNTR_SISB = value;
                }
            }
            // Propriedade que representa o campo: CO_ISPB_IF_CNPT
            public object CO_ISPB_IF_CNPT
            {
                get
                {
                    return _CO_ISPB_IF_CNPT;
                }
                set
                {
                    _CO_ISPB_IF_CNPT = value;
                }
            }
            // Propriedade que representa o campo: CO_PRAC
            public object CO_PRAC
            {
                get
                {
                    return _CO_PRAC;
                }
                set
                {
                    _CO_PRAC = value;
                }
            }
            // Propriedade que representa o campo: VA_MOED_ESTR
            public object VA_MOED_ESTR
            {
                get
                {
                    return _VA_MOED_ESTR;
                }
                set
                {
                    _VA_MOED_ESTR = value;
                }
            }
            // Propriedade que representa o campo: DT_LIQU_OPER_ATIV_MOED_ESTR
            public object DT_LIQU_OPER_ATIV_MOED_ESTR
            {
                get
                {
                    return _DT_LIQU_OPER_ATIV_MOED_ESTR;
                }
                set
                {
                    _DT_LIQU_OPER_ATIV_MOED_ESTR = value;
                }
            }
            // Propriedade que representa o campo: CO_SISB_COTR
            public object CO_SISB_COTR
            {
                get
                {
                    return _CO_SISB_COTR;
                }
                set
                {
                    _CO_SISB_COTR = value;
                }
            }
            // Propriedade que representa o campo: CO_CNAL_OPER_INTE
            public object CO_CNAL_OPER_INTE
            {
                get
                {
                    return _CO_CNAL_OPER_INTE;
                }
                set
                {
                    _CO_CNAL_OPER_INTE = value;
                }
            }
            // Propriedade que representa o campo: CO_SITU_PROC_MESG_SPB_RECB
            public object CO_SITU_PROC_MESG_SPB_RECB
            {
                get
                {
                    return _CO_SITU_PROC_MESG_SPB_RECB;
                }
                set
                {
                    _CO_SITU_PROC_MESG_SPB_RECB = value;
                }
            }
            // Propriedade que representa o campo: TP_CNAL_VEND
            public object TP_CNAL_VEND
            {
                get
                {
                    return _TP_CNAL_VEND;
                }
                set
                {
                    _TP_CNAL_VEND = value;
                }
            }
            // Propriedade que representa o campo: CD_SUB_PROD
            public object CD_SUB_PROD
            {
                get
                {
                    return _CD_SUB_PROD;
                }
                set
                {
                    _CD_SUB_PROD = value;
                }
            }
            // Propriedade que representa o campo: NR_IDEF_NEGO_BMC
            public object NR_IDEF_NEGO_BMC
            {
                get
                {
                    return _NR_IDEF_NEGO_BMC;
                }
                set
                {
                    _NR_IDEF_NEGO_BMC = value;
                }
            }
            // Propriedade que representa o campo: TP_NEGO
            public object TP_NEGO
            {
                get
                {
                    return _TP_NEGO;
                }
                set
                {
                    _TP_NEGO = value;
                }
            }
            // Propriedade que representa o campo: CD_ASSO_CAMB
            public object CD_ASSO_CAMB
            {
                get
                {
                    return _CD_ASSO_CAMB;
                }
                set
                {
                    _CD_ASSO_CAMB = value;
                }
            }
            // Propriedade que representa o campo: CD_OPER_ETRT
            public object CD_OPER_ETRT
            {
                get
                {
                    return _CD_OPER_ETRT;
                }
                set
                {
                    _CD_OPER_ETRT = value;
                }
            }
            // Propriedade que representa o campo: NR_CNPJ_CPF
            public object NR_CNPJ_CPF
            {
                get
                {
                    return _NR_CNPJ_CPF;
                }
                set
                {
                    _NR_CNPJ_CPF = value;
                }
            }
            // Propriedade que representa o campo: NR_OPER_CAMB_2
            public object NR_OPER_CAMB_2
            {
                get
                {
                    return _NR_OPER_CAMB_2;
                }
                set
                {
                    _NR_OPER_CAMB_2 = value;
                }
            }
            // Propriedade que representa o campo: TP_NEGO_INTB
            public object TP_NEGO_INTB
            {
                get
                {
                    return _TP_NEGO_INTB;
                }
                set
                {
                    _TP_NEGO_INTB = value;
                }
            }
            #endregion

            #region <<< OVERRRIDE - ToString() >>>
            public override string ToString()
            {
                //string CamposFormamToString = "|TP_OPER|NU_SEQU_OPER_ATIV|CO_ULTI_SITU_PROC|DT_OPER_ATIV|IN_ENTR_MANU|HO_ENVI_MESG_SPB|TP_ACAO_OPER_ATIV_EXEC|NU_COMD_ACAO_EXEC|DH_ULTI_ATLZ|TP_CPRO_OPER_ATIV|TP_CPRO_RETN_OPER_ATIV|VA_OPER_ATIV|NU_CTRL_MESG_SPB_ORIG|DT_OPER_ATIV|CO_CNTA_CUTD_SELIC_CNPT|NU_PRTC_OPER_LG|CO_PRAC|CO_MOED_ESTR|PE_TAXA_NEGO|VA_MOED_ESTR|DT_LIQU_OPER_ATIV|CO_SISB_COTR|NU_COMD_OPER|TP_CNAL_VEND|CD_SUB_PROD|B.CO_TEXT_XML|C.CO_MESG_SPB_REGT_OPER|NU_ATIV_MERC|CD_ASSO_CAMB|";
                StringBuilder Append = new StringBuilder();
                string Formato = "<{0}>{1}</{0}>";
                Append.Append("<OPER>");
                foreach (PropertyDescriptor propriedade in TypeDescriptor.GetProperties(this))
                {
                    Append.AppendFormat(Formato, propriedade.Name, propriedade.GetValue(this));
                }
                foreach (DictionaryEntry Item in HashNovasTAGs)
                {
                    Append.AppendFormat(Formato, Item.Key, Item.Value); 
                }
                Append.Append("</OPER>");
                return Append.ToString();
            }
            #endregion
        }
        #endregion

        #region <<< Propriedades >>>
        public EstruturaOperacao[] Itens
        {
            get
            {
                //Popula
                EstruturaOperacao[] lOperacao = null;
                Int32 lI;

                lOperacao = new EstruturaOperacao[_Lista.Count];

                for (lI = 0; lI < _Lista.Count; lI++)
                {
                    lOperacao[lI] = (EstruturaOperacao)_Lista[lI];
                }
                return lOperacao;
            }
        }
        public EstruturaOperacao TB_OPER_ATIV;
        #endregion

        #region <<< ProcessarOperacao >>>
        private void ProcessarOperacao(System.Data.DataView lDView)
        {
            Int32 lI;

            //Processa Lista
            _Lista.Clear();
            for (lI = 0; lI < lDView.Count; lI++)
            {
                EstruturaOperacao lOperacao = new EstruturaOperacao();

                lOperacao.NU_SEQU_OPER_ATIV = lDView[lI].Row["NU_SEQU_OPER_ATIV"];
                lOperacao.TP_OPER = lDView[lI].Row["TP_OPER"];
                lOperacao.CO_LOCA_LIQU = lDView[lI].Row["CO_LOCA_LIQU"];
                lOperacao.TP_LIQU_OPER_ATIV = lDView[lI].Row["TP_LIQU_OPER_ATIV"];
                lOperacao.CO_EMPR = lDView[lI].Row["CO_EMPR"];
                lOperacao.CO_USUA_CADR_OPER = lDView[lI].Row["CO_USUA_CADR_OPER"];
                lOperacao.HO_ENVI_MESG_SPB = lDView[lI].Row["HO_ENVI_MESG_SPB"];
                lOperacao.CO_OPER_ATIV = lDView[lI].Row["CO_OPER_ATIV"];
                lOperacao.NU_COMD_OPER = lDView[lI].Row["NU_COMD_OPER"];
                lOperacao.NU_COMD_OPER_RETN = lDView[lI].Row["NU_COMD_OPER_RETN"];
                lOperacao.DT_OPER_ATIV = lDView[lI].Row["DT_OPER_ATIV"];
                lOperacao.DT_OPER_ATIV_RETN = lDView[lI].Row["DT_OPER_ATIV_RETN"];
                lOperacao.CO_VEIC_LEGA = lDView[lI].Row["CO_VEIC_LEGA"];
                lOperacao.SG_SIST = lDView[lI].Row["SG_SIST"];
                lOperacao.CO_CNTA_CUTD_SELIC_VEIC_LEGA = lDView[lI].Row["CO_CNTA_CUTD_SELIC_VEIC_LEGA"];
                lOperacao.CO_CNPJ_CNPT = lDView[lI].Row["CO_CNPJ_CNPT"];
                lOperacao.CO_CNTA_CUTD_SELIC_CNPT = lDView[lI].Row["CO_CNTA_CUTD_SELIC_CNPT"];
                lOperacao.NO_CNPT = lDView[lI].Row["NO_CNPT"];
                lOperacao.IN_OPER_DEBT_CRED = lDView[lI].Row["IN_OPER_DEBT_CRED"];
                lOperacao.NU_ATIV_MERC = lDView[lI].Row["NU_ATIV_MERC"];
                lOperacao.DE_ATIV_MERC = lDView[lI].Row["DE_ATIV_MERC"];
                lOperacao.PU_ATIV_MERC = lDView[lI].Row["PU_ATIV_MERC"];
                lOperacao.QT_ATIV_MERC = lDView[lI].Row["QT_ATIV_MERC"];
                lOperacao.IN_ENTR_SAID_RECU_FINC = lDView[lI].Row["IN_ENTR_SAID_RECU_FINC"];
                lOperacao.DT_VENC_ATIV = lDView[lI].Row["DT_VENC_ATIV"];
                lOperacao.VA_OPER_ATIV = lDView[lI].Row["VA_OPER_ATIV"];
                lOperacao.VA_OPER_ATIV_REAJ = lDView[lI].Row["VA_OPER_ATIV_REAJ"];
                lOperacao.DT_LIQU_OPER_ATIV = lDView[lI].Row["DT_LIQU_OPER_ATIV"];
                lOperacao.TP_CPRO_OPER_ATIV = lDView[lI].Row["TP_CPRO_OPER_ATIV"];
                lOperacao.TP_CPRO_RETN_OPER_ATIV = lDView[lI].Row["TP_CPRO_RETN_OPER_ATIV"];
                lOperacao.IN_DISP_CONS = lDView[lI].Row["IN_DISP_CONS"];
                lOperacao.IN_ENVI_PREV_SIST_PJ = lDView[lI].Row["IN_ENVI_PREV_SIST_PJ"];
                lOperacao.IN_ENVI_RELZ_SIST_PJ = lDView[lI].Row["IN_ENVI_RELZ_SIST_PJ"];
                lOperacao.IN_ENVI_PREV_SIST_A6 = lDView[lI].Row["IN_ENVI_PREV_SIST_A6"];
                lOperacao.IN_ENVI_RELZ_SIST_A6 = lDView[lI].Row["IN_ENVI_RELZ_SIST_A6"];
                lOperacao.CO_ULTI_SITU_PROC = lDView[lI].Row["CO_ULTI_SITU_PROC"];
                lOperacao.TP_ACAO_OPER_ATIV_EXEC = lDView[lI].Row["TP_ACAO_OPER_ATIV_EXEC"];
                lOperacao.NU_COMD_ACAO_EXEC = lDView[lI].Row["NU_COMD_ACAO_EXEC"];
                lOperacao.IN_ENTR_MANU = lDView[lI].Row["IN_ENTR_MANU"];
                lOperacao.NU_PRTC_OPER_LG = lDView[lI].Row["NU_PRTC_OPER_LG"];
                lOperacao.NU_SEQU_CNCL_OPER_ATIV_MESG = lDView[lI].Row["NU_SEQU_CNCL_OPER_ATIV_MESG"];
                lOperacao.NU_CTRL_MESG_SPB_ORIG = lDView[lI].Row["NU_CTRL_MESG_SPB_ORIG"];
                lOperacao.PE_TAXA_NEGO = lDView[lI].Row["PE_TAXA_NEGO"];
                lOperacao.CO_TITL_CUTD = lDView[lI].Row["CO_TITL_CUTD"];
                lOperacao.CO_OPER_CETIP = lDView[lI].Row["CO_OPER_CETIP"];
                lOperacao.CO_ISPB_BANC_LIQU_CNPT = lDView[lI].Row["CO_ISPB_BANC_LIQU_CNPT"];
                lOperacao.DH_ULTI_ATLZ = lDView[lI].Row["DH_ULTI_ATLZ"];
                lOperacao.CO_ETCA_TRAB_ULTI_ATLZ = lDView[lI].Row["CO_ETCA_TRAB_ULTI_ATLZ"];
                lOperacao.CO_USUA_ULTI_ATLZ = lDView[lI].Row["CO_USUA_ULTI_ATLZ"];
                lOperacao.TP_IF_CRED_DEBT = lDView[lI].Row["TP_IF_CRED_DEBT"];
                lOperacao.CO_AGEN_COTR = lDView[lI].Row["CO_AGEN_COTR"];
                lOperacao.NU_CC_COTR = lDView[lI].Row["NU_CC_COTR"];
                lOperacao.PZ_DIAS_RETN_OPER_ATIV = lDView[lI].Row["PZ_DIAS_RETN_OPER_ATIV"];
                lOperacao.VA_OPER_ATIV_RETN = lDView[lI].Row["VA_OPER_ATIV_RETN"];
                lOperacao.TP_CNPT = lDView[lI].Row["TP_CNPT"];
                lOperacao.CO_CNPT_CAMR = lDView[lI].Row["CO_CNPT_CAMR"];
                lOperacao.CO_IDEF_LAST = lDView[lI].Row["CO_IDEF_LAST"];
                lOperacao.CO_PARP_CAMR = lDView[lI].Row["CO_PARP_CAMR"];
                lOperacao.TP_PGTO_LDL = lDView[lI].Row["TP_PGTO_LDL"];
                lOperacao.CO_GRUP_LANC_FINC = lDView[lI].Row["CO_GRUP_LANC_FINC"];
                lOperacao.CO_MOED_ESTR = lDView[lI].Row["CO_MOED_ESTR"];
                lOperacao.CO_CNTR_SISB = lDView[lI].Row["CO_CNTR_SISB"];
                lOperacao.CO_ISPB_IF_CNPT = lDView[lI].Row["CO_ISPB_IF_CNPT"];
                lOperacao.CO_PRAC = lDView[lI].Row["CO_PRAC"];
                lOperacao.VA_MOED_ESTR = lDView[lI].Row["VA_MOED_ESTR"];
                lOperacao.DT_LIQU_OPER_ATIV_MOED_ESTR = lDView[lI].Row["DT_LIQU_OPER_ATIV_MOED_ESTR"];
                lOperacao.CO_SISB_COTR = lDView[lI].Row["CO_SISB_COTR"];
                lOperacao.CO_CNAL_OPER_INTE = lDView[lI].Row["CO_CNAL_OPER_INTE"];
                lOperacao.CO_SITU_PROC_MESG_SPB_RECB = lDView[lI].Row["CO_SITU_PROC_MESG_SPB_RECB"];
                lOperacao.TP_CNAL_VEND = lDView[lI].Row["TP_CNAL_VEND"];
                lOperacao.CD_SUB_PROD = lDView[lI].Row["CD_SUB_PROD"];
                lOperacao.NR_IDEF_NEGO_BMC = lDView[lI].Row["NR_IDEF_NEGO_BMC"];
                lOperacao.TP_NEGO = lDView[lI].Row["TP_NEGO"];
                lOperacao.CD_ASSO_CAMB = lDView[lI].Row["CD_ASSO_CAMB"];
                lOperacao.CD_OPER_ETRT = lDView[lI].Row["CD_OPER_ETRT"];
                lOperacao.NR_CNPJ_CPF = lDView[lI].Row["NR_CNPJ_CPF"];
                lOperacao.NR_OPER_CAMB_2 = lDView[lI].Row["NR_OPER_CAMB_2"];
                lOperacao.TP_NEGO_INTB = lDView[lI].Row["TP_NEGO_INTB"];

                //Adiciona
                _Lista.Add(lOperacao);
            }
            if (_Lista.Count == 1) TB_OPER_ATIV = (EstruturaOperacao)_Lista[0];
            //Define DView Final
            _DView = lDView;
        }
        #endregion

        #region <<< ObterOperacao - OVERLOAD NU_SEQU_OPER_ATIV >>>
        /// <summary>
        ///Ler o xml de uma operação conforme o filtro especificado
        /// 
        /// ATENÇÃO: O Codigo da Operação é unico por Sistema
        /// Operações compromissada e redesconto repete o codigo da operação e o que desempata é o identificador do lastro
        /// </summary>
        /// <param name="codigoOperacao">O Codigo da Operação é unico por Sistema</param>
        /// <returns></returns>
        public EstruturaOperacao ObterOperacao(int numeroSequenciaOperacao)
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
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPS_TB_OPER_ATIV_02";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_OPER_ATIV(numeroSequenciaOperacao, ParameterDirection.Input));


                    _OraDA.Fill(DsMesg.TB_OPER_ATIV);

                    this.ProcessarOperacao(DsMesg.TB_OPER_ATIV.DefaultView);
                    return TB_OPER_ATIV;

                }
            }
            catch (Exception ex)
            {
                throw new Exception("ObterOperacao()" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterOperacao - OVERLOAD CO_OPER_ATIV >>>
        /// <summary>
        ///Ler o xml de uma operação conforme o filtro especificado
        /// 
        /// ATENÇÃO: O Codigo da Operação é unico por Sistema
        /// Operações compromissada e redesconto repete o codigo da operação e o que desempata é o identificador do lastro
        /// </summary>
        /// <param name="codigoOperacao">O Codigo da Operação é unico por Sistema</param>
        /// <returns></returns>
        public EstruturaOperacao ObterOperacao(string codigoOperacao, string siglaSistema)
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
                    _OracleCommand.CommandType = CommandType.StoredProcedure;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPS_TB_OPER_ATIV_06";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CO_OPER_ATIV(codigoOperacao, ParameterDirection.Input));
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.SG_SIST(siglaSistema, ParameterDirection.Input));

                    _OraDA.Fill(DsMesg.TB_OPER_ATIV);

                    this.ProcessarOperacao(DsMesg.TB_OPER_ATIV.DefaultView);
                    return TB_OPER_ATIV;

                }
            }
            catch (Exception ex)
            {
                throw new Exception("ObterOperacao() - OVERLOAD CO_OPER_ATIV" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterOperacaoXML >>>
        /// <summary>
        /// Ler o xml de uma operação conforme o filtro especificado
        /// ATENÇÃO: O Codigo da Operação é unico por Sistema
        /// Operações compromissada e redesconto repete o codigo da operação e o que desempata é o identificador do lastro
        /// </summary>
        /// <param name="codigoOperacao">O Codigo da Operação é unico por Sistema</param>
        /// <returns></returns>
        public DataTable ObterOperacaoXML(int codigoOperacao)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    DataTable dt = new DataTable();
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPS_TB_OPER_ATIV";

                    _OracleCommand.Parameters.Add(A8NETOracleParameter.CURSOR());
                    _OracleCommand.Parameters.Add(A8NETOracleParameter.NU_SEQU_OPER_ATIV(codigoOperacao, ParameterDirection.Input));

                    _OraDA.Fill(dt);

                    return dt;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ObterOperacaoXML()" + ex.ToString());
            }
        }
        #endregion

        #region <<< PersisteTB_OPER_ATIV() >>>
        /// <summary>
        /// utiliza da estrutura de datasets para inserir as informações na tabela TB_OPER_ATIV
        /// </summary>
        /// <param name="dataSetOperacao">dataSetOperacao</param>
        public void PersisteTB_OPER_ATIV(ref DsTB_OPER_ATIV dsTB_OPER_ATIV)
        {
            OracleConnection OracleConn = null;
            OracleCommand ComandoInsert = new OracleCommand();
            OracleDataAdapter AdapterORA = new OracleDataAdapter();

            try
            {
                using (OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    ComandoInsert.Connection = OracleConn;
                    ComandoInsert.CommandType = CommandType.StoredProcedure;
                    ComandoInsert.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPI_TB_OPER_ATIV";

                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_SEQU_OPER_ATIV(null, ParameterDirection.Output));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_OPER(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_LOCA_LIQU(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_LIQU_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_EMPR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_USUA_CADR_OPER(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.HO_ENVI_MESG_SPB(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_COMD_OPER(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_COMD_OPER_RETN(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DT_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DT_OPER_ATIV_RETN(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_VEIC_LEGA(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.SG_SIST(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_CNTA_CUTD_SELIC_VEIC_LEGA(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_CNPJ_CNPT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_CNTA_CUTD_SELIC_CNPT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NO_CNPT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_OPER_DEBT_CRED(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_ATIV_MERC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DE_ATIV_MERC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.PU_ATIV_MERC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.QT_ATIV_MERC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_ENTR_SAID_RECU_FINC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DT_VENC_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.VA_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.VA_OPER_ATIV_REAJ(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DT_LIQU_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_CPRO_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_CPRO_RETN_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_DISP_CONS(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_ENVI_PREV_SIST_PJ(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_ENVI_RELZ_SIST_PJ(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_ENVI_PREV_SIST_A6(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_ENVI_RELZ_SIST_A6(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_ULTI_SITU_PROC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_ACAO_OPER_ATIV_EXEC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_COMD_ACAO_EXEC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.IN_ENTR_MANU(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_PRTC_OPER_LG(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_SEQU_CNCL_OPER_ATIV_MESG(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_CTRL_MESG_SPB_ORIG(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.PE_TAXA_NEGO(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_TITL_CUTD(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_OPER_CETIP(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_ISPB_BANC_LIQU_CNPT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DH_ULTI_ATLZ(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_USUA_ULTI_ATLZ(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_IF_CRED_DEBT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_AGEN_COTR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NU_CC_COTR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.PZ_DIAS_RETN_OPER_ATIV(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.VA_OPER_ATIV_RETN(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_CNPT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_CNPT_CAMR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_IDEF_LAST(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_PARP_CAMR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_PGTO_LDL(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_GRUP_LANC_FINC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_MOED_ESTR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_CNTR_SISB(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_ISPB_IF_CNPT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_PRAC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.VA_MOED_ESTR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.DT_LIQU_OPER_ATIV_MOED_ESTR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_SISB_COTR(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_CNAL_OPER_INTE(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CO_SITU_PROC_MESG_SPB_RECB(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_CNAL_VEND(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CD_SUB_PROD(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NR_IDEF_NEGO_BMC(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_NEGO(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CD_ASSO_CAMB(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.CD_OPER_ETRT(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NR_CNPJ_CPF(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.NR_OPER_CAMB_2(null, ParameterDirection.Input));
                    ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_NEGO_INTB(null, ParameterDirection.Input));

                    // ESTÁ IMPLEMENTAÇÃO ESTÁ EM STAND-BY, AGUARDANDO PRIORIZAÇÃO PARA SER IMPLANTADA
                    //ComandoInsert.Parameters.Add(A8NETOracleParameter.NO_CLIE(null, ParameterDirection.Input));
                    //ComandoInsert.Parameters.Add(A8NETOracleParameter.CD_MOED_ISO(null, ParameterDirection.Input));
                    //ComandoInsert.Parameters.Add(A8NETOracleParameter.NR_PERC_TAXA_CAMB(null, ParameterDirection.Input));
                    //ComandoInsert.Parameters.Add(A8NETOracleParameter.TP_OPER_CAMB(null, ParameterDirection.Input));
                    //ComandoInsert.Parameters.Add(A8NETOracleParameter.NR_OPER_CAMB(null, ParameterDirection.Input));

                    AdapterORA.InsertCommand = ComandoInsert;

                    //atualiza os dados na base de dados
                    AdapterORA.Update(dsTB_OPER_ATIV.TB_OPER_ATIV);
                    dsTB_OPER_ATIV.AcceptChanges();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion

        #region <<< Inserir >>>
        public long Inserir(EstruturaOperacao parametro)
        {
            try
            {
                OracleParameter ParametroOUT = A8NETOracleParameter.NU_SEQU_OPER_ATIV(null, ParameterDirection.Output);

                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPI_TB_OPER_ATIV";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
                        ParametroOUT,
						A8NETOracleParameter.TP_OPER(parametro.TP_OPER, ParameterDirection.Input),
                        A8NETOracleParameter.CO_LOCA_LIQU(parametro.CO_LOCA_LIQU, ParameterDirection.Input),
                        A8NETOracleParameter.TP_LIQU_OPER_ATIV(parametro.TP_LIQU_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.CO_EMPR(parametro.CO_EMPR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_USUA_CADR_OPER(parametro.CO_USUA_CADR_OPER, ParameterDirection.Input),
                        A8NETOracleParameter.HO_ENVI_MESG_SPB(parametro.HO_ENVI_MESG_SPB, ParameterDirection.Input),
                        A8NETOracleParameter.CO_OPER_ATIV(parametro.CO_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.NU_COMD_OPER(parametro.NU_COMD_OPER, ParameterDirection.Input),
                        A8NETOracleParameter.NU_COMD_OPER_RETN(parametro.NU_COMD_OPER_RETN, ParameterDirection.Input),
                        A8NETOracleParameter.DT_OPER_ATIV(parametro.DT_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.DT_OPER_ATIV_RETN(parametro.DT_OPER_ATIV_RETN, ParameterDirection.Input),
                        A8NETOracleParameter.CO_VEIC_LEGA(parametro.CO_VEIC_LEGA, ParameterDirection.Input),
                        A8NETOracleParameter.SG_SIST(parametro.SG_SIST, ParameterDirection.Input),
                        A8NETOracleParameter.CO_CNTA_CUTD_SELIC_VEIC_LEGA(parametro.CO_CNTA_CUTD_SELIC_VEIC_LEGA, ParameterDirection.Input),
                        A8NETOracleParameter.CO_CNPJ_CNPT(parametro.CO_CNPJ_CNPT, ParameterDirection.Input),
                        A8NETOracleParameter.CO_CNTA_CUTD_SELIC_CNPT(parametro.CO_CNTA_CUTD_SELIC_CNPT, ParameterDirection.Input),
                        A8NETOracleParameter.NO_CNPT(parametro.NO_CNPT, ParameterDirection.Input),
                        A8NETOracleParameter.IN_OPER_DEBT_CRED(parametro.IN_OPER_DEBT_CRED, ParameterDirection.Input),
                        A8NETOracleParameter.NU_ATIV_MERC(parametro.NU_ATIV_MERC, ParameterDirection.Input),
                        A8NETOracleParameter.DE_ATIV_MERC(parametro.DE_ATIV_MERC, ParameterDirection.Input),
                        A8NETOracleParameter.PU_ATIV_MERC(parametro.PU_ATIV_MERC, ParameterDirection.Input),
                        A8NETOracleParameter.QT_ATIV_MERC(parametro.QT_ATIV_MERC, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENTR_SAID_RECU_FINC(parametro.IN_ENTR_SAID_RECU_FINC, ParameterDirection.Input),
                        A8NETOracleParameter.DT_VENC_ATIV(parametro.DT_VENC_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.VA_OPER_ATIV(parametro.VA_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.VA_OPER_ATIV_REAJ(parametro.VA_OPER_ATIV_REAJ, ParameterDirection.Input),
                        A8NETOracleParameter.DT_LIQU_OPER_ATIV(parametro.DT_LIQU_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.TP_CPRO_OPER_ATIV(parametro.TP_CPRO_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.TP_CPRO_RETN_OPER_ATIV(parametro.TP_CPRO_RETN_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.IN_DISP_CONS(parametro.IN_DISP_CONS, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENVI_PREV_SIST_PJ(parametro.IN_ENVI_PREV_SIST_PJ, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENVI_RELZ_SIST_PJ(parametro.IN_ENVI_RELZ_SIST_PJ, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENVI_PREV_SIST_A6(parametro.IN_ENVI_PREV_SIST_A6, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENVI_RELZ_SIST_A6(parametro.IN_ENVI_RELZ_SIST_A6, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ULTI_SITU_PROC(parametro.CO_ULTI_SITU_PROC, ParameterDirection.Input),
                        A8NETOracleParameter.TP_ACAO_OPER_ATIV_EXEC(parametro.TP_ACAO_OPER_ATIV_EXEC, ParameterDirection.Input),
                        A8NETOracleParameter.NU_COMD_ACAO_EXEC(parametro.NU_COMD_ACAO_EXEC, ParameterDirection.Input),
                        A8NETOracleParameter.IN_ENTR_MANU(parametro.IN_ENTR_MANU, ParameterDirection.Input),
                        A8NETOracleParameter.NU_PRTC_OPER_LG(parametro.NU_PRTC_OPER_LG, ParameterDirection.Input),
                        A8NETOracleParameter.NU_SEQU_CNCL_OPER_ATIV_MESG(parametro.NU_SEQU_CNCL_OPER_ATIV_MESG, ParameterDirection.Input),
                        A8NETOracleParameter.NU_CTRL_MESG_SPB_ORIG(parametro.NU_CTRL_MESG_SPB_ORIG, ParameterDirection.Input),
                        A8NETOracleParameter.PE_TAXA_NEGO(parametro.PE_TAXA_NEGO, ParameterDirection.Input),
                        A8NETOracleParameter.CO_TITL_CUTD(parametro.CO_TITL_CUTD, ParameterDirection.Input),
                        A8NETOracleParameter.CO_OPER_CETIP(parametro.CO_OPER_CETIP, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ISPB_BANC_LIQU_CNPT(parametro.CO_ISPB_BANC_LIQU_CNPT, ParameterDirection.Input),
                        A8NETOracleParameter.DH_ULTI_ATLZ(parametro.DH_ULTI_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(parametro.CO_ETCA_TRAB_ULTI_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.CO_USUA_ULTI_ATLZ(parametro.CO_USUA_ULTI_ATLZ, ParameterDirection.Input),
                        A8NETOracleParameter.TP_IF_CRED_DEBT(parametro.TP_IF_CRED_DEBT, ParameterDirection.Input),
                        A8NETOracleParameter.CO_AGEN_COTR(parametro.CO_AGEN_COTR, ParameterDirection.Input),
                        A8NETOracleParameter.NU_CC_COTR(parametro.NU_CC_COTR, ParameterDirection.Input),
                        A8NETOracleParameter.PZ_DIAS_RETN_OPER_ATIV(parametro.PZ_DIAS_RETN_OPER_ATIV, ParameterDirection.Input),
                        A8NETOracleParameter.VA_OPER_ATIV_RETN(parametro.VA_OPER_ATIV_RETN, ParameterDirection.Input),
                        A8NETOracleParameter.TP_CNPT(parametro.TP_CNPT, ParameterDirection.Input),
                        A8NETOracleParameter.CO_CNPT_CAMR(parametro.CO_CNPT_CAMR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_IDEF_LAST(parametro.CO_IDEF_LAST, ParameterDirection.Input),
                        A8NETOracleParameter.CO_PARP_CAMR(parametro.CO_PARP_CAMR, ParameterDirection.Input),
                        A8NETOracleParameter.TP_PGTO_LDL(parametro.TP_PGTO_LDL, ParameterDirection.Input),
                        A8NETOracleParameter.CO_GRUP_LANC_FINC(parametro.CO_GRUP_LANC_FINC, ParameterDirection.Input),
                        A8NETOracleParameter.CO_MOED_ESTR(parametro.CO_MOED_ESTR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_CNTR_SISB(parametro.CO_CNTR_SISB, ParameterDirection.Input),
                        A8NETOracleParameter.CO_ISPB_IF_CNPT(parametro.CO_ISPB_IF_CNPT, ParameterDirection.Input),
                        A8NETOracleParameter.CO_PRAC(parametro.CO_PRAC, ParameterDirection.Input),
                        A8NETOracleParameter.VA_MOED_ESTR(parametro.VA_MOED_ESTR, ParameterDirection.Input),
                        A8NETOracleParameter.DT_LIQU_OPER_ATIV_MOED_ESTR(parametro.DT_LIQU_OPER_ATIV_MOED_ESTR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SISB_COTR(parametro.CO_SISB_COTR, ParameterDirection.Input),
                        A8NETOracleParameter.CO_CNAL_OPER_INTE(parametro.CO_CNAL_OPER_INTE, ParameterDirection.Input),
                        A8NETOracleParameter.CO_SITU_PROC_MESG_SPB_RECB(parametro.CO_SITU_PROC_MESG_SPB_RECB, ParameterDirection.Input),
                        A8NETOracleParameter.TP_CNAL_VEND(parametro.TP_CNAL_VEND, ParameterDirection.Input),
                        A8NETOracleParameter.CD_SUB_PROD(parametro.CD_SUB_PROD, ParameterDirection.Input),
                        A8NETOracleParameter.NR_IDEF_NEGO_BMC(parametro.NR_IDEF_NEGO_BMC, ParameterDirection.Input),
                        A8NETOracleParameter.TP_NEGO(parametro.TP_NEGO, ParameterDirection.Input),
                        A8NETOracleParameter.CD_ASSO_CAMB(parametro.CD_ASSO_CAMB, ParameterDirection.Input),
                        A8NETOracleParameter.CD_OPER_ETRT(parametro.CD_OPER_ETRT, ParameterDirection.Input),
                        A8NETOracleParameter.NR_CNPJ_CPF(parametro.NR_CNPJ_CPF, ParameterDirection.Input),
                        A8NETOracleParameter.NR_OPER_CAMB_2(parametro.NR_OPER_CAMB_2, ParameterDirection.Input),
                        A8NETOracleParameter.TP_NEGO_INTB(parametro.TP_NEGO_INTB, ParameterDirection.Input)
                        });
                    _OracleCommand.ExecuteNonQuery();

                    return long.Parse(ParametroOUT.Value.ToString());
                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }
        #endregion

        #region <<< Atualizar >>>
        public void Atualizar(EstruturaOperacao parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPU_TB_OPER_ATIV";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
						A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.TP_OPER(parametro.TP_OPER, ParameterDirection.Input),
						A8NETOracleParameter.CO_LOCA_LIQU(parametro.CO_LOCA_LIQU, ParameterDirection.Input),
						A8NETOracleParameter.TP_LIQU_OPER_ATIV(parametro.TP_LIQU_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.CO_EMPR(parametro.CO_EMPR, ParameterDirection.Input),
						A8NETOracleParameter.CO_USUA_CADR_OPER(parametro.CO_USUA_CADR_OPER, ParameterDirection.Input),
						A8NETOracleParameter.HO_ENVI_MESG_SPB(parametro.HO_ENVI_MESG_SPB, ParameterDirection.Input),
						A8NETOracleParameter.CO_OPER_ATIV(parametro.CO_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.NU_COMD_OPER(parametro.NU_COMD_OPER, ParameterDirection.Input),
						A8NETOracleParameter.NU_COMD_OPER_RETN(parametro.NU_COMD_OPER_RETN, ParameterDirection.Input),
						A8NETOracleParameter.DT_OPER_ATIV(parametro.DT_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.DT_OPER_ATIV_RETN(parametro.DT_OPER_ATIV_RETN, ParameterDirection.Input),
						A8NETOracleParameter.CO_VEIC_LEGA(parametro.CO_VEIC_LEGA, ParameterDirection.Input),
						A8NETOracleParameter.SG_SIST(parametro.SG_SIST, ParameterDirection.Input),
						A8NETOracleParameter.CO_CNTA_CUTD_SELIC_VEIC_LEGA(parametro.CO_CNTA_CUTD_SELIC_VEIC_LEGA, ParameterDirection.Input),
						A8NETOracleParameter.CO_CNPJ_CNPT(parametro.CO_CNPJ_CNPT, ParameterDirection.Input),
						A8NETOracleParameter.CO_CNTA_CUTD_SELIC_CNPT(parametro.CO_CNTA_CUTD_SELIC_CNPT, ParameterDirection.Input),
						A8NETOracleParameter.NO_CNPT(parametro.NO_CNPT, ParameterDirection.Input),
						A8NETOracleParameter.IN_OPER_DEBT_CRED(parametro.IN_OPER_DEBT_CRED, ParameterDirection.Input),
						A8NETOracleParameter.NU_ATIV_MERC(parametro.NU_ATIV_MERC, ParameterDirection.Input),
						A8NETOracleParameter.DE_ATIV_MERC(parametro.DE_ATIV_MERC, ParameterDirection.Input),
						A8NETOracleParameter.PU_ATIV_MERC(parametro.PU_ATIV_MERC, ParameterDirection.Input),
						A8NETOracleParameter.QT_ATIV_MERC(parametro.QT_ATIV_MERC, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENTR_SAID_RECU_FINC(parametro.IN_ENTR_SAID_RECU_FINC, ParameterDirection.Input),
						A8NETOracleParameter.DT_VENC_ATIV(parametro.DT_VENC_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.VA_OPER_ATIV(parametro.VA_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.VA_OPER_ATIV_REAJ(parametro.VA_OPER_ATIV_REAJ, ParameterDirection.Input),
						A8NETOracleParameter.DT_LIQU_OPER_ATIV(parametro.DT_LIQU_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.TP_CPRO_OPER_ATIV(parametro.TP_CPRO_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.TP_CPRO_RETN_OPER_ATIV(parametro.TP_CPRO_RETN_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.IN_DISP_CONS(parametro.IN_DISP_CONS, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENVI_PREV_SIST_PJ(parametro.IN_ENVI_PREV_SIST_PJ, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENVI_RELZ_SIST_PJ(parametro.IN_ENVI_RELZ_SIST_PJ, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENVI_PREV_SIST_A6(parametro.IN_ENVI_PREV_SIST_A6, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENVI_RELZ_SIST_A6(parametro.IN_ENVI_RELZ_SIST_A6, ParameterDirection.Input),
						A8NETOracleParameter.CO_ULTI_SITU_PROC(parametro.CO_ULTI_SITU_PROC, ParameterDirection.Input),
						A8NETOracleParameter.TP_ACAO_OPER_ATIV_EXEC(parametro.TP_ACAO_OPER_ATIV_EXEC, ParameterDirection.Input),
						A8NETOracleParameter.NU_COMD_ACAO_EXEC(parametro.NU_COMD_ACAO_EXEC, ParameterDirection.Input),
						A8NETOracleParameter.IN_ENTR_MANU(parametro.IN_ENTR_MANU, ParameterDirection.Input),
						A8NETOracleParameter.NU_PRTC_OPER_LG(parametro.NU_PRTC_OPER_LG, ParameterDirection.Input),
						A8NETOracleParameter.NU_SEQU_CNCL_OPER_ATIV_MESG(parametro.NU_SEQU_CNCL_OPER_ATIV_MESG, ParameterDirection.Input),
						A8NETOracleParameter.NU_CTRL_MESG_SPB_ORIG(parametro.NU_CTRL_MESG_SPB_ORIG, ParameterDirection.Input),
						A8NETOracleParameter.PE_TAXA_NEGO(parametro.PE_TAXA_NEGO, ParameterDirection.Input),
						A8NETOracleParameter.CO_TITL_CUTD(parametro.CO_TITL_CUTD, ParameterDirection.Input),
						A8NETOracleParameter.CO_OPER_CETIP(parametro.CO_OPER_CETIP, ParameterDirection.Input),
						A8NETOracleParameter.CO_ISPB_BANC_LIQU_CNPT(parametro.CO_ISPB_BANC_LIQU_CNPT, ParameterDirection.Input),
						A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(parametro.CO_ETCA_TRAB_ULTI_ATLZ, ParameterDirection.Input),
						A8NETOracleParameter.CO_USUA_ULTI_ATLZ(parametro.CO_USUA_ULTI_ATLZ, ParameterDirection.Input),
						A8NETOracleParameter.TP_IF_CRED_DEBT(parametro.TP_IF_CRED_DEBT, ParameterDirection.Input),
						A8NETOracleParameter.CO_AGEN_COTR(parametro.CO_AGEN_COTR, ParameterDirection.Input),
						A8NETOracleParameter.NU_CC_COTR(parametro.NU_CC_COTR, ParameterDirection.Input),
						A8NETOracleParameter.PZ_DIAS_RETN_OPER_ATIV(parametro.PZ_DIAS_RETN_OPER_ATIV, ParameterDirection.Input),
						A8NETOracleParameter.VA_OPER_ATIV_RETN(parametro.VA_OPER_ATIV_RETN, ParameterDirection.Input),
						A8NETOracleParameter.TP_CNPT(parametro.TP_CNPT, ParameterDirection.Input),
						A8NETOracleParameter.CO_CNPT_CAMR(parametro.CO_CNPT_CAMR, ParameterDirection.Input),
						A8NETOracleParameter.CO_IDEF_LAST(parametro.CO_IDEF_LAST, ParameterDirection.Input),
						A8NETOracleParameter.CO_PARP_CAMR(parametro.CO_PARP_CAMR, ParameterDirection.Input),
						A8NETOracleParameter.TP_PGTO_LDL(parametro.TP_PGTO_LDL, ParameterDirection.Input),
						A8NETOracleParameter.CO_GRUP_LANC_FINC(parametro.CO_GRUP_LANC_FINC, ParameterDirection.Input),
						A8NETOracleParameter.CO_MOED_ESTR(parametro.CO_MOED_ESTR, ParameterDirection.Input),
						A8NETOracleParameter.CO_CNTR_SISB(parametro.CO_CNTR_SISB, ParameterDirection.Input),
						A8NETOracleParameter.CO_ISPB_IF_CNPT(parametro.CO_ISPB_IF_CNPT, ParameterDirection.Input),
						A8NETOracleParameter.CO_PRAC(parametro.CO_PRAC, ParameterDirection.Input),
						A8NETOracleParameter.VA_MOED_ESTR(parametro.VA_MOED_ESTR, ParameterDirection.Input),
						A8NETOracleParameter.DT_LIQU_OPER_ATIV_MOED_ESTR(parametro.DT_LIQU_OPER_ATIV_MOED_ESTR, ParameterDirection.Input),
						A8NETOracleParameter.CO_SISB_COTR(parametro.CO_SISB_COTR, ParameterDirection.Input),
						A8NETOracleParameter.CO_CNAL_OPER_INTE(parametro.CO_CNAL_OPER_INTE, ParameterDirection.Input),
						A8NETOracleParameter.CO_SITU_PROC_MESG_SPB_RECB(parametro.CO_SITU_PROC_MESG_SPB_RECB, ParameterDirection.Input),
						A8NETOracleParameter.TP_CNAL_VEND(parametro.TP_CNAL_VEND, ParameterDirection.Input),
						A8NETOracleParameter.CD_SUB_PROD(parametro.CD_SUB_PROD, ParameterDirection.Input),
						A8NETOracleParameter.NR_IDEF_NEGO_BMC(parametro.NR_IDEF_NEGO_BMC, ParameterDirection.Input),
						A8NETOracleParameter.TP_NEGO(parametro.TP_NEGO, ParameterDirection.Input),
						A8NETOracleParameter.CD_ASSO_CAMB(parametro.CD_ASSO_CAMB, ParameterDirection.Input),
						A8NETOracleParameter.CD_OPER_ETRT(parametro.CD_OPER_ETRT, ParameterDirection.Input),
						A8NETOracleParameter.NR_CNPJ_CPF(parametro.NR_CNPJ_CPF, ParameterDirection.Input),
                        A8NETOracleParameter.NR_OPER_CAMB_2(parametro.NR_OPER_CAMB_2, ParameterDirection.Input),
                        A8NETOracleParameter.TP_NEGO_INTB(parametro.TP_NEGO_INTB, ParameterDirection.Input)
                        });
                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("OperacaoDAO.Atualizar() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< AtualizarStatus >>>
        public void AtualizarStatus(EstruturaOperacao parametro)
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    //Abrir conexao banco de dados
                    OracleConn.Open();

                    _OracleCommand.Parameters.Clear();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandText = "A8PROC.PKG_A8_TB_OPER_ATIV.SPU_STATUS";

                    _OracleCommand.Parameters.AddRange(new OracleParameter[]{
						A8NETOracleParameter.NU_SEQU_OPER_ATIV(parametro.NU_SEQU_OPER_ATIV, ParameterDirection.Input),
						
						A8NETOracleParameter.CO_ULTI_SITU_PROC(parametro.CO_ULTI_SITU_PROC, ParameterDirection.Input),
						A8NETOracleParameter.CO_ETCA_TRAB_ULTI_ATLZ(parametro.CO_ETCA_TRAB_ULTI_ATLZ, ParameterDirection.Input),
						A8NETOracleParameter.CO_USUA_ULTI_ATLZ(parametro.CO_USUA_ULTI_ATLZ, ParameterDirection.Input)}
                        );
                    _OracleCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("OperacaoDAO.AtualizarStatus() - " + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterTipoBackoffice >>>
        /// <summary>
        /// Obter o TipoBackiffice de uma determinada Operacao
        /// </summary>
        /// <param name="numeroSequenciaOperacao">Identificador único da Operação</param>
        /// <returns></returns>
        public int ObterTipoBackoffice(OperacaoDAO operacaoDATA, DataTable dtVeiculoLegal)
        {
            try
            {
               return int.Parse(dtVeiculoLegal.Select(string.Format(@"CO_VEIC_LEGA = '{0}' 
                                                                  AND SG_SIST      = '{1}'", 
                                                                  operacaoDATA.TB_OPER_ATIV.CO_VEIC_LEGA,
                                                                  operacaoDATA.TB_OPER_ATIV.SG_SIST))[0]["TP_BKOF"].ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("ObterTipoBackoffice()" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterTipoMensagemRetorno >>>
        /// <summary>
        /// Obter o TipoMensagemRetorno de uma determinada Operacao
        /// </summary>
        /// <param name="numeroSequenciaOperacao">Identificador único da Operação</param>
        /// <returns></returns>
        public string ObterTipoMensagemRetorno(int tipoOperacao, DataTable dtTipoOper)
        {
            try
            {
                return dtTipoOper.Select("TP_OPER = " + tipoOperacao)[0]["TP_MESG_RETN_INTE"].ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("ObterTipoMensagemRetorno()" + ex.ToString());
            }
        }
        #endregion

        #region <<< ObterProximoNumeroSequenciaOperacao >>>
        public long ObterProximoNumeroSequenciaOperacao()
        {
            try
            {
                using (OracleConnection OracleConn = new OracleConnection(base.GetStringConnection()))
                {
                    OracleConn.Open();
                    _OracleCommand.Connection = OracleConn;
                    _OracleCommand.CommandType = CommandType.Text;
                    _OracleCommand.CommandText = "SELECT A8.SQ_A8_NU_SEQU_OPER_ATIV.NEXTVAL FROM DUAL";
                    return (long)_OracleCommand.ExecuteOracleScalar();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("OperacaoDAO.ObterProximoNumeroSequenciaOperacao() - " + ex.ToString());
            }
        }
        #endregion

    }
}

