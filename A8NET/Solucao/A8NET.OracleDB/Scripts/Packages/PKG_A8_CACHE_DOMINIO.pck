CREATE OR REPLACE PACKAGE A8PROC.PKG_A8_CACHE_DOMINIO IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE SPS_TB_CTRL_PROC_OPER_ATIV(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_TIPO_OPER(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_FCAO_SIST_TIPO_OPER(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_PARM_FCAO_SIST(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_SITU_SPB_SITU_PROC(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_MENSAGEM(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_MBS_GRUPO(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_PARM_FCAO_SIST_EXCE(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_VEIC_LEGA(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_CTRL_DOMI(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_PRODUTO(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_TIPO_OPER_CNTD_ATRB(
   p_cur_out_CURSOR OUT tp_cursor
);

end PKG_A8_CACHE_DOMINIO;
/
CREATE OR REPLACE PACKAGE BODY A8PROC.PKG_A8_CACHE_DOMINIO IS

/******************************************************************************************
 Objetos criados: SPS_TB_CTRL_PROC_OPER_ATIV
 Descricao:       Consulta dados na tabela de Controle de Processamento de Operações
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_CTRL_PROC_OPER_ATIV(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT tp_oper, 
          co_situ_proc, 
          no_proc_oper_ativ, 
          in_envi_prev_pj, 
          tp_liqu_oper_ativ, 
          in_envi_prev_a6, 
          in_envi_relz_pj, 
          in_envi_relz_soli_a6, 
          in_envi_relz_conf_a6, 
          in_veri_regr_conf, 
          in_veri_regr_cncl, 
          in_veri_regr_libe, 
          in_envi_mesg_retn, 
          in_envi_mesg_spb, 
          in_disp_lanc_cnta_crrt, 
          in_esto_pj_a6, 
          in_envi_aler,
          in_envi_prev_pj_me,
          in_envi_relz_pj_me
   FROM   A8.TB_CTRL_PROC_OPER_ATIV;

END SPS_TB_CTRL_PROC_OPER_ATIV;  


/******************************************************************************************
 Objetos criados: SPS_TB_TIPO_OPER
 Descricao:       Consulta dados na tabela de Tipo de Operações
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_TIPO_OPER(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT tp_oper, 
          tp_bkof, 
          co_loca_liqu, 
          no_tipo_oper, 
          tp_mesg_recb_inte, 
          tp_mesg_retn_inte, 
          co_mesg_spb_regt_oper, 
          dt_inic_vige, 
          dt_fim_vige, 
          co_usua_ulti_atlz, 
          co_etca_trab_ulti_atlz, 
          dh_ulti_atlz, 
          co_oper_selic
   FROM   A8.TB_TIPO_OPER;
   
END SPS_TB_TIPO_OPER;


/******************************************************************************************
 Objetos criados: SPS_TB_FCAO_SIST_TIPO_OPER
 Descricao:       Consulta dados na tabela de Função Sistema Tipo Operação
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_FCAO_SIST_TIPO_OPER(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT tp_oper, 
          co_fcao_sist, 
          nu_orde_exec_fcao_sist
   FROM A8.TB_FCAO_SIST_TIPO_OPER;
   
END SPS_TB_FCAO_SIST_TIPO_OPER;


/******************************************************************************************
 Objetos criados: SPS_TB_PARM_FCAO_SIST
 Descricao:       Consulta dados na tabela de Parâmetro Função Sistema
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_PARM_FCAO_SIST(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT tp_oper, 
          co_fcao_sist, 
          tp_bkof, 
          co_empr, 
          in_fcao_sist_autm, 
          co_usua_ulti_atlz, 
          co_etca_trab_ulti_atlz, 
          dh_ulti_atlz, 
          tp_cond_sald 
   FROM A8.TB_PARM_FCAO_SIST;
   
END SPS_TB_PARM_FCAO_SIST;


/******************************************************************************************
 Objetos criados: SPS_TB_SITU_SPB_SITU_PROC
 Descricao:       Consulta dados na tabela de Situacao SPB
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_SITU_SPB_SITU_PROC(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR
  
   SELECT 
      A.sg_grup_mesg_spb, 
      A.no_tag, 
      A.de_domi, 
      A.co_situ_proc_oper_ativ, 
      A.co_situ_proc_mesg_spb, 
      A.co_usua_ulti_atlz, 
      A.co_etca_trab_ulti_atlz, 
      A.dh_ulti_atlz,
      B.SQ_TIPO_TAG
   FROM 
      A8.TB_SITU_SPB_SITU_PROC A,
      A8.TB_TAG                B  
   WHERE 
      A.NO_TAG  = B.NO_TAG;

   
END SPS_TB_SITU_SPB_SITU_PROC;

/******************************************************************************************
 Objetos criados: SPS_TB_MENSAGEM
 Descricao:       Consulta dados na tabela de Mensagem
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_MENSAGEM(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR
  
   SELECT 
    A.sq_mesg, 
    A.sq_even, 
    A.co_mesg, 
    A.no_mesg, 
    A.no_tag_prin_mesg, 
    A.id_usua_ulti_atlz, 
    A.dh_ulti_atlz,
    B.SQ_TIPO_FLUX
   FROM 
      A8.TB_MENSAGEM A,
      A8.TB_EVENTO   B  
   WHERE 
      A.SQ_EVEN  = B.SQ_EVEN;
END SPS_TB_MENSAGEM;
   
/******************************************************************************************
 Objetos criados: SPS_MBS_GRUPO
 Descricao:       Consulta dados na tabela de Grupo MBS
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_MBS_GRUPO(
   p_cur_out_CURSOR OUT tp_cursor
)
IS

vSELECT           VARCHAR2(5000);

BEGIN

     vSELECT := '';
     vSELECT := vSELECT || 'SELECT  M.cd_usr '
                        || '       ,M.cd_gr_usr '
                        || '       ,M.nm_gr_usr ';
     vSELECT := vSELECT || 'FROM   A8.MBS_GRUPO M ';

     OPEN p_cur_out_CURSOR FOR vSELECT;

END SPS_MBS_GRUPO;   
   
/******************************************************************************************
 Objetos criados: SPS_TB_PARM_FCAO_SIST_EXCE
 Descricao:       Consulta dados na tabela de Parametro de Funcao de Sistema
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_PARM_FCAO_SIST_EXCE(
   p_cur_out_CURSOR OUT tp_cursor
)
IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT tp_oper, 
          co_fcao_sist, 
          tp_bkof, 
          co_empr, 
          sg_sist, 
          co_loca_liqu, 
          co_grup_usua, 
          co_usua_ulti_atlz, 
          co_etca_trab_ulti_atlz, 
          dh_ulti_atlz, 
          tp_cond_sald
   FROM A8.TB_PARM_FCAO_SIST_EXCE;

END SPS_TB_PARM_FCAO_SIST_EXCE;

/******************************************************************************************
 Objetos criados: SPS_TB_VEIC_LEGA
 Descricao:       Consulta dados na tabela de Veiculo Legal
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_VEIC_LEGA(
   p_cur_out_CURSOR OUT tp_cursor
)
IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT co_veic_lega, 
          sg_sist, 
          co_empr, 
          co_grup_veic_lega, 
          tp_bkof, 
          no_veic_lega, 
          no_redu_veic_lega, 
          co_cnpj_veic_lega, 
          dt_inic_vige, 
          dt_fim_vige, 
          co_usua_ulti_atlz, 
          co_etca_trab_ulti_atlz, 
          dh_ulti_atlz, 
          co_cnta_cutd_padr_selic, 
          tp_titl_bma, 
          id_part_camr_cetip, 
          co_titl_bma
   FROM A8.TB_VEIC_LEGA;


END SPS_TB_VEIC_LEGA;

/******************************************************************************************
 Objetos criados: SPS_TB_CTRL_DOMI
 Descricao:       Consulta dados na tabela de Controle de Dominio
 Autor:           Ivan Tabarino
 Data:            06/mar/2011
******************************************************************************************/
PROCEDURE SPS_TB_CTRL_DOMI(
   p_cur_out_CURSOR OUT tp_cursor
)
IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT CTDO.NO_ATRB,
          TTAG.NO_TIPO_TAG,
          CTDO.TP_CNTR_DOMI,
          CTDO.DE_DOMI,
          DECODE(DOMI.CO_DOMI, NULL, CTDO.CO_DOMI, DOMI.CO_DOMI) AS CO_DOMI
   FROM   A8.TB_CTRL_DOMI CTDO,
          A8.TB_DOMINIO   DOMI,
          A8.TB_TIPO_TAG  TTAG
   WHERE  CTDO.CO_DOMI       =  TTAG.NO_TIPO_TAG (+)
   AND    TTAG.SQ_TIPO_TAG   =  DOMI.SQ_TIPO_TAG (+)
   AND    CTDO.TP_CNTR_DOMI IN (1, 2)
   ORDER  BY CTDO.NO_ATRB, DOMI.CO_DOMI;


END SPS_TB_CTRL_DOMI;

/******************************************************************************************
 Objetos criados: SPS_TB_PRODUTO
 Descricao:       Consulta dados na tabela de Produto
 Autor:           Ivan Tabarino
 Data:            08/mar/2011
******************************************************************************************/
PROCEDURE SPS_TB_PRODUTO(
   p_cur_out_CURSOR OUT tp_cursor
)
IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT CO_PROD,
          CO_EMPR_FUSI,
          SQ_ITEM_CAIX,
          DE_PROD,
          QT_DIAS_MAIR_VALO,
          VA_MINI_MAIR_VALO,
          QT_REGT_MAIR_VALO,
          DT_INIC_VIGE,
          DT_FIM_VIGE,
          ID_USUA_ULTI_ATLZ,
          DH_ULTI_ATLZ
   FROM   A8.TB_PRODUTO;


END SPS_TB_PRODUTO;



/******************************************************************************************
 Objetos criados: SPS_TB_TIPO_OPER_CNTD_ATRB
 Descricao:       Consulta dados na tabela de Tipo Operacao Conteudo Atributo
 Autor:           Bruno Oliveira
 Data:            15/mar/2012
******************************************************************************************/
PROCEDURE SPS_TB_TIPO_OPER_CNTD_ATRB(
   p_cur_out_CURSOR OUT tp_cursor
)
IS

BEGIN

OPEN p_cur_out_CURSOR FOR
   SELECT TP_OPER, 
          DE_CNTD_ATRB, 
          NO_ATRB_MESG, 
          CO_USUA_ULTI_ATLZ, 
          CO_ETCA_TRAB_ULTI_ATLZ, 
          DH_ULTI_ATLZ
   FROM   A8.TB_TIPO_OPER_CNTD_ATRB;

END SPS_TB_TIPO_OPER_CNTD_ATRB;


END PKG_A8_CACHE_DOMINIO;
/
