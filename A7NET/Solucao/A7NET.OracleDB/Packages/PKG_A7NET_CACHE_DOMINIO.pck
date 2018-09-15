CREATE OR REPLACE PACKAGE PKG_A7NET_CACHE_DOMINIO IS

TYPE tp_cursor IS REF CURSOR;


PROCEDURE SPS_TB_MESG(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_TIPO_MESG(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_EMPRESA_HO(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_SIST(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_REGR_TRAP_MESG(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_MENSAGEM_SPB(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_ENDE_FILA_MQSE(
   p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPS_TB_TIPO_OPER(
   p_cur_out_CURSOR OUT tp_cursor
);

END PKG_A7NET_CACHE_DOMINIO;
/
CREATE OR REPLACE PACKAGE BODY PKG_A7NET_CACHE_DOMINIO IS

/******************************************************************************************
 Objetos criados: SPS_TB_MESG
 Descricao:       Busca estrutura da tabela TB_MESG
 Autor:           Ivan Tabarino
 Data:            28/03/2011
******************************************************************************************/
PROCEDURE SPS_TB_MESG(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT CO_MESG,
          TP_FORM_MESG_SAID,
          DH_MESG,
          TP_MESG,
          CO_EMPR_ORIG,
          SG_SIST_ORIG,
          '' AS SG_SIST_DEST,
          DH_INIC_VIGE_REGR_TRAP,
          CO_CMPO_ATRB_IDEF_MESG,
          CO_MESG_MQSE,
          CO_TEXT_XML_ENTR,
          CO_TEXT_XML_SAID,
          '' AS TX_DTLH_OCOR_ERRO,
          '' AS TX_CNTD_ENTR,
          '' AS TX_CNTD_SAID
   FROM   A7.TB_MESG
   WHERE  Rownum = 1;

   
END SPS_TB_MESG;

/******************************************************************************************
 Objetos criados: SPS_TB_TIPO_MESG
 Descricao:       Busca os dados da tabela TB_TIPO_MESG
 Autor:           Ivan Tabarino
 Data:            30/03/2011
******************************************************************************************/
PROCEDURE SPS_TB_TIPO_MESG(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT TP_MESG,
          TP_FORM_MESG_SAID,
          DT_INIC_VIGE_MESG,
          NO_TIPO_MESG,
          TP_NATZ_MESG,
          TP_CTER_DELI,
          CO_PRIO_FILA_SAID_MESG,
          DT_FIM_VIGE_MESG,
          CO_TEXT_XML,
          NO_TITU_MESG,
          CO_USUA_ULTI_ATLZ,
          CO_ETCA_TRAB_ULTI_ATLZ,
          DH_ULTI_ATLZ
   FROM   A7.TB_TIPO_MESG
   WHERE (DT_FIM_VIGE_MESG IS NULL
   OR     DT_FIM_VIGE_MESG >= SYSDATE);


END SPS_TB_TIPO_MESG;

/******************************************************************************************
 Objetos criados: SPS_TB_EMPRESA_HO
 Descricao:       Busca os dados da tabela TB_EMPRESA_HO
 Autor:           Ivan Tabarino
 Data:            30/03/2011
******************************************************************************************/
PROCEDURE SPS_TB_EMPRESA_HO(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT CO_EMPR,
          CO_EMPR_FUSI,
          CO_BANC,
          NU_CNPJ,
          NO_EMPR,
          NO_REDU_EMPR,
          SQ_ISPB,
          CO_GRUP_CPUL,
          DT_INIC_VIGE,
          DT_FIM_VIGE,
          ID_USUA_ULTI_ATLZ,
          DH_ULTI_ATLZ,
          ID_PART_CAMR_CETIP
   FROM   A8.TB_EMPRESA_HO
   WHERE  DT_INIC_VIGE <= SYSDATE
   AND   (DT_FIM_VIGE  IS NULL
   OR     DT_FIM_VIGE  >= SYSDATE);


END SPS_TB_EMPRESA_HO;

/******************************************************************************************
 Objetos criados: SPS_TB_SIST
 Descricao:       Busca os dados da tabela TB_SIST
 Autor:           Ivan Tabarino
 Data:            30/03/2011
******************************************************************************************/
PROCEDURE SPS_TB_SIST(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT SG_SIST,
          CO_EMPR,
          NO_SIST,
          DT_INIC_VIGE_SIST,
          DT_FIM_VIGE_SIST,
          CO_USUA_ULTI_ATLZ,
          CO_ETCA_TRAB_ULTI_ATLZ,
          DH_ULTI_ATLZ
   FROM   A7.TB_SIST
   WHERE  DT_INIC_VIGE_SIST <= SYSDATE
   AND   (DT_FIM_VIGE_SIST  IS NULL
   OR     DT_FIM_VIGE_SIST  >= SYSDATE);


END SPS_TB_SIST;

/******************************************************************************************
 Objetos criados: SPS_TB_REGR_TRAP_MESG
 Descricao:       Busca os dados da tabela TB_REGR_TRAP_MESG
 Autor:           Ivan Tabarino
 Data:            30/03/2011
******************************************************************************************/
PROCEDURE SPS_TB_REGR_TRAP_MESG(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT REGR_TRAP.TP_MESG,
          REGR_TRAP.SG_SIST_ORIG,
          REGR_TRAP.CO_EMPR_ORIG,
          REGR_TRAP.DH_INIC_VIGE_REGR_TRAP,
          REGR_TRAP.TP_FORM_MESG_SAID,
          REGR_TRAP.TP_FORM_MESG_ENTR,
          REGR_TRAP.TP_CTER_DELI AS TP_CTER_DELI_ENTR,
          TIPO_MESG.TP_CTER_DELI AS TP_CTER_DELI_SAID,
          REGR_TRAP.DT_FIM_VIGE_REGR_TRAP,
          REGR_TRAP.IN_EXIS_REGR_TRNF,
          REGR_TRAP.CO_TEXT_XML AS CO_TEXT_XML_REGR,
          TIPO_MESG.CO_TEXT_XML AS CO_TEXT_XML_MESG,
          REGR_TRAP.CO_USUA_ULTI_ATLZ,
          REGR_TRAP.CO_ETCA_TRAB_ULTI_ATLZ,
          REGR_TRAP.DH_ULTI_ATLZ,
          REGR_SIST.SG_SIST_DEST,
          REGR_SIST.CO_EMPR_DEST,
          TIPO_MESG.TP_NATZ_MESG,
          TIPO_MESG.NO_TITU_MESG
   FROM   A7.TB_REGR_TRAP_MESG REGR_TRAP,
          A7.TB_REGR_SIST_DEST REGR_SIST,
          A7.TB_TIPO_MESG      TIPO_MESG
   WHERE  REGR_TRAP.TP_MESG                 = REGR_SIST.TP_MESG
   AND    REGR_TRAP.TP_MESG                 = TIPO_MESG.TP_MESG
   AND    REGR_TRAP.SG_SIST_ORIG            = REGR_SIST.SG_SIST_ORIG
   AND    REGR_TRAP.CO_EMPR_ORIG            = REGR_SIST.CO_EMPR_ORIG
   AND    REGR_TRAP.TP_FORM_MESG_SAID       = REGR_SIST.TP_FORM_MESG_SAID
   AND    REGR_TRAP.TP_FORM_MESG_SAID       = TIPO_MESG.TP_FORM_MESG_SAID
   AND    REGR_TRAP.DH_INIC_VIGE_REGR_TRAP  = REGR_SIST.DH_INIC_VIGE_REGR_TRAP
   AND    REGR_TRAP.DH_INIC_VIGE_REGR_TRAP <= SYSDATE
   AND   (REGR_TRAP.DT_FIM_VIGE_REGR_TRAP  IS NULL
   OR     REGR_TRAP.DT_FIM_VIGE_REGR_TRAP  >= SYSDATE);


END SPS_TB_REGR_TRAP_MESG;

/******************************************************************************************
 Objetos criados: SPS_TB_MENSAGEM_SPB
 Descricao:       Busca os dados das Mensagens SPB para montagem da mensagem de saida
 Autor:           Ivan Tabarino
 Data:            14/04/2011
******************************************************************************************/
PROCEDURE SPS_TB_MENSAGEM_SPB(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT RTRIM(MESG.CO_MESG)           AS CO_MESG,
          RTRIM(TAG.NO_TAG)             AS NO_TAG,
          RTRIM(TIPO_TAG.NO_TIPO_TAG)   AS NO_TIPO_TAG,
          NVL(TIPO_TAG.QT_CASA_DECI, 0) AS QT_CASA_DECI
   FROM   A8.TB_MENSAGEM     MESG,
          A8.TB_TAG          TAG,
          A8.TB_TAG_MENSAGEM TAG_MESG,
          A8.TB_TIPO_TAG     TIPO_TAG
   WHERE  MESG.SQ_MESG          =    TAG_MESG.SQ_MESG
   AND    TAG_MESG.SQ_TAG       =    TAG.SQ_TAG
   AND    TAG.SQ_TIPO_TAG       =    TIPO_TAG.SQ_TIPO_TAG
   AND  ((TIPO_TAG.NO_TIPO_TAG  LIKE '%Data%'
   OR     TIPO_TAG.NO_TIPO_TAG  LIKE '%Hora%'
   OR     TIPO_TAG.NO_TIPO_TAG  LIKE '%Ano%'
   OR     TIPO_TAG.NO_TIPO_TAG  LIKE '%Mes%')
   OR     TIPO_TAG.QT_CASA_DECI >    0)
   ORDER  BY 1, 2, 3;
   

END SPS_TB_MENSAGEM_SPB;

/******************************************************************************************
 Objetos criados: SPS_TB_ENDE_FILA_MQSE
 Descricao:       Busca os dados das Filas dos Sistemas Cadastrados
 Autor:           Ivan Tabarino
 Data:            18/04/2011
******************************************************************************************/
PROCEDURE SPS_TB_ENDE_FILA_MQSE(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT SG_SIST_DEST,           
          CO_EMPR_DEST,
          NO_FILA_MQSE,           
          CO_USUA_ULTI_ATLZ,           
          CO_ETCA_TRAB_ULTI_ATLZ,           
          DH_ULTI_ATLZ  
   FROM   A7.TB_ENDE_FILA_MQSE   
   WHERE  CO_EMPR_DEST = 558;
   

END SPS_TB_ENDE_FILA_MQSE;

/******************************************************************************************
 Objetos criados: SPS_TB_TIPO_OPER
 Descricao:       Busca os dados das Filas dos Sistemas Cadastrados
 Autor:           Ivan Tabarino
 Data:            19/04/2011
******************************************************************************************/
PROCEDURE SPS_TB_TIPO_OPER(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT TP_OPER,
          TP_MESG_RECB_INTE,
          TP_MESG_RETN_INTE
   FROM  A8.TB_TIPO_OPER
   WHERE DT_INIC_VIGE <= SYSDATE
   AND  (DT_FIM_VIGE  IS NULL
   OR    DT_FIM_VIGE  >= SYSDATE)
   ORDER BY TP_OPER;


END SPS_TB_TIPO_OPER;
  
END PKG_A7NET_CACHE_DOMINIO;
/
