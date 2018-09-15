CREATE OR REPLACE PACKAGE A7PROC.PKG_A7_CACHE_DOMINIO IS

TYPE tp_cursor IS REF CURSOR;


PROCEDURE SPS_TB_REGR_SIST_DEST(
   p_cur_out_CURSOR OUT tp_cursor
);

end PKG_A7_CACHE_DOMINIO;
/
CREATE OR REPLACE PACKAGE BODY A7PROC.PKG_A7_CACHE_DOMINIO IS


/******************************************************************************************
 Objetos criados: SPS_TB_REGR_SIST_DEST
 Descricao:       Consulta dados na tabela de Regra Sistema Destino
 Autor:           Bruno Oliveira
 Data:            28/fev/2011
******************************************************************************************/
PROCEDURE SPS_TB_REGR_SIST_DEST(
   p_cur_out_CURSOR OUT tp_cursor
)

IS

BEGIN

OPEN p_cur_out_CURSOR FOR

   SELECT DEST.TP_FORM_MESG_SAID, 
          DEST.TP_MESG, 
          DEST.SG_SIST_ORIG, 
          DEST.CO_EMPR_ORIG, 
          DEST.DH_INIC_VIGE_REGR_TRAP, 
          DEST.SG_SIST_DEST, 
          DEST.CO_EMPR_DEST
   FROM   A7.TB_REGR_SIST_DEST DEST,
          A7.TB_REGR_TRAP_MESG REGR
   WHERE  DEST.DH_INIC_VIGE_REGR_TRAP  = REGR.DH_INIC_VIGE_REGR_TRAP
   AND    DEST.CO_EMPR_ORIG            = REGR.CO_EMPR_ORIG
   AND    DEST.SG_SIST_ORIG            = REGR.SG_SIST_ORIG
   AND    DEST.TP_MESG                 = REGR.TP_MESG
   AND   (REGR.DT_FIM_VIGE_REGR_TRAP  IS NULL
   OR     REGR.DT_FIM_VIGE_REGR_TRAP  >= SYSDATE);

   
END SPS_TB_REGR_SIST_DEST;
  
END PKG_A7_CACHE_DOMINIO;
/
