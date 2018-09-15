CREATE OR REPLACE PACKAGE PKG_A8_INTEGRACAO_PJ IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE SPS_SQ_A8_NU_SEQU_REME_PJ(
   p_num_out_RETORNO  OUT  NUMBER
);

PROCEDURE SPS_TB_PRODUTO(
   P_CO_EMPR            IN A8.TB_EMPRESA_HO.CO_EMPR%TYPE,
   P_CO_PROD            IN A8.TB_PRODUTO.CO_PROD%TYPE,
   P_VA_MINI_MAIR_VALO OUT A8.TB_PRODUTO.VA_MINI_MAIR_VALO%TYPE
);

PROCEDURE SPI_TB_HIST_ENVI_INFO_GEST(
   P_NU_SEQU_OPER_ATIV       IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.NU_SEQU_OPER_ATIV%TYPE,
   P_CO_SITU_MOVI_GEST_CAIX  IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_SITU_MOVI_GEST_CAIX%TYPE,
   P_CO_TEXT_XML             IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_TEXT_XML%TYPE
);

END PKG_A8_INTEGRACAO_PJ;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_INTEGRACAO_PJ IS

/******************************************************************************************
 Objetos criados: SPS_SQ_A8_NU_SEQU_REME_PJ
 Descricao:       Busca valor da sequence SQ_A8_NU_SEQU_REME_PJ.
 Autor:           Ivan Tabarino
 Data:            23/02/2012
******************************************************************************************/
PROCEDURE SPS_SQ_A8_NU_SEQU_REME_PJ(
   p_num_out_RETORNO  OUT  NUMBER
)

IS

BEGIN

   SELECT A8.SQ_A8_NU_SEQU_REME_PJ.NEXTVAL
   INTO   p_num_out_RETORNO
   FROM   DUAL;


END SPS_SQ_A8_NU_SEQU_REME_PJ;

/******************************************************************************************
 Objetos criados: SPS_TB_PRODUTO
 Descricao:       Busca valor da sequence TB_PRODUTO.
 Autor:           Ivan Tabarino
 Data:            27/02/2012
******************************************************************************************/
PROCEDURE SPS_TB_PRODUTO(
   P_CO_EMPR            IN A8.TB_EMPRESA_HO.CO_EMPR%TYPE,
   P_CO_PROD            IN A8.TB_PRODUTO.CO_PROD%TYPE,
   P_VA_MINI_MAIR_VALO OUT A8.TB_PRODUTO.VA_MINI_MAIR_VALO%TYPE
)

IS

BEGIN

   SELECT   PROD.VA_MINI_MAIR_VALO
   INTO     P_VA_MINI_MAIR_VALO
   FROM     A8.TB_PRODUTO       PROD,
            A8.TB_EMPRESA_HO    EMPR
   WHERE    PROD.CO_EMPR_FUSI   = EMPR.CO_EMPR_FUSI
   AND      PROD.CO_PROD        = P_CO_PROD
   AND      EMPR.CO_EMPR        = P_CO_EMPR;


END SPS_TB_PRODUTO;

/******************************************************************************************
 Objetos criados: SPI_TB_HIST_ENVI_INFO_GEST
 Descricao:       Insere dados na tabela TB_HIST_ENVI_INFO_GEST_CAIX
 Autor:           Ivan Tabarino
 Data:            23/02/2012
******************************************************************************************/
PROCEDURE SPI_TB_HIST_ENVI_INFO_GEST(
   P_NU_SEQU_OPER_ATIV       IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.NU_SEQU_OPER_ATIV%TYPE,
   P_CO_SITU_MOVI_GEST_CAIX  IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_SITU_MOVI_GEST_CAIX%TYPE,
   P_CO_TEXT_XML             IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_TEXT_XML%TYPE
)
IS

BEGIN

   INSERT INTO A8.TB_HIST_ENVI_INFO_GEST_CAIX (NU_SEQU_OPER_ATIV,
                                               DH_ENVI_GEST_CAIX,
                                               CO_SITU_MOVI_GEST_CAIX,
                                               CO_TEXT_XML)
   VALUES                                     (P_NU_SEQU_OPER_ATIV,
                                               SYSDATE,
                                               P_CO_SITU_MOVI_GEST_CAIX,
                                               P_CO_TEXT_XML);
                           

EXCEPTION
   WHEN dup_val_on_index THEN
      raise_application_error(-20000, 'Chave única já existe.', TRUE);
   WHEN OTHERS THEN
      raise_application_error(-20999, 'Erro SPI_TB_HIST_ENVI_INFO_GEST: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
      
END SPI_TB_HIST_ENVI_INFO_GEST;
  
END PKG_A8_INTEGRACAO_PJ;
/
