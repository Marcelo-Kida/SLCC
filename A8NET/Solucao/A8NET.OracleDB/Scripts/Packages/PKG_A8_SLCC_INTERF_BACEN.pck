CREATE OR REPLACE PACKAGE PKG_A8_SLCC_INTERF_BACEN IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE A8F_ISPB_COMP
(  p_vch_COD_COMP         IN A8.TB_INSTITUICAO_SPB.CO_ISPB%TYPE,
   p_num_VL_TP_CONSULTA   IN A8.TB_INSTITUICAO_SPB.IN_ENVI_CRIP%TYPE,
   ps_return              OUT VARCHAR2,
   ps_statusCode          OUT NUMBER,
   ps_statusDescription   OUT VARCHAR2);

END PKG_A8_SLCC_INTERF_BACEN;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_SLCC_INTERF_BACEN IS

/********************************************************************************************************
Nome Lógico     :  PKG_A8_SLCC_INTERF_BACEN
Descrição       :  Function de consulta de registros na tabela TB_INSTITUICAO_SPB
Retorno         :  Código ISPB ou Código de Compensação
Autor           :  Cleber Santos
Data Criação    :  10/06/2011 17:52:13
Comentario      :  -
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE A8F_ISPB_COMP(
   p_vch_COD_COMP            IN A8.TB_INSTITUICAO_SPB.CO_ISPB%TYPE,
   p_num_VL_TP_CONSULTA      IN A8.TB_INSTITUICAO_SPB.IN_ENVI_CRIP%TYPE,
   ps_return                 OUT VARCHAR2,
   ps_statusCode             OUT NUMBER,
   ps_statusDescription      OUT VARCHAR2
)
    
IS

   v_vch_RESULTADO     CHAR(8);

BEGIN

   IF p_num_VL_TP_CONSULTA = 1 THEN

      SELECT CO_ISPB INTO v_vch_RESULTADO
      FROM   A8.TB_INSTITUICAO_SPB
      WHERE  TRUNC(DT_INIC_VIGE) <= TRUNC(SYSDATE)
	         AND    (TRUNC(DT_FIM_VIGE) >= TRUNC(SYSDATE) OR DT_FIM_VIGE IS NULL)
	         AND    CO_CPEN = p_vch_COD_COMP;

   ELSIF p_num_VL_TP_CONSULTA = 2 THEN

      SELECT CO_CPEN INTO v_vch_RESULTADO
      FROM   A8.TB_INSTITUICAO_SPB
      WHERE  TRUNC(DT_INIC_VIGE) <= TRUNC(SYSDATE)
	         AND    (TRUNC(DT_FIM_VIGE) >= TRUNC(SYSDATE) OR DT_FIM_VIGE IS NULL)
	         AND    CO_ISPB = p_vch_COD_COMP;

   END IF;

   ps_return := v_vch_RESULTADO;
   ps_statusCode := 0;
   ps_statusDescription := '';
    
  EXCEPTION
    WHEN NO_DATA_FOUND THEN
      --RAISE_APPLICATION_ERROR(-20001, 'Execucao - Registro nao encontrado!'  || SQLCODE || ' - ' || SQLERRM);
      ps_return := p_vch_COD_COMP;
      ps_statusCode := 3;
      ps_statusDescription := 'Execucao - Registro nao encontrado!'  || SQLCODE || ' - ' || SQLERRM;
      
    WHEN OTHERS THEN
      --RAISE_APPLICATION_ERROR(-20006, 'Execucao - Erro na execucao da Function A8F_ISPB_COMP '  || SQLCODE || ' - ' || SQLERRM);
      ps_return := p_vch_COD_COMP;
      ps_statusCode := 9;
      ps_statusDescription := 'Cod : 20006 - Execucao - Erro na execucao da Function A8F_ISPB_COMP '  || SQLCODE || ' - ' || SQLERRM;
      
END A8F_ISPB_COMP;

END PKG_A8_SLCC_INTERF_BACEN;
/
