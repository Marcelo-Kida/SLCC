CREATE OR REPLACE PACKAGE PKG_A8_CONCILIACAO_OPER_MESG IS

  -- Author  : Bruno Oliveira
  -- Created : 16/03/2012
  -- Purpose : Realizar todas as operações necessárias (inserts, updates, selects) 
  --           para conciliação de Operações e Mensagens SPB
  -- Tabelas : A8.TB_CNCL_OPER_ATIV
  --           A8.TB_MESG_RECB_SPB_CNCL 
  --           A8.TB_JUST_CNCL_OPER_ATIV_MESG

  -- Public type declarations
  TYPE tp_cursor IS REF CURSOR;

  PROCEDURE SPI_TB_CNCL_OPER_ATIV
  (
     p_NU_SEQU_CNCL_OPER_ATIV_MESG   IN A8.TB_CNCL_OPER_ATIV.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE,
     p_NU_SEQU_OPER_ATIV             IN A8.TB_CNCL_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
     p_NU_CTRL_IF                    IN A8.TB_CNCL_OPER_ATIV.NU_CTRL_IF%TYPE,
     p_DH_REGT_MESG_SPB              IN A8.TB_CNCL_OPER_ATIV.DH_REGT_MESG_SPB%TYPE,
     p_QT_ATIV_MERC_CNCL             IN A8.TB_CNCL_OPER_ATIV.QT_ATIV_MERC_CNCL%TYPE,
     p_NU_SEQU_CNTR_REPE             IN A8.TB_CNCL_OPER_ATIV.NU_SEQU_CNTR_REPE%TYPE
  );

  PROCEDURE SPS_TB_CNCL_OPER_ATIV(
     p_NU_SEQU_CNCL_OPER_ATIV_MESG   IN  A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE,
     p_cur_out_CURSOR                OUT tp_cursor
  );
  
  PROCEDURE SPS_TB_CNCL_OPER_ATIV_02(
     p_NU_SEQU_OPER_ATIV              IN A8.TB_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
		 p_cur_out_CURSOR                 OUT tp_cursor
  );
  
  PROCEDURE SPI_TB_JUST_CNCL_OPERATIV_MESG(
      p_NU_SEQU_CNCL_OPER_ATIV_MESG   OUT A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE,
      p_TP_JUST_CNCL                  IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.TP_JUST_CNCL%TYPE,
      p_TX_JUST                       IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.TX_JUST%TYPE,
      p_CO_USUA_ATLZ                  IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.CO_USUA_ATLZ%TYPE,
      p_CO_ETCA_TRAB_ATLZ             IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.CO_ETCA_TRAB_ATLZ%TYPE
  );
  
  PROCEDURE CONCILIAR_OPERACAO_X_BMC0015(
      p_NU_SEQU_OPER_ATIV             IN A8.TB_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE
     ,p_CO_ULTI_SITU_PROC_OPERACAO    IN A8.TB_OPER_ATIV.CO_ULTI_SITU_PROC%TYPE
     ,p_CO_ULTI_SITU_PROC_MSGSPB      IN A8.TB_MESG_RECB_ENVI_SPB.CO_ULTI_SITU_PROC%TYPE
     ,p_NU_SEQU_CNCL_OPER_ATIV_MESG   OUT A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE
     ,p_RETORNO                       OUT NUMBER
  );
   
END PKG_A8_CONCILIACAO_OPER_MESG;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_CONCILIACAO_OPER_MESG IS

-----------------------------------------------------------------------------
-- NOME       : SPI_TB_CNCL_OPER_ATIV
-----------------------------------------------------------------------------
-- AUTOR      : Bruno Oliveira
-- CRIADO     : 16/03/2012
--
-- OBJETIVO   : selecionar registros da tabela A8.TB_MESG_RECB_SPB_CNCL
-- PARAMETROS : ...
-- RETORNO    : ...
--
-- ARQUIVOS E TABELAS UTILIZADOS:
-- ...
-----------------------------------------------------------------------------
-- HISTORICO
-- DATA        AUTOR  DETALHES
-----------------------------------------------------------------------------
PROCEDURE SPI_TB_CNCL_OPER_ATIV
(
   p_NU_SEQU_CNCL_OPER_ATIV_MESG   IN A8.TB_CNCL_OPER_ATIV.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE,
   p_NU_SEQU_OPER_ATIV             IN A8.TB_CNCL_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
   p_NU_CTRL_IF                    IN A8.TB_CNCL_OPER_ATIV.NU_CTRL_IF%TYPE,
   p_DH_REGT_MESG_SPB              IN A8.TB_CNCL_OPER_ATIV.DH_REGT_MESG_SPB%TYPE,
   p_QT_ATIV_MERC_CNCL             IN A8.TB_CNCL_OPER_ATIV.QT_ATIV_MERC_CNCL%TYPE,
   p_NU_SEQU_CNTR_REPE             IN A8.TB_CNCL_OPER_ATIV.NU_SEQU_CNTR_REPE%TYPE
) IS

BEGIN

  INSERT INTO A8.TB_CNCL_OPER_ATIV(
    	nu_sequ_cncl_oper_ativ_mesg, 
      nu_sequ_oper_ativ, 
      nu_ctrl_if, 
      dh_regt_mesg_spb, 
      qt_ativ_merc_cncl, 
      nu_sequ_cntr_repe
  )
  VALUES(
     p_NU_SEQU_CNCL_OPER_ATIV_MESG,
     p_NU_SEQU_OPER_ATIV,
     p_NU_CTRL_IF,
     p_DH_REGT_MESG_SPB,
     p_QT_ATIV_MERC_CNCL,
     p_NU_SEQU_CNTR_REPE	
  );

EXCEPTION
 WHEN OTHERS THEN
    raise_application_error(-20999, 'Erro SPI_TB_CNCL_OPER_ATIV: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
  
END SPI_TB_CNCL_OPER_ATIV;
  

-----------------------------------------------------------------------------
-- NOME       : SPS_TB_CNCL_OPER_ATIV
-----------------------------------------------------------------------------
-- AUTOR      : Bruno Oliveira
-- CRIADO     : 16/03/2012
--
-- OBJETIVO   : selecionar registros da tabela A8.TB_MESG_RECB_SPB_CNCL
-- PARAMETROS : ...
-- RETORNO    : ...
--
-- ARQUIVOS E TABELAS UTILIZADOS:
-- ...
-----------------------------------------------------------------------------
-- HISTORICO
-- DATA        AUTOR  DETALHES
-----------------------------------------------------------------------------
PROCEDURE SPS_TB_CNCL_OPER_ATIV(
   p_NU_SEQU_CNCL_OPER_ATIV_MESG  IN  A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE,
   p_cur_out_CURSOR               OUT tp_cursor
) IS

BEGIN

 OPEN p_cur_out_CURSOR FOR

   SELECT
      nu_sequ_cncl_oper_ativ_mesg, 
      nu_sequ_oper_ativ, 
      nu_ctrl_if, 
      dh_regt_mesg_spb, 
      qt_ativ_merc_cncl, 
      nu_sequ_cntr_repe
   
   FROM A8.TB_CNCL_OPER_ATIV 
   
   WHERE nu_sequ_cncl_oper_ativ_mesg = p_NU_SEQU_CNCL_OPER_ATIV_MESG
   ;

EXCEPTION
 WHEN OTHERS THEN
    raise_application_error(-20999, 'Erro SPS_TB_CNCL_OPER_ATIV: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);

END SPS_TB_CNCL_OPER_ATIV;


-----------------------------------------------------------------------------
-- NOME       : SPS_TB_CNCL_OPER_ATIV_02
-----------------------------------------------------------------------------
-- AUTOR      : Bruno Oliveira
-- CRIADO     : 28/05/2012
--
-- OBJETIVO   : selecionar registros da tabela A8.TB_MESG_RECB_SPB_CNCL
--              a partir de uma Operacao
-- PARAMETROS : ...
-- RETORNO    : ...
--
-- ARQUIVOS E TABELAS UTILIZADOS:
-- ...
-----------------------------------------------------------------------------
-- HISTORICO
-- DATA        AUTOR  DETALHES
-----------------------------------------------------------------------------
PROCEDURE SPS_TB_CNCL_OPER_ATIV_02(
   p_NU_SEQU_OPER_ATIV              IN A8.TB_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
   p_cur_out_CURSOR                 OUT tp_cursor
) IS

BEGIN

 OPEN p_cur_out_CURSOR FOR

   SELECT
      nu_sequ_cncl_oper_ativ_mesg, 
      nu_sequ_oper_ativ, 
      nu_ctrl_if, 
      dh_regt_mesg_spb, 
      qt_ativ_merc_cncl, 
      nu_sequ_cntr_repe
   
   FROM A8.TB_CNCL_OPER_ATIV 
   
   WHERE nu_sequ_oper_ativ  = p_NU_SEQU_OPER_ATIV
   ;

EXCEPTION
 WHEN OTHERS THEN
    raise_application_error(-20999, 'Erro SPS_TB_CNCL_OPER_ATIV_02: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);

END SPS_TB_CNCL_OPER_ATIV_02;



-----------------------------------------------------------------------------
-- NOME       : SPI_TB_JUST_CNCL_OPERATIV_MESG
-----------------------------------------------------------------------------
-- AUTOR      : Bruno Oliveira
-- CRIADO     : 16/03/2012
--
-- OBJETIVO   : selecionar registros da tabela A8.TB_MESG_RECB_SPB_CNCL
-- PARAMETROS : ...
-- RETORNO    : ...
--
-- ARQUIVOS E TABELAS UTILIZADOS:
-- ...
-----------------------------------------------------------------------------
-- HISTORICO
-- DATA        AUTOR  DETALHES
-----------------------------------------------------------------------------
PROCEDURE SPI_TB_JUST_CNCL_OPERATIV_MESG(
    p_NU_SEQU_CNCL_OPER_ATIV_MESG   OUT A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE,
    p_TP_JUST_CNCL                  IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.TP_JUST_CNCL%TYPE,
    p_TX_JUST                       IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.TX_JUST%TYPE,
    p_CO_USUA_ATLZ                  IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.CO_USUA_ATLZ%TYPE,
    p_CO_ETCA_TRAB_ATLZ             IN A8.TB_JUST_CNCL_OPER_ATIV_MESG.CO_ETCA_TRAB_ATLZ%TYPE
) IS

BEGIN

  INSERT INTO A8.TB_JUST_CNCL_OPER_ATIV_MESG(
    nu_sequ_cncl_oper_ativ_mesg, 
    tp_just_cncl, 
    tx_just, 
    co_usua_atlz, 
    co_etca_trab_atlz, 
    dh_just_cncl
  )
  VALUES(
    A8.SQ_A8_NU_SEQU_CNCL_OPER_MESG.NEXTVAL,
    p_TP_JUST_CNCL,             
    p_TX_JUST,                  
    p_CO_USUA_ATLZ,             
    p_CO_ETCA_TRAB_ATLZ,        
    SYSDATE             
  ); 
  
  SELECT A8.SQ_A8_NU_SEQU_CNCL_OPER_MESG.CURRVAL INTO p_NU_SEQU_CNCL_OPER_ATIV_MESG FROM DUAL;

EXCEPTION
 WHEN OTHERS THEN
    raise_application_error(-20999, 'Erro SPI_TB_JUST_CNCL_OPERATIV_MESG: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);

END SPI_TB_JUST_CNCL_OPERATIV_MESG;


-----------------------------------------------------------------------------
-- NOME       : CONCILIAR_OPERACAO_BMC0015
-----------------------------------------------------------------------------
-- AUTOR      : Bruno Oliveira
-- CRIADO     : 27/03/2012
--
-- OBJETIVO   : verificar se operação concilia com alguma BMC0015
-- PARAMETROS : ...
-- p_RETORNO  : 0 = conciliação NOK
--              1 = conciliação OK
--             >1 conciliação NOK devido existir mais de uma BMC0015 a conciliar
-- p_NU_SEQU_CNCL_OPER_ATIV_MESG é retornado <> 0 somente se p_RETORNO = 1
-----------------------------------------------------------------------------
-- HISTORICO
-- DATA        AUTOR  DETALHES
-----------------------------------------------------------------------------
PROCEDURE CONCILIAR_OPERACAO_X_BMC0015(
    p_NU_SEQU_OPER_ATIV             IN A8.TB_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE
   ,p_CO_ULTI_SITU_PROC_OPERACAO    IN A8.TB_OPER_ATIV.CO_ULTI_SITU_PROC%TYPE
   ,p_CO_ULTI_SITU_PROC_MSGSPB      IN A8.TB_MESG_RECB_ENVI_SPB.CO_ULTI_SITU_PROC%TYPE
   ,p_NU_SEQU_CNCL_OPER_ATIV_MESG   OUT A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE
   ,p_RETORNO                       OUT NUMBER
) IS

   -- VARIAVEIS
   V_NU_SEQU_CNCL_OPER_ATIV_MESG  A8.TB_JUST_CNCL_OPER_ATIV_MESG.NU_SEQU_CNCL_OPER_ATIV_MESG%TYPE;
   V_NU_CTRL_IF                   A8.TB_MESG_RECB_SPB_CNCL.NU_CTRL_IF%TYPE;
   V_DH_REGT_MESG_SPB             A8.TB_MESG_RECB_SPB_CNCL.DH_REGT_MESG_SPB%TYPE;
   V_NU_SEQU_CNTR_REPE            A8.TB_MESG_RECB_SPB_CNCL.NU_SEQU_CNTR_REPE%TYPE;

   -- CURSORES
   CURSOR curConciliacao (p_SITUACAO_OPERACAO A8.TB_OPER_ATIV.CO_ULTI_SITU_PROC%TYPE
                         ,p_SITUACAO_MSGSPB   A8.TB_MESG_RECB_ENVI_SPB.CO_ULTI_SITU_PROC%TYPE)
   IS
     SELECT 
      C.NU_CTRL_IF, C.DH_REGT_MESG_SPB, C.NU_SEQU_CNTR_REPE
   FROM 
        A8.TB_OPER_ATIV          O
       ,A8.TB_MESG_RECB_SPB_CNCL C
       ,A8.TB_MESG_RECB_ENVI_SPB M
   WHERE
     ----- filtros ---------------
         O.NU_SEQU_OPER_ATIV = p_NU_SEQU_OPER_ATIV
     AND M.CO_MESG_SPB       = 'BMC0015'
     AND O.CO_ULTI_SITU_PROC = p_SITUACAO_OPERACAO
     AND M.CO_ULTI_SITU_PROC = p_SITUACAO_MSGSPB
     ----- campos conciliação -----
     AND O.CD_ASSO_CAMB      = C.CD_ASSO_CAMB
     AND O.DT_OPER_ATIV      = C.DT_OPER
     AND O.IN_OPER_DEBT_CRED = C.IN_OPER_DEBT_CRED
     AND O.PE_TAXA_NEGO      = C.PE_TAXA_NEGO
     AND O.VA_OPER_ATIV      = C.VA_FINC
     AND O.VA_MOED_ESTR      = C.VA_MOED_ESTR
     AND O.DT_LIQU_OPER_ATIV = C.DT_LIQU
     ----- relacionamentos --------
     AND C.NU_CTRL_IF        = M.NU_CTRL_IF
     AND C.DH_REGT_MESG_SPB  = M.DH_REGT_MESG_SPB
     AND C.NU_SEQU_CNTR_REPE = M.NU_SEQU_CNTR_REPE;

BEGIN  

    p_RETORNO := 0;
    p_NU_SEQU_CNCL_OPER_ATIV_MESG := 0;

    FOR rowConciliacao IN curConciliacao(p_CO_ULTI_SITU_PROC_OPERACAO, p_CO_ULTI_SITU_PROC_MSGSPB) LOOP
    
       p_RETORNO := curConciliacao%ROWCOUNT;
    
       IF p_RETORNO = 1 THEN
          V_NU_CTRL_IF := rowConciliacao.Nu_Ctrl_If;
          V_DH_REGT_MESG_SPB := rowConciliacao.Dh_Regt_Mesg_Spb;  
          V_NU_SEQU_CNTR_REPE := rowConciliacao.Nu_Sequ_Cntr_Repe;
       END IF;
    
    END LOOP;

    -- faz conciliação somente se tiver conciliado com apenas uma BMC0015
    IF p_RETORNO = 1 THEN
          
          -- gera conciliação          
          SPI_TB_JUST_CNCL_OPERATIV_MESG(V_NU_SEQU_CNCL_OPER_ATIV_MESG, 
                                         NULL, 
                                         NULL, 
                                         'SISTEMA', 
                                         'SERVIDOR');

          -- concilia Operação com a Mensagem SPB
          SPI_TB_CNCL_OPER_ATIV(V_NU_SEQU_CNCL_OPER_ATIV_MESG, 
                                p_NU_SEQU_OPER_ATIV, 
                                V_NU_CTRL_IF,
                                V_DH_REGT_MESG_SPB, 
                                0,
                                V_NU_SEQU_CNTR_REPE);
           
           -- retorna PK da conciliação gerada
           p_NU_SEQU_CNCL_OPER_ATIV_MESG := V_NU_SEQU_CNCL_OPER_ATIV_MESG;
    
    END IF;

EXCEPTION
 WHEN OTHERS THEN
    raise_application_error(-20999, 'Erro CONCILIAR_OPERACAO_X_BMC0015: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);

END CONCILIAR_OPERACAO_X_BMC0015;


END PKG_A8_CONCILIACAO_OPER_MESG;
/
