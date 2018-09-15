CREATE OR REPLACE PACKAGE PKG_A8_TB_HIST_SITU_ACAO_MESG IS

-- Author  : MAPS
-- Created : 08/03/2011 17:47:15
-- Purpose : Realizar operações basicas de Insert, Update, Delete e Select na tabela TB_HIST_SITU_ACAO_MESG_SPB

-- Public type declarations
TYPE tp_cursor IS REF CURSOR;

-- Public constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Public variable declarations
--<VariableName> <Datatype>;

-- Public function and procedure declarations
PROCEDURE SPI_TB_HIST_SITU_ACAO_MESG_SPB(
	P_NU_SEQU_CNTR_REPE	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_SEQU_CNTR_REPE%TYPE default null,
	P_NU_CTRL_IF	            IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE default null,
	P_DH_REGT_MESG_SPB	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_REGT_MESG_SPB%TYPE default null,
	P_DH_SITU_ACAO_MESG_SPB	  IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE default null,
	P_CO_SITU_PROC	          IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_SITU_PROC%TYPE default null,
	P_TP_ACAO_MESG_SPB	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TP_ACAO_MESG_SPB%TYPE default null,
	P_TX_CNTD_ANTE_ACAO	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TX_CNTD_ANTE_ACAO%TYPE default null,
	P_CO_USUA_ULTI_ATLZ	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_USUA_ULTI_ATLZ%TYPE default null,
	P_CO_ETCA_TRAB_ULTI_ATLZ	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_ETCA_TRAB_ULTI_ATLZ%TYPE default null
);

PROCEDURE SPS_TB_HIST_SITU_ACAO_MESG_SPB(
  p_curselect OUT tp_cursor
);

PROCEDURE SPS_TB_HIST_MAX
(
  P_NU_CTRL_IF              IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE,
  P_DH_SITU_ACAO_MESG_SPB   OUT A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE
);

PROCEDURE SPU_TB_HIST_SITU_ACAO_MESG_SPB(
	P_NU_SEQU_CNTR_REPE	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_SEQU_CNTR_REPE%TYPE,
	P_NU_CTRL_IF	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE,
	P_DH_REGT_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_REGT_MESG_SPB%TYPE,
	P_DH_SITU_ACAO_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE,
	P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_SITU_PROC%TYPE,
	P_TP_ACAO_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TP_ACAO_MESG_SPB%TYPE,
	P_TX_CNTD_ANTE_ACAO	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TX_CNTD_ANTE_ACAO%TYPE,
	P_CO_USUA_ULTI_ATLZ	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_USUA_ULTI_ATLZ%TYPE,
	P_CO_ETCA_TRAB_ULTI_ATLZ	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_ETCA_TRAB_ULTI_ATLZ%TYPE
);

PROCEDURE SPE_TB_HIST_SITU_ACAO_MESG_SPB(
   P_NU_SEQU_CNTR_REPE	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_SEQU_CNTR_REPE%TYPE,
   P_NU_CTRL_IF	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE,
   P_DH_REGT_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_REGT_MESG_SPB%TYPE,
   P_DH_SITU_ACAO_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE,
   P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_SITU_PROC%TYPE
);

END PKG_A8_TB_HIST_SITU_ACAO_MESG;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_TB_HIST_SITU_ACAO_MESG IS

-- Private type declarations
--type <TypeName> is <Datatype>;
-- Private constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Private variable declarations
--<VariableName> <Datatype>;

-- Function and procedure implementations

/********************************************************************************************************
Nome Lógico     :	SPI_TB_HIST_SITU_ACAO_MESG_SPB
Descrição       :	Procedure de inclusão de registros na tabela TB_HIST_SITU_ACAO_MESG_SPB 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 17:47:15 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPI_TB_HIST_SITU_ACAO_MESG_SPB(
 	P_NU_SEQU_CNTR_REPE	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_SEQU_CNTR_REPE%TYPE default null,
	P_NU_CTRL_IF	            IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE default null,
	P_DH_REGT_MESG_SPB	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_REGT_MESG_SPB%TYPE default null,
	P_DH_SITU_ACAO_MESG_SPB	  IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE default null,
	P_CO_SITU_PROC	          IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_SITU_PROC%TYPE default null,
	P_TP_ACAO_MESG_SPB	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TP_ACAO_MESG_SPB%TYPE default null,
	P_TX_CNTD_ANTE_ACAO	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TX_CNTD_ANTE_ACAO%TYPE default null,
	P_CO_USUA_ULTI_ATLZ	      IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_USUA_ULTI_ATLZ%TYPE default null,
	P_CO_ETCA_TRAB_ULTI_ATLZ	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_ETCA_TRAB_ULTI_ATLZ%TYPE default null
  )

IS

BEGIN

  INSERT INTO A8.TB_HIST_SITU_ACAO_MESG_SPB(
		NU_SEQU_CNTR_REPE,
		NU_CTRL_IF,
		DH_REGT_MESG_SPB,
		DH_SITU_ACAO_MESG_SPB,
		CO_SITU_PROC,
		TP_ACAO_MESG_SPB, 
		TX_CNTD_ANTE_ACAO, 
		CO_USUA_ULTI_ATLZ, 
		CO_ETCA_TRAB_ULTI_ATLZ 
  )
  VALUES(
		P_NU_SEQU_CNTR_REPE,
		P_NU_CTRL_IF,
    P_DH_REGT_MESG_SPB,
		P_DH_SITU_ACAO_MESG_SPB,
		P_CO_SITU_PROC,
		P_TP_ACAO_MESG_SPB,
		P_TX_CNTD_ANTE_ACAO,
		P_CO_USUA_ULTI_ATLZ,
		P_CO_ETCA_TRAB_ULTI_ATLZ
  );
END SPI_TB_HIST_SITU_ACAO_MESG_SPB;


/********************************************************************************************************
Nome Lógico     :	SPS_TB_HIST_SITU_ACAO_MESG_SPB
Descrição       :	Seleciona os dados da tabela TB_HIST_SITU_ACAO_MESG_SPB 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 17:47:15 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPS_TB_HIST_SITU_ACAO_MESG_SPB(
  p_curselect OUT tp_cursor
) IS        

BEGIN 
 
	  -- ***** Seleciona os Dados ***** 
  OPEN p_curselect FOR 
  SELECT  
		a.NU_SEQU_CNTR_REPE,
		a.NU_CTRL_IF,
		a.DH_REGT_MESG_SPB,
		a.DH_SITU_ACAO_MESG_SPB,
		a.CO_SITU_PROC,
		a.TP_ACAO_MESG_SPB,
		a.TX_CNTD_ANTE_ACAO,
		a.CO_USUA_ULTI_ATLZ,
		a.CO_ETCA_TRAB_ULTI_ATLZ
  	FROM  
  		A8.TB_HIST_SITU_ACAO_MESG_SPB a;
END SPS_TB_HIST_SITU_ACAO_MESG_SPB; 

/********************************************************************************************************
Nome Lógico     :	SPS_TB_HIST_SITU_ACAO_MESG_SPB
Descrição       :	Seleciona os dados da tabela TB_HIST_SITU_ACAO_MESG_SPB 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 17:47:15 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPS_TB_HIST_MAX
(
  P_NU_CTRL_IF              IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE,
  P_DH_SITU_ACAO_MESG_SPB   OUT A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE
) 
IS        

BEGIN 
 
  SELECT  
		MAX(a.DH_SITU_ACAO_MESG_SPB) INTO P_DH_SITU_ACAO_MESG_SPB
  FROM  A8.TB_HIST_SITU_ACAO_MESG_SPB a
    WHERE NU_CTRL_IF = P_NU_CTRL_IF;
    
END SPS_TB_HIST_MAX; 


/********************************************************************************************************
Nome Lógico     :	SPU_TB_HIST_SITU_ACAO_MESG_SPB
Descrição       :	Procedure de atualização de registros na tabela TB_HIST_SITU_ACAO_MESG_SPB 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 17:47:15 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPU_TB_HIST_SITU_ACAO_MESG_SPB(
	P_NU_SEQU_CNTR_REPE	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_SEQU_CNTR_REPE%TYPE,
	P_NU_CTRL_IF	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE,
	P_DH_REGT_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_REGT_MESG_SPB%TYPE,
	P_DH_SITU_ACAO_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE,
	P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_SITU_PROC%TYPE,
	P_TP_ACAO_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TP_ACAO_MESG_SPB%TYPE,
	P_TX_CNTD_ANTE_ACAO	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.TX_CNTD_ANTE_ACAO%TYPE,
	P_CO_USUA_ULTI_ATLZ	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_USUA_ULTI_ATLZ%TYPE,
	P_CO_ETCA_TRAB_ULTI_ATLZ	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_ETCA_TRAB_ULTI_ATLZ%TYPE
) IS 
BEGIN 
 
	UPDATE A8.TB_HIST_SITU_ACAO_MESG_SPB SET
		TP_ACAO_MESG_SPB = P_TP_ACAO_MESG_SPB,
		TX_CNTD_ANTE_ACAO = P_TX_CNTD_ANTE_ACAO,
		CO_USUA_ULTI_ATLZ = P_CO_USUA_ULTI_ATLZ,
		CO_ETCA_TRAB_ULTI_ATLZ = P_CO_ETCA_TRAB_ULTI_ATLZ
	WHERE 
		NU_SEQU_CNTR_REPE = P_NU_SEQU_CNTR_REPE
		AND NU_CTRL_IF = P_NU_CTRL_IF
		AND DH_REGT_MESG_SPB = P_DH_REGT_MESG_SPB
		AND DH_SITU_ACAO_MESG_SPB = P_DH_SITU_ACAO_MESG_SPB
		AND CO_SITU_PROC = P_CO_SITU_PROC
  ;
END SPU_TB_HIST_SITU_ACAO_MESG_SPB; 


/********************************************************************************************************
Nome Lógico     :	SPE_TB_HIST_SITU_ACAO_MESG_SPB
Descrição       :	Procedure de exclusao de registros da tabela TB_HIST_SITU_ACAO_MESG_SPB 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 17:47:15 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPE_TB_HIST_SITU_ACAO_MESG_SPB(
   P_NU_SEQU_CNTR_REPE	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_SEQU_CNTR_REPE%TYPE,
   P_NU_CTRL_IF	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.NU_CTRL_IF%TYPE,
   P_DH_REGT_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_REGT_MESG_SPB%TYPE,
   P_DH_SITU_ACAO_MESG_SPB	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.DH_SITU_ACAO_MESG_SPB%TYPE,
   P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_MESG_SPB.CO_SITU_PROC%TYPE
) IS
BEGIN

	DELETE FROM
		A8.TB_HIST_SITU_ACAO_MESG_SPB
	WHERE
		NU_SEQU_CNTR_REPE = p_NU_SEQU_CNTR_REPE
		AND NU_CTRL_IF = p_NU_CTRL_IF
		AND DH_REGT_MESG_SPB = p_DH_REGT_MESG_SPB
		AND DH_SITU_ACAO_MESG_SPB = p_DH_SITU_ACAO_MESG_SPB
		AND CO_SITU_PROC = p_CO_SITU_PROC;
  
END SPE_TB_HIST_SITU_ACAO_MESG_SPB;


END PKG_A8_TB_HIST_SITU_ACAO_MESG;
/
