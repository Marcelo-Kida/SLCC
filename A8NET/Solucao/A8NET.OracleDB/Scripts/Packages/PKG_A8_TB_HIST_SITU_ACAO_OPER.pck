CREATE OR REPLACE PACKAGE PKG_A8_TB_HIST_SITU_ACAO_OPER IS

-- Author  : MAPS
-- Created : 08/03/2011 15:59:13
-- Purpose : Realizar operações basicas de Insert, Update, Delete e Select na tabela TB_HIST_SITU_ACAO_OPER_ATIV

-- Public type declarations
TYPE tp_cursor IS REF CURSOR;

-- Public constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Public variable declarations
--<VariableName> <Datatype>;

-- Public function and procedure declarations
PROCEDURE SPI_TB_HIST_SITU_ACAO_OPER(
	P_NU_SEQU_OPER_ATIV IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
	P_DH_SITU_ACAO_OPER_ATIV IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.DH_SITU_ACAO_OPER_ATIV%TYPE,
	P_CO_SITU_PROC           IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE,
	P_TP_ACAO_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TP_ACAO_OPER_ATIV%TYPE DEFAULT NULL,
	P_TP_JUST_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TP_JUST_SITU_PROC%TYPE DEFAULT NULL,
	P_TX_CNTD_ANTE_ACAO	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TX_CNTD_ANTE_ACAO%TYPE DEFAULT NULL,
	P_CO_USUA_ATLZ	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_USUA_ATLZ%TYPE,
	P_CO_ETCA_USUA_ATLZ	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_ETCA_USUA_ATLZ%TYPE
);

PROCEDURE SPU_TB_HIST_SITU_ACAO_OPER(
	P_NU_SEQU_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
	P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE,
  P_TP_JUST_SITU_PROC IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TP_JUST_SITU_PROC%TYPE
);

PROCEDURE SPE_TB_HIST_SITU_ACAO_OPER(
   P_NU_SEQU_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
   P_DH_SITU_ACAO_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.DH_SITU_ACAO_OPER_ATIV%TYPE,
   P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE
);

PROCEDURE SPS_TB_HIST_SITU_ACAO_OPER(
  p_curselect OUT tp_cursor
);

PROCEDURE SPS_TB_HIST_MAX(
  P_NU_SEQU_OPER_ATIV       IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
  P_CO_SITU_PROC            IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE,
  P_DH_SITU_ACAO_OPER_ATIV   OUT A8.TB_HIST_SITU_ACAO_OPER_ATIV.DH_SITU_ACAO_OPER_ATIV%TYPE
);
END PKG_A8_TB_HIST_SITU_ACAO_OPER;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_TB_HIST_SITU_ACAO_OPER IS

-- Private type declarations
--type <TypeName> is <Datatype>;
-- Private constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Private variable declarations
--<VariableName> <Datatype>;

-- Function and procedure implementations

/********************************************************************************************************
Nome Lógico     :	SPI_TB_HIST_SITU_ACAO_OPER_ATIV
Descrição       :	Procedure de inclusão de registros na tabela TB_HIST_SITU_ACAO_OPER_ATIV 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:13 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPI_TB_HIST_SITU_ACAO_OPER(
	P_NU_SEQU_OPER_ATIV IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
	P_DH_SITU_ACAO_OPER_ATIV IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.DH_SITU_ACAO_OPER_ATIV%TYPE,
	P_CO_SITU_PROC           IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE,
	P_TP_ACAO_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TP_ACAO_OPER_ATIV%TYPE DEFAULT NULL,
	P_TP_JUST_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TP_JUST_SITU_PROC%TYPE DEFAULT NULL,
	P_TX_CNTD_ANTE_ACAO	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TX_CNTD_ANTE_ACAO%TYPE DEFAULT NULL,
	P_CO_USUA_ATLZ	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_USUA_ATLZ%TYPE,
	P_CO_ETCA_USUA_ATLZ	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_ETCA_USUA_ATLZ%TYPE
) IS
BEGIN

  INSERT INTO A8.TB_HIST_SITU_ACAO_OPER_ATIV(
		NU_SEQU_OPER_ATIV,
		DH_SITU_ACAO_OPER_ATIV,
		CO_SITU_PROC,
		TP_ACAO_OPER_ATIV, 
		TP_JUST_SITU_PROC, 
		TX_CNTD_ANTE_ACAO, 
		CO_USUA_ATLZ, 
		CO_ETCA_USUA_ATLZ 
  )
  VALUES(
		P_NU_SEQU_OPER_ATIV,
		P_DH_SITU_ACAO_OPER_ATIV,
		P_CO_SITU_PROC,
		P_TP_ACAO_OPER_ATIV,
		P_TP_JUST_SITU_PROC,
		P_TX_CNTD_ANTE_ACAO,
		P_CO_USUA_ATLZ,
		P_CO_ETCA_USUA_ATLZ
  );
END SPI_TB_HIST_SITU_ACAO_OPER;


PROCEDURE SPU_TB_HIST_SITU_ACAO_OPER(
	P_NU_SEQU_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
	P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE,
  P_TP_JUST_SITU_PROC IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.TP_JUST_SITU_PROC%TYPE
) IS 
BEGIN 
 
 	UPDATE A8.TB_HIST_SITU_ACAO_OPER_ATIV 
     SET TP_JUST_SITU_PROC      = P_TP_JUST_SITU_PROC
   WHERE NU_SEQU_OPER_ATIV      = P_NU_SEQU_OPER_ATIV
	 	 AND CO_SITU_PROC           = P_CO_SITU_PROC
 		 AND DH_SITU_ACAO_OPER_ATIV = (SELECT MAX(DH_SITU_ACAO_OPER_ATIV)
                                  FROM A8.TB_HIST_SITU_ACAO_OPER_ATIV
                                  WHERE NU_SEQU_OPER_ATIV  = P_NU_SEQU_OPER_ATIV 
                                    AND CO_SITU_PROC       = P_CO_SITU_PROC)
  ;

END SPU_TB_HIST_SITU_ACAO_OPER;



/********************************************************************************************************
Nome Lógico     :	SPE_TB_HIST_SITU_ACAO_OPER_ATIV
Descrição       :	Procedure de exclusao de registros da tabela TB_HIST_SITU_ACAO_OPER_ATIV 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:13 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPE_TB_HIST_SITU_ACAO_OPER(
   P_NU_SEQU_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
   P_DH_SITU_ACAO_OPER_ATIV	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.DH_SITU_ACAO_OPER_ATIV%TYPE,
   P_CO_SITU_PROC	IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE
) IS
BEGIN

	DELETE FROM
		A8.TB_HIST_SITU_ACAO_OPER_ATIV
	WHERE
		NU_SEQU_OPER_ATIV = p_NU_SEQU_OPER_ATIV
		AND DH_SITU_ACAO_OPER_ATIV = p_DH_SITU_ACAO_OPER_ATIV
		AND CO_SITU_PROC = p_CO_SITU_PROC;
  
END SPE_TB_HIST_SITU_ACAO_OPER;

/********************************************************************************************************
Nome Lógico     :	SPS_TB_HIST_SITU_ACAO_OPER_ATIV
Descrição       :	Seleciona os dados da tabela TB_HIST_SITU_ACAO_OPER_ATIV 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:13 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPS_TB_HIST_SITU_ACAO_OPER(
  p_curselect OUT tp_cursor
) IS        

BEGIN 
 
	  -- ***** Seleciona os Dados ***** 
  OPEN p_curselect FOR 
  SELECT  
		a.NU_SEQU_OPER_ATIV,
		a.DH_SITU_ACAO_OPER_ATIV,
		a.CO_SITU_PROC,
		a.TP_ACAO_OPER_ATIV,
		a.TP_JUST_SITU_PROC,
		a.TX_CNTD_ANTE_ACAO,
		a.CO_USUA_ATLZ,
		a.CO_ETCA_USUA_ATLZ
  	FROM  
  		A8.TB_HIST_SITU_ACAO_OPER_ATIV a;
END SPS_TB_HIST_SITU_ACAO_OPER; 

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
PROCEDURE SPS_TB_HIST_MAX(
  P_NU_SEQU_OPER_ATIV       IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.NU_SEQU_OPER_ATIV%TYPE,
  P_CO_SITU_PROC            IN A8.TB_HIST_SITU_ACAO_OPER_ATIV.CO_SITU_PROC%TYPE,
  P_DH_SITU_ACAO_OPER_ATIV   OUT A8.TB_HIST_SITU_ACAO_OPER_ATIV.DH_SITU_ACAO_OPER_ATIV%TYPE

) IS        

BEGIN 
 
  SELECT  
		MAX(a.DH_SITU_ACAO_OPER_ATIV) INTO P_DH_SITU_ACAO_OPER_ATIV
  FROM  A8.TB_HIST_SITU_ACAO_OPER_ATIV a
    WHERE NU_SEQU_OPER_ATIV = P_NU_SEQU_OPER_ATIV
      AND CO_SITU_PROC      = P_CO_SITU_PROC;  
    
END SPS_TB_HIST_MAX; 




END PKG_A8_TB_HIST_SITU_ACAO_OPER;
/
