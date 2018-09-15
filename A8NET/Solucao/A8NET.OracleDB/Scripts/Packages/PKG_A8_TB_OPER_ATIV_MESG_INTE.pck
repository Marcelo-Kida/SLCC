CREATE OR REPLACE PACKAGE PKG_A8_TB_OPER_ATIV_MESG_INTE IS

-- Author  : MAPS
-- Created : 08/03/2011 15:59:12
-- Purpose : Realizar operações basicas de Insert, Update, Delete e Select na tabela TB_OPER_ATIV_MESG_INTE

-- Public type declarations
TYPE tp_cursor IS REF CURSOR;

-- Public constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Public variable declarations
--<VariableName> <Datatype>;

-- Public function and procedure declarations
PROCEDURE SPI_TB_OPER_ATIV_MESG_INTE(
	P_NU_SEQU_OPER_ATIV IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
	P_DH_MESG_INTE      IN A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE,
	P_TP_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_MESG_INTE%TYPE,
	P_TP_SOLI_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_SOLI_MESG_INTE%TYPE,
	P_CO_TEXT_XML	IN A8.TB_OPER_ATIV_MESG_INTE.CO_TEXT_XML%TYPE,
	P_TP_FORM_MESG_SAID	IN A8.TB_OPER_ATIV_MESG_INTE.TP_FORM_MESG_SAID%TYPE
);

PROCEDURE SPS_TB_OPER_ATIV_MESG_INTE(
  P_NU_SEQU_OPER_ATIV  IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE
 ,p_curselect          OUT tp_cursor
);

PROCEDURE SPU_TB_OPER_ATIV_MESG_INTE(
	P_NU_SEQU_OPER_ATIV	IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
	P_DH_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE,
	P_TP_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_MESG_INTE%TYPE,
	P_TP_SOLI_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_SOLI_MESG_INTE%TYPE,
	P_CO_TEXT_XML	IN A8.TB_OPER_ATIV_MESG_INTE.CO_TEXT_XML%TYPE,
	P_TP_FORM_MESG_SAID	IN A8.TB_OPER_ATIV_MESG_INTE.TP_FORM_MESG_SAID%TYPE
);

PROCEDURE SPE_TB_OPER_ATIV_MESG_INTE(
   P_NU_SEQU_OPER_ATIV	IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
   P_DH_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE
);

PROCEDURE SPS_MAX(
  P_NU_SEQU_OPER_ATIV  IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
  P_DH_MESG_INTE       OUT A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE
);

END PKG_A8_TB_OPER_ATIV_MESG_INTE;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_TB_OPER_ATIV_MESG_INTE IS

-- Private type declarations
--type <TypeName> is <Datatype>;
-- Private constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Private variable declarations
--<VariableName> <Datatype>;

-- Function and procedure implementations

/********************************************************************************************************
Nome Lógico     :	SPI_TB_OPER_ATIV_MESG_INTE
Descrição       :	Procedure de inclusão de registros na tabela TB_OPER_ATIV_MESG_INTE 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:12 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPI_TB_OPER_ATIV_MESG_INTE(
	P_NU_SEQU_OPER_ATIV IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
	P_DH_MESG_INTE      IN A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE,
	P_TP_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_MESG_INTE%TYPE,
	P_TP_SOLI_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_SOLI_MESG_INTE%TYPE,
	P_CO_TEXT_XML	IN A8.TB_OPER_ATIV_MESG_INTE.CO_TEXT_XML%TYPE,
	P_TP_FORM_MESG_SAID	IN A8.TB_OPER_ATIV_MESG_INTE.TP_FORM_MESG_SAID%TYPE
) IS
BEGIN

  INSERT INTO A8.TB_OPER_ATIV_MESG_INTE(
		NU_SEQU_OPER_ATIV,
		DH_MESG_INTE,
		TP_MESG_INTE, 
		TP_SOLI_MESG_INTE, 
		CO_TEXT_XML, 
		TP_FORM_MESG_SAID 
  )
  VALUES(
		P_NU_SEQU_OPER_ATIV,
		P_DH_MESG_INTE,
		P_TP_MESG_INTE,
		P_TP_SOLI_MESG_INTE,
		P_CO_TEXT_XML,
		P_TP_FORM_MESG_SAID
  );
END SPI_TB_OPER_ATIV_MESG_INTE;

/********************************************************************************************************
Nome Lógico     :	SPU_TB_OPER_ATIV_MESG_INTE
Descrição       :	Procedure de atualização de registros na tabela TB_OPER_ATIV_MESG_INTE 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:12 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPU_TB_OPER_ATIV_MESG_INTE(
	P_NU_SEQU_OPER_ATIV	IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
	P_DH_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE,
	P_TP_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_MESG_INTE%TYPE,
	P_TP_SOLI_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.TP_SOLI_MESG_INTE%TYPE,
	P_CO_TEXT_XML	IN A8.TB_OPER_ATIV_MESG_INTE.CO_TEXT_XML%TYPE,
	P_TP_FORM_MESG_SAID	IN A8.TB_OPER_ATIV_MESG_INTE.TP_FORM_MESG_SAID%TYPE
) IS 
BEGIN 
 
	UPDATE A8.TB_OPER_ATIV_MESG_INTE SET
		TP_MESG_INTE = P_TP_MESG_INTE,
		TP_SOLI_MESG_INTE = P_TP_SOLI_MESG_INTE,
		CO_TEXT_XML = P_CO_TEXT_XML,
		TP_FORM_MESG_SAID = P_TP_FORM_MESG_SAID
	WHERE 
		NU_SEQU_OPER_ATIV = P_NU_SEQU_OPER_ATIV
		AND DH_MESG_INTE = P_DH_MESG_INTE
  ;
END SPU_TB_OPER_ATIV_MESG_INTE; 


/********************************************************************************************************
Nome Lógico     :	SPE_TB_OPER_ATIV_MESG_INTE
Descrição       :	Procedure de exclusao de registros da tabela TB_OPER_ATIV_MESG_INTE 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:12 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPE_TB_OPER_ATIV_MESG_INTE(
   P_NU_SEQU_OPER_ATIV	IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
   P_DH_MESG_INTE	IN A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE
) IS
BEGIN

	DELETE FROM
		A8.TB_OPER_ATIV_MESG_INTE
	WHERE
		NU_SEQU_OPER_ATIV = p_NU_SEQU_OPER_ATIV
		AND DH_MESG_INTE = p_DH_MESG_INTE;
  
END SPE_TB_OPER_ATIV_MESG_INTE;

/********************************************************************************************************
Nome Lógico     :	SPS_TB_OPER_ATIV_MESG_INTE
Descrição       :	Seleciona os dados da tabela TB_OPER_ATIV_MESG_INTE 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:12 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPS_TB_OPER_ATIV_MESG_INTE(
  P_NU_SEQU_OPER_ATIV  IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE
 ,p_curselect OUT tp_cursor
) IS        

BEGIN

	  -- ***** Seleciona os Dados ***** 
  OPEN p_curselect FOR 
  SELECT  
		a.NU_SEQU_OPER_ATIV,
		a.DH_MESG_INTE,
		a.TP_MESG_INTE,
		a.TP_SOLI_MESG_INTE,
		a.CO_TEXT_XML,
		a.TP_FORM_MESG_SAID
  	FROM  
  		A8.TB_OPER_ATIV_MESG_INTE a
    WHERE 
      NU_SEQU_OPER_ATIV = P_NU_SEQU_OPER_ATIV;
      
END SPS_TB_OPER_ATIV_MESG_INTE; 

/********************************************************************************************************
Nome Lógico     :	SPE_TB_OPER_ATIV_MESG_INTE
Descrição       :	Procedure de exclusao de registros da tabela TB_OPER_ATIV_MESG_INTE 
Retorno         :	-
Autor           :	Fernando Grassi Chaves
Data Criação    :	08/03/2011 15:59:12 
Comentario      :	-
----------------------------------------------------------------------------------------------------------
Alterado        :
Data            :
Motivo          :
Solicitado Por  :
**********************************************************************************************************/
PROCEDURE SPS_MAX(
  P_NU_SEQU_OPER_ATIV  IN A8.TB_OPER_ATIV_MESG_INTE.NU_SEQU_OPER_ATIV%TYPE,
  P_DH_MESG_INTE       OUT A8.TB_OPER_ATIV_MESG_INTE.DH_MESG_INTE%TYPE
) IS        

BEGIN 
 
  SELECT   MAX(DH_MESG_INTE) INTO P_DH_MESG_INTE
  FROM     A8.TB_OPER_ATIV_MESG_INTE    
  WHERE    NU_SEQU_OPER_ATIV       =   P_NU_SEQU_OPER_ATIV;  
    
END SPS_MAX; 



END PKG_A8_TB_OPER_ATIV_MESG_INTE;
/
