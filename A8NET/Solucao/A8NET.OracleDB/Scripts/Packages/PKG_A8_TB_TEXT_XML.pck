CREATE OR REPLACE PACKAGE PKG_A8_TB_TEXT_XML IS

-- Author  : MAPS
-- Created : 08/03/2011 15:59:13
-- Purpose : Realizar operações basicas de Insert, Update, Delete e Select na tabela TB_TEXT_XML

-- Public type declarations
TYPE tp_cursor IS REF CURSOR;

-- Public constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Public variable declarations
--<VariableName> <Datatype>;

-- Public function and procedure declarations
PROCEDURE SPI_TB_TEXT_XML
(
  P_CO_TEXT_XML           IN OUT A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
  P_NU_SEQU_TEXT_XML      IN  A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE,
  P_TX_XML                IN  A8.TB_TEXT_XML.TX_XML%TYPE
);

PROCEDURE SPS_TB_TEXT_XML(
	P_CO_TEXT_XML	IN A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
  p_cur_out_CURSOR OUT tp_cursor
);

PROCEDURE SPU_TB_TEXT_XML(
	P_CO_TEXT_XML	IN A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
	P_NU_SEQU_TEXT_XML	IN A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE,
	P_TX_XML	IN A8.TB_TEXT_XML.TX_XML%TYPE
);

PROCEDURE SPE_TB_TEXT_XML(
   P_CO_TEXT_XML	IN A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
   P_NU_SEQU_TEXT_XML	IN A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE
);

END PKG_A8_TB_TEXT_XML;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_TB_TEXT_XML IS

-- Private type declarations
--type <TypeName> is <Datatype>;
-- Private constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Private variable declarations
--<VariableName> <Datatype>;

-- Function and procedure implementations

/********************************************************************************************************
Nome Lógico     :	SPI_TB_TEXT_XML
Descrição       :	Procedure de inclusão de registros na tabela TB_TEXT_XML 
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
PROCEDURE SPI_TB_TEXT_XML
(
  P_CO_TEXT_XML           IN OUT A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
  P_NU_SEQU_TEXT_XML      IN  A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE,
  P_TX_XML                IN  A8.TB_TEXT_XML.TX_XML%TYPE
)

IS

BEGIN

     IF P_CO_TEXT_XML IS NULL OR P_CO_TEXT_XML = 0  THEN
          SELECT A8.SQ_A8_CO_TEXT_XML.NEXTVAL INTO P_CO_TEXT_XML FROM DUAL;
     END IF;
         
     INSERT INTO A8.TB_TEXT_XML (
            CO_TEXT_XML,
            NU_SEQU_TEXT_XML,
            TX_XML
            )
     VALUES (
            A8.SQ_A8_CO_TEXT_XML.CURRVAL,
            P_NU_SEQU_TEXT_XML,
            P_TX_XML
     );

EXCEPTION
   WHEN dup_val_on_index THEN
      raise_application_error(-20000, 'Chave única já existe.', TRUE);
   WHEN OTHERS THEN
      raise_application_error(-20999, 'Erro SPI_TB_TEXT_XML: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
END SPI_TB_TEXT_XML;


/********************************************************************************************************
Nome Lógico     :	SPS_TB_TEXT_XML
Descrição       :	Seleciona os dados da tabela TB_TEXT_XML 
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
PROCEDURE SPS_TB_TEXT_XML(
	P_CO_TEXT_XML	IN A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
  p_cur_out_CURSOR OUT tp_cursor
) IS        

BEGIN 
 
	  -- ***** Seleciona os Dados ***** 
  OPEN p_cur_out_CURSOR FOR 
  SELECT  
		a.CO_TEXT_XML,
		a.NU_SEQU_TEXT_XML,
		a.TX_XML
  FROM  
  		A8.TB_TEXT_XML a
  WHERE a.co_text_xml = P_CO_TEXT_XML
  ORDER BY a.nu_sequ_text_xml;

END SPS_TB_TEXT_XML; 


/********************************************************************************************************
Nome Lógico     :	SPU_TB_TEXT_XML
Descrição       :	Procedure de atualização de registros na tabela TB_TEXT_XML 
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
PROCEDURE SPU_TB_TEXT_XML(
	P_CO_TEXT_XML	IN A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
	P_NU_SEQU_TEXT_XML	IN A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE,
	P_TX_XML	IN A8.TB_TEXT_XML.TX_XML%TYPE
) IS 
BEGIN 
 
	UPDATE A8.TB_TEXT_XML SET
		TX_XML = P_TX_XML
	WHERE 
		CO_TEXT_XML = P_CO_TEXT_XML
		AND NU_SEQU_TEXT_XML = P_NU_SEQU_TEXT_XML
  ;
END SPU_TB_TEXT_XML; 


/********************************************************************************************************
Nome Lógico     :	SPE_TB_TEXT_XML
Descrição       :	Procedure de exclusao de registros da tabela TB_TEXT_XML 
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
PROCEDURE SPE_TB_TEXT_XML(
   P_CO_TEXT_XML	IN A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
   P_NU_SEQU_TEXT_XML	IN A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE
) IS
BEGIN

	DELETE FROM
		A8.TB_TEXT_XML
	WHERE
		CO_TEXT_XML = p_CO_TEXT_XML
		AND NU_SEQU_TEXT_XML = p_NU_SEQU_TEXT_XML;
  
END SPE_TB_TEXT_XML;


END PKG_A8_TB_TEXT_XML;
/
