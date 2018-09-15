CREATE OR REPLACE PACKAGE PKG_A8_TB_HIST_ENVI_INFO_GEST IS

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
PROCEDURE SPI_TB_HIST_ENVI_INFO_GEST(
   P_NU_SEQU_OPER_ATIV       IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.NU_SEQU_OPER_ATIV%TYPE,
   P_DH_ENVI_GEST_CAIX       IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.DH_ENVI_GEST_CAIX%TYPE,
   P_CO_SITU_MOVI_GEST_CAIX  IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_SITU_MOVI_GEST_CAIX%TYPE,
   P_CO_TEXT_XML             IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_TEXT_XML%TYPE
);

END PKG_A8_TB_HIST_ENVI_INFO_GEST;
/
CREATE OR REPLACE PACKAGE BODY PKG_A8_TB_HIST_ENVI_INFO_GEST IS

-- Private type declarations
--type <TypeName> is <Datatype>;
-- Private constant declarations
--<ConstantName> constant <Datatype> := <Value>;
-- Private variable declarations
--<VariableName> <Datatype>;

-- Function and procedure implementations

/********************************************************************************************************
Nome Lógico     :	SPI_TB_HIST_ENVI_INFO_GEST
Descrição       :	Procedure de inclusão de registros na tabela TB_HIST_ENVI_INFO_GEST_CAIX
Retorno         :	-
Autor           :	Ivan Tabarino
Data Criação    :	08/02/2012 18:59:13 
Comentario      :	-
**********************************************************************************************************/
PROCEDURE SPI_TB_HIST_ENVI_INFO_GEST(
   P_NU_SEQU_OPER_ATIV       IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.NU_SEQU_OPER_ATIV%TYPE,
   P_DH_ENVI_GEST_CAIX       IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.DH_ENVI_GEST_CAIX%TYPE,
   P_CO_SITU_MOVI_GEST_CAIX  IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_SITU_MOVI_GEST_CAIX%TYPE,
   P_CO_TEXT_XML             IN  A8.TB_HIST_ENVI_INFO_GEST_CAIX.CO_TEXT_XML%TYPE
)
IS

BEGIN

   INSERT INTO A8.TB_HIST_ENVI_INFO_GEST_CAIX(NU_SEQU_OPER_ATIV,
                                              DH_ENVI_GEST_CAIX,
                                              CO_SITU_MOVI_GEST_CAIX,
                                              CO_TEXT_XML)
   VALUES                                    (P_NU_SEQU_OPER_ATIV,
                                              P_DH_ENVI_GEST_CAIX,
                                              P_CO_SITU_MOVI_GEST_CAIX,
                                              P_CO_TEXT_XML);


END SPI_TB_HIST_ENVI_INFO_GEST;

END PKG_A8_TB_HIST_ENVI_INFO_GEST;
/
