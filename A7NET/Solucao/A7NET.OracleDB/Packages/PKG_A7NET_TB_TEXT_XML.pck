CREATE OR REPLACE PACKAGE PKG_A7NET_TB_TEXT_XML IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE SPS_TB_TEXT_XML(
   p_num_CO_TEXT_XML   IN  A7.TB_TEXT_XML.CO_TEXT_XML%TYPE,
   p_cur_out_CURSOR   OUT  tp_cursor
);

PROCEDURE SPS_TB_TEXT_XML_A8(
   p_cur_out_CURSOR   OUT  tp_cursor
);

PROCEDURE SPI_TB_TEXT_XML(
  p_num_CO_TEXT_XML           IN OUT A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
  p_num_NU_SEQU_TEXT_XML      IN     A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE,
  p_vch_TX_XML                IN     A8.TB_TEXT_XML.TX_XML%TYPE
);

END PKG_A7NET_TB_TEXT_XML;
/
CREATE OR REPLACE PACKAGE BODY PKG_A7NET_TB_TEXT_XML IS

/******************************************************************************************
 Objetos criados: SPS_TB_TEXT_XML
 Descricao:       Busca dados da tabela TB_TEXT_XML
 Autor:           Ivan Tabarino
 Data:            04/04/2011
******************************************************************************************/
PROCEDURE SPS_TB_TEXT_XML(
   p_num_CO_TEXT_XML   IN  A7.TB_TEXT_XML.CO_TEXT_XML%TYPE,
   p_cur_out_CURSOR   OUT  tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT TX_XML
   FROM   A7.TB_TEXT_XML
   WHERE  CO_TEXT_XML = p_num_CO_TEXT_XML
   ORDER  BY NU_SEQU_TEXT_XML;

   
END SPS_TB_TEXT_XML;

/******************************************************************************************
 Objetos criados: SPS_TB_TEXT_XML_A8
 Descricao:       Busca dados da tabela A8.TB_TEXT_XML
 Autor:           Ivan Tabarino
 Data:            14/04/2011
******************************************************************************************/
PROCEDURE SPS_TB_TEXT_XML_A8(
   p_cur_out_CURSOR   OUT  tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT TX_XML   
   FROM   A8.TB_TEXT_XML 
   WHERE  CO_TEXT_XML = 0
   ORDER  BY NU_SEQU_TEXT_XML;

   
END SPS_TB_TEXT_XML_A8;

/******************************************************************************************
 Objetos criados: SPI_TB_TEXT_XML
 Descricao:       Insere dados na tabela A7.TB_TEXT_XML
 Autor:           Ivan Tabarino
 Data:            18/04/2011
******************************************************************************************/
PROCEDURE SPI_TB_TEXT_XML(
  p_num_CO_TEXT_XML           IN OUT A8.TB_TEXT_XML.CO_TEXT_XML%TYPE,
  p_num_NU_SEQU_TEXT_XML      IN     A8.TB_TEXT_XML.NU_SEQU_TEXT_XML%TYPE,
  p_vch_TX_XML                IN     A8.TB_TEXT_XML.TX_XML%TYPE
)

IS

BEGIN

   IF p_num_CO_TEXT_XML IS NULL 
   OR p_num_CO_TEXT_XML  = 0 THEN
      SELECT A7.SQ_A7_CO_TEXT_XML.NEXTVAL 
      INTO   p_num_CO_TEXT_XML 
      FROM   DUAL;
   END IF;
         
   INSERT INTO A7.TB_TEXT_XML(CO_TEXT_XML,
                              NU_SEQU_TEXT_XML,
                              TX_XML)
   VALUES                    (p_num_CO_TEXT_XML,
                              p_num_NU_SEQU_TEXT_XML,
                              p_vch_TX_XML);
                               

EXCEPTION
   WHEN dup_val_on_index THEN
      raise_application_error(-20000, 'Chave única já existe.', TRUE);
   WHEN OTHERS THEN
      raise_application_error(-20999, 'Erro SPI_TB_TEXT_XML: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
END SPI_TB_TEXT_XML;
  
END PKG_A7NET_TB_TEXT_XML;
/
