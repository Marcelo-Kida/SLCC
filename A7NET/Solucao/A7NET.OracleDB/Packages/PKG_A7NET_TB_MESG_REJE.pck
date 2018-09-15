CREATE OR REPLACE PACKAGE PKG_A7NET_TB_MESG_REJE IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE SPI_TB_MESG_REJE(
   p_vch_CO_MESG_MQSE           IN A7.TB_MESG_REJE.CO_MESG_MQSE%TYPE,
   p_num_CO_OCOR_MESG           IN A7.TB_MESG_REJE.CO_OCOR_MESG%TYPE,
   p_num_CO_TEXT_XML            IN A7.TB_MESG_REJE.CO_TEXT_XML%TYPE,
   p_vch_NO_ARQU_ENTR_FILA_MQSE IN A7.TB_MESG_REJE.NO_ARQU_ENTR_FILA_MQSE%TYPE,
   p_vch_TX_DTLH_OCOR_ERRO      IN A7.TB_MESG_REJE.TX_DTLH_OCOR_ERRO%TYPE DEFAULT NULL
);

END PKG_A7NET_TB_MESG_REJE;
/
CREATE OR REPLACE PACKAGE BODY PKG_A7NET_TB_MESG_REJE IS

/******************************************************************************************
 Objetos criados: SPI_TB_MESG_REJE
 Descricao:       Insere dados na tabela TB_MESG_REJE
 Autor:           Ivan Tabarino
 Data:            20/04/2011
******************************************************************************************/
PROCEDURE SPI_TB_MESG_REJE(
   p_vch_CO_MESG_MQSE           IN A7.TB_MESG_REJE.CO_MESG_MQSE%TYPE,
   p_num_CO_OCOR_MESG           IN A7.TB_MESG_REJE.CO_OCOR_MESG%TYPE,
   p_num_CO_TEXT_XML            IN A7.TB_MESG_REJE.CO_TEXT_XML%TYPE,
   p_vch_NO_ARQU_ENTR_FILA_MQSE IN A7.TB_MESG_REJE.NO_ARQU_ENTR_FILA_MQSE%TYPE,
   p_vch_TX_DTLH_OCOR_ERRO      IN A7.TB_MESG_REJE.TX_DTLH_OCOR_ERRO%TYPE DEFAULT NULL
)

IS

BEGIN

   INSERT INTO A7.TB_MESG_REJE(CO_MESG_MQSE,
                               CO_OCOR_MESG,
                               DH_RECB_MESG,
                               CO_TEXT_XML,
                               DH_ENTR_FILA_MQSE,
                               NO_ARQU_ENTR_FILA_MQSE,
                               TX_DTLH_OCOR_ERRO) 
   VALUES                     (p_vch_CO_MESG_MQSE,
                               p_num_CO_OCOR_MESG,
                               SYSDATE,
                               p_num_CO_TEXT_XML,
                               SYSDATE,
                               p_vch_NO_ARQU_ENTR_FILA_MQSE,
                               p_vch_TX_DTLH_OCOR_ERRO);
                           

EXCEPTION
   WHEN dup_val_on_index THEN
      raise_application_error(-20000, 'Chave única já existe.', TRUE);
   WHEN OTHERS THEN
      raise_application_error(-20999, 'Erro SPI_TB_MESG_REJE: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
END SPI_TB_MESG_REJE;
  
END PKG_A7NET_TB_MESG_REJE;
/
