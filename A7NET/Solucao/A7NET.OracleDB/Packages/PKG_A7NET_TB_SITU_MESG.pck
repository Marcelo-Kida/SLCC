CREATE OR REPLACE PACKAGE PKG_A7NET_TB_SITU_MESG IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE SPI_TB_SITU_MESG(
   p_num_CO_MESG            IN A7.TB_SITU_MESG.CO_MESG%TYPE,
   p_num_CO_OCOR_MESG       IN A7.TB_SITU_MESG.CO_OCOR_MESG%TYPE,
   p_vch_TX_DTLH_OCOR_ERRO  IN A7.TB_SITU_MESG.TX_DTLH_OCOR_ERRO%TYPE DEFAULT NULL
);

END PKG_A7NET_TB_SITU_MESG;
/
CREATE OR REPLACE PACKAGE BODY PKG_A7NET_TB_SITU_MESG IS

/******************************************************************************************
 Objetos criados: SPI_TB_SITU_MESG
 Descricao:       Insere dados na tabela TB_SITU_MESG
 Autor:           Ivan Tabarino
 Data:            19/04/2011
******************************************************************************************/
PROCEDURE SPI_TB_SITU_MESG(
   p_num_CO_MESG            IN A7.TB_SITU_MESG.CO_MESG%TYPE,
   p_num_CO_OCOR_MESG       IN A7.TB_SITU_MESG.CO_OCOR_MESG%TYPE,
   p_vch_TX_DTLH_OCOR_ERRO  IN A7.TB_SITU_MESG.TX_DTLH_OCOR_ERRO%TYPE DEFAULT NULL
)

IS
   v_num_NU_SEQU_SITU_MESG      NUMBER;

BEGIN

   -- Insere Situação de Recebimento Bem Sucedido
   SELECT NVL(MAX(NU_SEQU_SITU_MESG), 0) + 1
   INTO   v_num_NU_SEQU_SITU_MESG
   FROM   A7.TB_SITU_MESG
   WHERE  CO_MESG = p_num_CO_MESG;

   INSERT INTO  A7.TB_SITU_MESG (NU_SEQU_SITU_MESG,
                                 CO_MESG,
                                 CO_OCOR_MESG,
                                 DH_OCOR_MESG,
                                 TX_DTLH_OCOR_ERRO)
   VALUES                       (v_num_NU_SEQU_SITU_MESG, 
                                 p_num_CO_MESG, 
                                 1, -- Recebimento Bem Sucedido
                                 SYSDATE,
                                 p_vch_TX_DTLH_OCOR_ERRO);
  
   -- Insere Situação Enviada pelo A7
   SELECT NVL(MAX(NU_SEQU_SITU_MESG), 0) + 1
   INTO   v_num_NU_SEQU_SITU_MESG
   FROM   A7.TB_SITU_MESG
   WHERE  CO_MESG = p_num_CO_MESG;
  
   INSERT INTO  A7.TB_SITU_MESG (NU_SEQU_SITU_MESG,
                                 CO_MESG,
                                 CO_OCOR_MESG,
                                 DH_OCOR_MESG,
                                 TX_DTLH_OCOR_ERRO)
   VALUES                       (v_num_NU_SEQU_SITU_MESG, 
                                 p_num_CO_MESG, 
                                 p_num_CO_OCOR_MESG,
                                 SYSDATE,
                                 SUBSTR(p_vch_TX_DTLH_OCOR_ERRO,1,500));
                                                                
                                                                
EXCEPTION
   WHEN dup_val_on_index THEN
      raise_application_error(-20000, 'Chave única já existe.', TRUE);
   WHEN OTHERS THEN
      raise_application_error(-20999, 'Erro SPI_TB_SITU_MESG: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
END SPI_TB_SITU_MESG;
  
END PKG_A7NET_TB_SITU_MESG;
/
