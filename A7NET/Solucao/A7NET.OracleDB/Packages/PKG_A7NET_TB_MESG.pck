CREATE OR REPLACE PACKAGE PKG_A7NET_TB_MESG IS

TYPE tp_cursor IS REF CURSOR;

PROCEDURE SPS_TB_MESG(
   p_vch_CO_CMPO_ATRB_IDEF_MESG   IN  A7.TB_MESG.CO_CMPO_ATRB_IDEF_MESG%TYPE,
   p_cur_out_CURSOR              OUT  tp_cursor
);

PROCEDURE SPI_TB_MESG(
   p_num_out_CO_MESG_OUT        OUT A7.TB_MESG.CO_MESG%TYPE,
   p_chr_SG_SIST_ORIG           IN  A7.TB_MESG.SG_SIST_ORIG%TYPE,
   p_vch_CO_MESG_MQSE           IN  A7.TB_MESG.CO_MESG_MQSE%TYPE DEFAULT NULL,
   p_vch_TP_MESG                IN  A7.TB_MESG.TP_MESG%TYPE,
   p_num_CO_EMPR_ORIG           IN  A7.TB_MESG.CO_EMPR_ORIG%TYPE,
   p_dat_DH_INIC_VIGE_REGR_TRAP IN  A7.TB_MESG.DH_INIC_VIGE_REGR_TRAP%TYPE,
   p_vch_CO_CMPO_ATRB_IDEF_MESG IN  A7.TB_MESG.CO_CMPO_ATRB_IDEF_MESG%TYPE DEFAULT NULL,
   p_num_CO_TEXT_XML_ENTR       IN  A7.TB_MESG.CO_TEXT_XML_ENTR%TYPE,
   p_num_CO_TEXT_XML_SAID       IN  A7.TB_MESG.CO_TEXT_XML_SAID%TYPE DEFAULT NULL,
   p_num_TP_FORM_MESG_SAID      IN  A7.TB_MESG.TP_FORM_MESG_SAID%TYPE
);

END PKG_A7NET_TB_MESG;
/
CREATE OR REPLACE PACKAGE BODY PKG_A7NET_TB_MESG IS

/******************************************************************************************
 Objetos criados: SPS_TB_MESG
 Descricao:       Busca dados da tabela TB_MESG
 Autor:           Ivan Tabarino
 Data:            14/04/2011
******************************************************************************************/
PROCEDURE SPS_TB_MESG(
   p_vch_CO_CMPO_ATRB_IDEF_MESG   IN  A7.TB_MESG.CO_CMPO_ATRB_IDEF_MESG%TYPE,
   p_cur_out_CURSOR              OUT  tp_cursor
)

IS

BEGIN

   OPEN   p_cur_out_CURSOR FOR
   SELECT MESG.CO_MESG,
          MESG.CO_CMPO_ATRB_IDEF_MESG,
          MESG.CO_TEXT_XML_ENTR,
          OCOR.DE_ABRV_OCOR_MESG,
          OCOR.DE_OCOR_MESG,
          SITU.NU_SEQU_SITU_MESG
   FROM   A7.TB_MESG       MESG,
          A7.TB_OCOR_MESG  OCOR,
          A7.TB_SITU_MESG  SITU
   WHERE  SITU.CO_MESG                = MESG.CO_MESG
   AND    SITU.CO_OCOR_MESG           = OCOR.CO_OCOR_MESG
   AND    MESG.CO_CMPO_ATRB_IDEF_MESG = p_vch_CO_CMPO_ATRB_IDEF_MESG
   AND    SITU.NU_SEQU_SITU_MESG      = (SELECT MAX(SITU2.NU_SEQU_SITU_MESG)
                                         FROM   A7.TB_MESG      MESG2,
                                                A7.TB_SITU_MESG SITU2
                                         WHERE  MESG2.CO_MESG = SITU2.CO_MESG
                                         AND    MESG2.CO_CMPO_ATRB_IDEF_MESG = p_vch_CO_CMPO_ATRB_IDEF_MESG);


END SPS_TB_MESG;

/******************************************************************************************
 Objetos criados: SPI_TB_MESG
 Descricao:       Insere dados na tabela TB_MESG
 Autor:           Ivan Tabarino
 Data:            19/04/2011
******************************************************************************************/
PROCEDURE SPI_TB_MESG(
   p_num_out_CO_MESG_OUT        OUT A7.TB_MESG.CO_MESG%TYPE,
   p_chr_SG_SIST_ORIG           IN  A7.TB_MESG.SG_SIST_ORIG%TYPE,
   p_vch_CO_MESG_MQSE           IN  A7.TB_MESG.CO_MESG_MQSE%TYPE DEFAULT NULL,
   p_vch_TP_MESG                IN  A7.TB_MESG.TP_MESG%TYPE,
   p_num_CO_EMPR_ORIG           IN  A7.TB_MESG.CO_EMPR_ORIG%TYPE,
   p_dat_DH_INIC_VIGE_REGR_TRAP IN  A7.TB_MESG.DH_INIC_VIGE_REGR_TRAP%TYPE,
   p_vch_CO_CMPO_ATRB_IDEF_MESG IN  A7.TB_MESG.CO_CMPO_ATRB_IDEF_MESG%TYPE DEFAULT NULL,
   p_num_CO_TEXT_XML_ENTR       IN  A7.TB_MESG.CO_TEXT_XML_ENTR%TYPE,
   p_num_CO_TEXT_XML_SAID       IN  A7.TB_MESG.CO_TEXT_XML_SAID%TYPE DEFAULT NULL,
   p_num_TP_FORM_MESG_SAID      IN  A7.TB_MESG.TP_FORM_MESG_SAID%TYPE
)

IS

BEGIN

   SELECT A7.SQ_A7_CO_MESG.NEXTVAL
   INTO   p_num_out_CO_MESG_OUT
   FROM   DUAL;

   INSERT INTO A7.TB_MESG (CO_MESG, 
                           SG_SIST_ORIG, 
                           CO_MESG_MQSE, 
                           TP_MESG, 
                           CO_EMPR_ORIG, 
                           DH_INIC_VIGE_REGR_TRAP, 
                           DH_MESG, 
                           CO_CMPO_ATRB_IDEF_MESG, 
                           CO_TEXT_XML_ENTR,
                           CO_TEXT_XML_SAID,
                           TP_FORM_MESG_SAID)
   VALUES                 (p_num_out_CO_MESG_OUT,
                           p_chr_SG_SIST_ORIG, 
                           p_vch_CO_MESG_MQSE, 
                           p_vch_TP_MESG, 
                           p_num_CO_EMPR_ORIG, 
                           p_dat_DH_INIC_VIGE_REGR_TRAP, 
                           SYSDATE, 
                           p_vch_CO_CMPO_ATRB_IDEF_MESG, 
                           p_num_CO_TEXT_XML_ENTR, 
                           p_num_CO_TEXT_XML_SAID, 
                           p_num_TP_FORM_MESG_SAID);
                           

EXCEPTION
   WHEN dup_val_on_index THEN
      raise_application_error(-20000, 'Chave única já existe.', TRUE);
   WHEN OTHERS THEN
      raise_application_error(-20999, 'Erro SPI_TB_MESG: [' || SQLCODE || '-' || SQLERRM || ']', FALSE);
END SPI_TB_MESG;
  
END PKG_A7NET_TB_MESG;
/
