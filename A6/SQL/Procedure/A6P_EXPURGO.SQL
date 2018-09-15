CREATE OR REPLACE PACKAGE A6K_EXPURGO AS
--
--***************
--	Procedures
--***************
--
PROCEDURE gravarhistoricoexecucao
(	protina		IN		NUMBER
,	pstatus		IN		NUMBER
,	perro		IN		VARCHAR2
);
--
PROCEDURE obter_data_corte;
--
PROCEDURE PROCESSAR;
--
END;
/
CREATE OR REPLACE PACKAGE BODY a6k_expurgo
IS
   data_corte      DATE;
   qt_processada   NUMBER;

   PROCEDURE gravarhistoricoexecucao (
      protina   IN   NUMBER,
      pstatus   IN   NUMBER,
      perro     IN   VARCHAR2
   )
   IS
   BEGIN
      INSERT INTO a7.tb_hist_exec_roti_batch (
                     co_roti_batch,
                     dh_fim_exec,
                     in_exec_suce,
                     de_erro_exec
                  )
           VALUES (protina, SYSDATE, pstatus, SUBSTR (perro, 1, 200));
   EXCEPTION
      WHEN OTHERS
      THEN
         raise_application_error (
            SQLCODE,
            'ERRO NA ROTINA DE GRAVACAO DE HISTORICO DE EXECUCAO: ' || SQLERRM
         );
   END gravarhistoricoexecucao;

   PROCEDURE obter_data_corte
   IS
      vdatacorte   DATE;
      vdiascorte   NUMBER;
   BEGIN
      vdiascorte := a8proc.a8k_expurgo.obterqtddiasexpurgo () - 1;
      SELECT TRUNC ((SYSDATE - vdiascorte))
        INTO data_corte
        FROM dual;
   END obter_data_corte;

   PROCEDURE msg (mensagem IN VARCHAR2)
   IS
   BEGIN
      DBMS_OUTPUT.put_line (
         TO_CHAR (SYSDATE, 'yyyy-mm-dd hh24:mi:ss') || ': ' || mensagem
      );
   END;

   PROCEDURE processar
   IS
   BEGIN
      -- Chamada da procedure para exclusividade de área de Rollback
      SYS.dbP_muda_rbs('RBSBIG','ON');
      
      -- Otem a data de corte para o expurgo
      obter_data_corte;
      msg ('Expurgo da tabela A6.TB_CTRL_PROC_REME.');
      -- Expurgo da tabela A6.TB_CTRL_PROC_REME
      INSERT INTO a6hist.tb_ctrl_proc_reme
                  (dt_reme,
                   tp_mesg,
                   sg_sist,
                   co_veic_lega,
                   qt_regt_info_sist_orig,
                   qt_regt_reje,
                   qt_regt_recb,
                   dh_inic_reme_orig,
                   dh_fim_reme_orig,
                   dh_inic_proc_reme,
                   dh_fim_proc_reme,
                   co_reme_prmr,
                   co_reme_ulti,
                   va_totl_reme,
                   co_situ_reme,
                   dh_ulti_atlz
                  )
         SELECT dt_reme,
                tp_mesg,
                sg_sist,
                co_veic_lega,
                qt_regt_info_sist_orig,
                qt_regt_reje,
                qt_regt_recb,
                dh_inic_reme_orig,
                dh_fim_reme_orig,
                dh_inic_proc_reme,
                dh_fim_proc_reme,
                co_reme_prmr,
                co_reme_ulti,
                va_totl_reme,
                co_situ_reme,
                dh_ulti_atlz
           FROM a6.tb_ctrl_proc_reme
          WHERE dt_reme < data_corte;

      DELETE
        FROM a6.tb_ctrl_proc_reme
       WHERE dt_reme < data_corte;
       
      COMMIT;
      SYS.dbP_muda_rbs('RBSBIG','ON');
      
      msg ('Expurgo da tabela A6.TB_MOVI_CAIX_SUB_RESE.');
-- Expurgo da tabela A6.TB_MOVI_CAIX_SUB_RESE
      INSERT INTO a6hist.tb_movi_caix_sub_rese
                  (co_movi_caix_sub_rese,
                   co_empr,
                   tp_oper,
                   dh_movi_caix_sub_rese,
                   sg_sist,
                   co_loca_liqu,
                   co_veic_lega,
                   co_item_caix,
                   de_tipo_liqu,
                   va_movi_caix_sub_rese,
                   in_movi_entr_said,
                   co_cnpj_cnpt,
                   no_cnpt,
                   de_ativ,
                   nu_sequ_oper_ativ,
                   co_situ_movi_caix_sub_rese,
                   tp_gera_movi,
                   tx_moti_gera_movi,
                   co_item_caix_nive_01,
                   co_item_caix_nive_02,
                   co_item_caix_nive_03,
                   co_item_caix_nive_04,
                   co_item_caix_nive_05,
                   co_usua_ulti_atlz,
                   co_etca_trab_ulti_atlz,
                   dh_ulti_atlz,
		   dt_retn_oper
                  )
         SELECT co_movi_caix_sub_rese,
                co_empr,
                tp_oper,
                dh_movi_caix_sub_rese,
                sg_sist,
                co_loca_liqu,
                co_veic_lega,
                co_item_caix,
                de_tipo_liqu,
                va_movi_caix_sub_rese,
                in_movi_entr_said,
                co_cnpj_cnpt,
                no_cnpt,
                de_ativ,
                nu_sequ_oper_ativ,
                co_situ_movi_caix_sub_rese,
                tp_gera_movi,
                tx_moti_gera_movi,
                co_item_caix_nive_01,
                co_item_caix_nive_02,
                co_item_caix_nive_03,
                co_item_caix_nive_04,
                co_item_caix_nive_05,
                co_usua_ulti_atlz,
                co_etca_trab_ulti_atlz,
                dh_ulti_atlz,
		dt_retn_oper
           FROM a6.tb_movi_caix_sub_rese
          WHERE dh_movi_caix_sub_rese < data_corte;
      DELETE
        FROM a6.tb_movi_caix_sub_rese
       WHERE dh_movi_caix_sub_rese < data_corte;
       
      COMMIT;
      SYS.dbP_muda_rbs('RBSBIG','ON');

      msg ('Expurgo da tabela A6.TB_MOVI_FUTU_OPER_ATIV.');
-- Expurgo da tabela A6.TB_MOVI_FUTU_OPER_ATIV
      INSERT INTO a6hist.tb_movi_futu_oper_ativ
                  (dh_reme,
                   sg_sist,
                   co_reme_movi_futu,
                   co_empr,
                   dt_fech_proc,
                   co_loca_liqu,
                   de_tipo_liqu,
                   dt_liqu_oper,
                   in_movi_entr_said,
                   va_liqu_oper,
                   de_ativ,
                   co_veic_lega,
                   co_item_caix,
                   co_cnpj_cnpt,
                   no_cnpt,
                   tp_gera_movi,
                   tx_moti_gera_movi,
                   co_item_caix_nive_01,
                   co_item_caix_nive_02,
                   co_item_caix_nive_03,
                   co_item_caix_nive_04,
                   co_item_caix_nive_05,
                   co_usua_ulti_atlz,
                   co_etca_trab_ulti_atlz,
                   dh_ulti_atlz,
                   co_prod,
                   dt_retn_oper
                  )
         SELECT dh_reme,
                sg_sist,
                co_reme_movi_futu,
                co_empr,
                dt_fech_proc,
                co_loca_liqu,
                de_tipo_liqu,
                dt_liqu_oper,
                in_movi_entr_said,
                va_liqu_oper,
                de_ativ,
                co_veic_lega,
                co_item_caix,
                co_cnpj_cnpt,
                no_cnpt,
                tp_gera_movi,
                tx_moti_gera_movi,
                co_item_caix_nive_01,
                co_item_caix_nive_02,
                co_item_caix_nive_03,
                co_item_caix_nive_04,
                co_item_caix_nive_05,
                co_usua_ulti_atlz,
                co_etca_trab_ulti_atlz,
                dh_ulti_atlz,
                co_prod,
                dt_retn_oper
           FROM a6.tb_movi_futu_oper_ativ
          WHERE dt_liqu_oper < data_corte;
      DELETE
        FROM a6.tb_movi_futu_oper_ativ
       WHERE dt_liqu_oper < data_corte;
       
      COMMIT;
      SYS.dbP_muda_rbs('RBSBIG','ON');

      msg ('Expurgo da tabela A6.TB_SALD_FECH_VEIC_LEGA.');
-- Expurgo da tabela A6.TB_SALD_FECH_VEIC_LEGA
      INSERT INTO a6hist.tb_sald_fech_veic_lega
                  (dt_apur_sald,
                   co_veic_lega,
                   sg_sist,
                   dh_envi_sald,
                   va_sald_fech,
                   dh_ulti_atlz
                  )
         SELECT dt_apur_sald,
                co_veic_lega,
                sg_sist,
                dh_envi_sald,
                va_sald_fech,
                dh_ulti_atlz
           FROM a6.tb_sald_fech_veic_lega
          WHERE dt_apur_sald < data_corte;
      DELETE
        FROM a6.tb_sald_fech_veic_lega
       WHERE dt_apur_sald < data_corte;
       
      COMMIT;
      SYS.dbP_muda_rbs('RBSBIG','ON');
      
      INSERT INTO a6hist.tb_posi_caix_sub_rese
                  (dt_caix_sub_rese,
                   co_veic_lega,
                   sg_sist,
                   co_situ_caix_sub_rese,
                   va_utlz_aber_caix_sub_rese,
                   va_aber_rece_caix_sub_rese,
                   va_aber_info_caix_sub_rese,
                   va_movi_caix_sub_rese,
                   va_fech_caix_sub_rese,
                   co_usua_ulti_atlz,
                   co_etca_trab_ulti_atlz,
                   dh_ulti_atlz
                  )
         SELECT dt_caix_sub_rese,
                co_veic_lega,
                sg_sist,
                co_situ_caix_sub_rese,
                va_utlz_aber_caix_sub_rese,
                va_aber_rece_caix_sub_rese,
                va_aber_info_caix_sub_rese,
                va_movi_caix_sub_rese,
                va_fech_caix_sub_rese,
                co_usua_ulti_atlz,
                co_etca_trab_ulti_atlz,
                dh_ulti_atlz
           FROM a6.tb_posi_caix_sub_rese
          WHERE dt_caix_sub_rese < data_corte;
      DELETE
        FROM a6.tb_posi_caix_sub_rese
       WHERE dt_caix_sub_rese < data_corte;
      msg ('Expurgo da tabela A6.TB_REME_REJE.');
-- Expurgo da tabela A6.TB_REME_REJE
      INSERT INTO a6hist.tb_reme_reje
                  (sg_sist_orig_info,
                   tp_mesg_inte,
                   co_empr,
                   co_text_xml_reje,
                   tx_xml_erro,
                   dh_reme_reje,
                   co_veic_lega
                  )
         SELECT sg_sist_orig_info,
                tp_mesg_inte,
                co_empr,
                co_text_xml_reje,
                tx_xml_erro,
                dh_reme_reje,
                co_veic_lega
           FROM a6.tb_reme_reje
          WHERE dh_reme_reje < data_corte;
      INSERT INTO a6hist.tb_text_xml
                  (co_text_xml, nu_sequ_text_xml, tx_xml)
         SELECT co_text_xml, nu_sequ_text_xml, tx_xml
           FROM a6.tb_text_xml
          WHERE co_text_xml IN (SELECT co_text_xml_reje
                                  FROM a6.tb_reme_reje
                                 WHERE dh_reme_reje < data_corte);
      DELETE
        FROM a6.tb_text_xml
       WHERE co_text_xml IN (SELECT co_text_xml_reje
                               FROM a6.tb_reme_reje
                              WHERE dh_reme_reje < data_corte);
      DELETE
        FROM a6.tb_reme_reje
       WHERE dh_reme_reje < data_corte;
       
      COMMIT;
      SYS.dbP_muda_rbs('RBSBIG','ON');

      msg ('Fim do expurgo do sistema A6.');
      gravarhistoricoexecucao (3, 1, NULL);
      
      COMMIT;
      
   EXCEPTION
      WHEN OTHERS
      THEN
         ROLLBACK;
         gravarhistoricoexecucao (3, 2, SUBSTR (SQLERRM, 1, 200));
         COMMIT;
         DBMS_OUTPUT.put_line (SQLCODE || ' - ' || SQLERRM);
         -- Chamada da procedure para exclusividade de área de Rollback
         SYS.dbP_muda_rbs('RBSBIG','OFF');
   END processar;
END a6k_expurgo;
/

GRANT EXECUTE ON A6K_EXPURGO TO RL01_A6_MANUTENCAO;
GRANT EXECUTE ON A6K_EXPURGO TO RL02_A6_CONSULTA;
