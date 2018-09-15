CREATE OR REPLACE PACKAGE A8K_EXPURGO AS
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
--**************
--	Functions
--**************
--
FUNCTION obterqtddiasexpurgo RETURN NUMBER;
--
END;
/
--
CREATE OR REPLACE PACKAGE BODY a8k_expurgo
IS
	data_corte			DATE;
	qt_processada		NUMBER;
--
--***********************************************************
--
FUNCTION obterqtddiasexpurgo RETURN NUMBER
--
IS
--
	inicio		NUMBER;
	fim			NUMBER;
	valor		NUMBER;
	strxml		VARCHAR(32000);
--
	CURSOR curtextxml
	IS
		SELECT	TX_XML
		FROM	A8.TB_TEXT_XML
		WHERE	CO_TEXT_XML = 0
		AND		NU_SEQU_TEXT_XML < 5
		ORDER	BY NU_SEQU_TEXT_XML;
--
	BEGIN
		strxml := '';
		--
		FOR rowtextxml IN curtextxml
		LOOP
			strxml := strxml || rowtextxml.tx_xml;
		END LOOP;
		--
		inicio	:= INSTR (strxml,'>',INSTR (strxml,'<BASE_HISTORICA') + 1) + 1;
		fim		:= INSTR (strxml, '</BASE_HISTORICA>');
		valor	:= TO_NUMBER (SUBSTR (strxml, inicio, fim - inicio));
		--
		RETURN valor;
		--
	END obterqtddiasexpurgo;
--
--***********************************************************
--
PROCEDURE gravarhistoricoexecucao
(	protina		IN		NUMBER,
	pstatus		IN		NUMBER,
	perro		IN		VARCHAR2
)
--
IS
--
	BEGIN
	--
		INSERT	INTO A7.TB_HIST_EXEC_ROTI_BATCH
				(CO_ROTI_BATCH,
				DH_FIM_EXEC,
				IN_EXEC_SUCE,
				DE_ERRO_EXEC)
		VALUES	(protina, SYSDATE, pstatus, SUBSTR (perro, 1, 200));
		--
		EXCEPTION
			WHEN OTHERS THEN
				raise_application_error (SQLCODE,'ERRO NA ROTINA DE GRAVACAO DE HISTORICO DE EXECUCAO: ' || SQLERRM);
	--
	END gravarhistoricoexecucao;
--
--***********************************************************
--
PROCEDURE obter_data_corte
--
IS
--
	vdatacorte		DATE;
	vdiascorte		NUMBER;
	--
	BEGIN
		--
		vdiascorte := a8k_expurgo.obterqtddiasexpurgo () - 1;
		--
		SELECT	TRUNC ((SYSDATE - vdiascorte))
		INTO	data_corte
		FROM	DUAL;
	--
	END obter_data_corte;
--
--***********************************************************
--
PROCEDURE msg 
(	mensagem	IN		VARCHAR2
)
--
IS
--
	BEGIN
		DBMS_OUTPUT.put_line (TO_CHAR (SYSDATE, 'yyyy-mm-dd hh24:mi:ss') || ': ' || mensagem);
	END;
--
--***********************************************************
--
PROCEDURE processar
--
IS
--
	CURSOR curoperativ
	IS
	--
	SELECT	/*+ RULE */ NU_SEQU_OPER_ATIV
	FROM	A8.TB_OPER_ATIV
	WHERE	DT_OPER_ATIV < DATA_CORTE
	AND		(DT_LIQU_OPER_ATIV < TRUNC (SYSDATE + 1) OR DT_LIQU_OPER_ATIV IS NULL)
	AND		(DT_OPER_ATIV_RETN < TRUNC (SYSDATE + 1) OR DT_OPER_ATIV_RETN IS NULL);
	--
	CURSOR curmesgrecbenvispb (pnu_sequ_oper_ativ NUMBER)
	IS
	--
	SELECT	NU_CTRL_IF
	,		DH_REGT_MESG_SPB
	FROM	A8.TB_MESG_RECB_ENVI_SPB
	WHERE	(NU_SEQU_OPER_ATIV = pnu_sequ_oper_ativ OR (NU_SEQU_OPER_ATIV IS NULL AND DH_REGT_MESG_SPB < data_corte));
	--
	CURSOR curcncloperativ (pnu_sequ_oper_ativ NUMBER)
	IS
	--
	SELECT	NU_CTRL_IF
	,		DH_REGT_MESG_SPB
	FROM	A8.TB_CNCL_OPER_ATIV
	WHERE	NU_SEQU_OPER_ATIV = pnu_sequ_oper_ativ;
	--
	BEGIN
  
    execute immediate 'alter session set session_cached_cursors=100';
  
		--
		-- Chamada da procedure para exclusividade de área de Rollback
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		-- Obtem a data de corte para o Expurgo
		obter_data_corte;
		qt_processada := 0;
		--
		FOR rowoperativ IN curoperativ
			LOOP
			--
			--************************************
			--Expurgo da tabela A8.TB_ALER_EMIT.
			--************************************
			--
			INSERT	INTO A8HIST.TB_ALER_EMIT
					(NU_SEQU_ALER_EMIT,
					NU_SEQU_OPER_ATIV,
					CO_OPER_ATIV,
					CO_FATO_GERA_ALER,
					DH_EMIS_ALER,
					TP_BKOF,
					NO_VEIC_LEGA,
					NO_CNPT,
					NO_TIPO_OPER,
					NU_COMD_OPER,
					VA_OPER_ATIV,
					DH_DESA_ALER,
					CO_USUA_ULTI_ATLZ,
					CO_ETCA_TRAB_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			SELECT	NU_SEQU_ALER_EMIT,
					NU_SEQU_OPER_ATIV,
					CO_OPER_ATIV,
					CO_FATO_GERA_ALER,
					DH_EMIS_ALER,
					TP_BKOF,
					NO_VEIC_LEGA,
					NO_CNPT,
					NO_TIPO_OPER,
					NU_COMD_OPER,
					VA_OPER_ATIV,
					DH_DESA_ALER,
					CO_USUA_ULTI_ATLZ,
					CO_ETCA_TRAB_ULTI_ATLZ,
					DH_ULTI_ATLZ
			FROM	A8.TB_ALER_EMIT
			WHERE	(DH_EMIS_ALER < data_corte AND NU_SEQU_OPER_ATIV IS NULL)
			OR		NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_ALER_EMIT
			WHERE	(DH_EMIS_ALER < data_corte AND NU_SEQU_OPER_ATIV IS NULL)
			OR		NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--****************************************************
			-- Expurgo da tabela A8.TB_HIST_ENVI_INFO_GEST_CAIX.
			--****************************************************
			--
			INSERT	INTO A8HIST.TB_TEXT_XML
					(CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML)
			SELECT	CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML
			FROM	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
									FROM	A8.TB_HIST_ENVI_INFO_GEST_CAIX
									WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ);
			--
			DELETE	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
									FROM	A8.TB_HIST_ENVI_INFO_GEST_CAIX
									WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ);
			--
			INSERT	INTO A8HIST.TB_HIST_ENVI_INFO_GEST_CAIX
					(NU_SEQU_OPER_ATIV,
					DH_ENVI_GEST_CAIX,
					CO_SITU_MOVI_GEST_CAIX,
					CO_TEXT_XML)
			SELECT	NU_SEQU_OPER_ATIV,
					DH_ENVI_GEST_CAIX,
					CO_SITU_MOVI_GEST_CAIX,
					CO_TEXT_XML
			FROM	A8.TB_HIST_ENVI_INFO_GEST_CAIX
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_HIST_ENVI_INFO_GEST_CAIX
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--***********************************************
			-- Expurgo da tabela A8.TB_OPER_ATIV_MESG_INTE.
			--***********************************************
			--
			INSERT	INTO A8HIST.TB_TEXT_XML
					(CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML)
			SELECT	CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML
			FROM	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
									FROM	A8.TB_OPER_ATIV_MESG_INTE
									WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ);
			--
			DELETE	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
									FROM	A8.TB_OPER_ATIV_MESG_INTE
									WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ);
			--
			INSERT	INTO A8HIST.TB_OPER_ATIV_MESG_INTE
					(NU_SEQU_OPER_ATIV,
					DH_MESG_INTE,
					TP_MESG_INTE,
					TP_SOLI_MESG_INTE,
					CO_TEXT_XML,
					TP_FORM_MESG_SAID)
			SELECT	NU_SEQU_OPER_ATIV,
					DH_MESG_INTE,
					TP_MESG_INTE,
					TP_SOLI_MESG_INTE,
					CO_TEXT_XML,
					TP_FORM_MESG_SAID
			FROM	A8.TB_OPER_ATIV_MESG_INTE
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_OPER_ATIV_MESG_INTE
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--****************************************************
			-- Expurgo da tabela A8.TB_HIST_SITU_ACAO_OPER_ATIV.
			--****************************************************
			--
			INSERT	INTO A8HIST.TB_HIST_SITU_ACAO_OPER_ATIV
					(NU_SEQU_OPER_ATIV,
					DH_SITU_ACAO_OPER_ATIV,
					CO_SITU_PROC,
					TP_ACAO_OPER_ATIV,
					TP_JUST_SITU_PROC,
					TX_CNTD_ANTE_ACAO,
					CO_USUA_ATLZ,
					CO_ETCA_USUA_ATLZ)
			SELECT	NU_SEQU_OPER_ATIV,
					DH_SITU_ACAO_OPER_ATIV,
					CO_SITU_PROC,
					TP_ACAO_OPER_ATIV,
					TP_JUST_SITU_PROC,
					TX_CNTD_ANTE_ACAO,
					CO_USUA_ATLZ,
					CO_ETCA_USUA_ATLZ
			FROM	A8.TB_HIST_SITU_ACAO_OPER_ATIV
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_HIST_SITU_ACAO_OPER_ATIV
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--***************************************************************************
			-- Expurgo da tabela A8.TB_CNCL_OPER_ATIV (nu_sequ_oper_ativ sem mensagem).
			--***************************************************************************
			--
			INSERT	INTO A8HIST.TB_CNCL_OPER_ATIV
					(NU_SEQU_CNCL_OPER_ATIV_MESG,
					NU_SEQU_OPER_ATIV,
					NU_CTRL_IF,
					DH_REGT_MESG_SPB,
					QT_ATIV_MERC_CNCL,
					NU_SEQU_CNTR_REPE)
			SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG,
					NU_SEQU_OPER_ATIV,
					NU_CTRL_IF,
					DH_REGT_MESG_SPB,
					QT_ATIV_MERC_CNCL,
					NU_SEQU_CNTR_REPE
			FROM	A8.TB_CNCL_OPER_ATIV
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ
			AND		NU_CTRL_IF IS NULL
			AND		DH_REGT_MESG_SPB IS NULL;
			--
			DELETE	A8.TB_CNCL_OPER_ATIV
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ
			AND		NU_CTRL_IF IS NULL
			AND		DH_REGT_MESG_SPB IS NULL;
			--
			--***************************************************************************
			--
			FOR rowcncloperativ IN curcncloperativ (rowoperativ.nu_sequ_oper_ativ)
				LOOP
				--
				--******************************************
				-- Expurgo da tabela A8.TB_CNCL_OPER_ATIV.
				--******************************************
				--
				INSERT	INTO A8HIST.TB_CNCL_OPER_ATIV
						(NU_SEQU_CNCL_OPER_ATIV_MESG,
						NU_SEQU_OPER_ATIV,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						QT_ATIV_MERC_CNCL,
						NU_SEQU_CNTR_REPE)
				SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG,
						NU_SEQU_OPER_ATIV,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						QT_ATIV_MERC_CNCL,
						NU_SEQU_CNTR_REPE
				FROM	A8.TB_CNCL_OPER_ATIV
				WHERE	NU_CTRL_IF = rowcncloperativ.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowcncloperativ.dh_regt_mesg_spb;
				--
				DELETE	A8.TB_CNCL_OPER_ATIV
				WHERE	NU_CTRL_IF = rowcncloperativ.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowcncloperativ.dh_regt_mesg_spb;
				--
				--**********************************************
				-- Expurgo da tabela A8.TB_MESG_RECB_SPB_CNCL.
				--**********************************************
				--
				INSERT	INTO A8HIST.TB_MESG_RECB_SPB_CNCL
						(NU_SEQU_CNTR_REPE,
						NU_CTRL_IF,
						DT_OPER,
						CO_OPER_SELIC,
						NU_COMD_OPER_ORIG,
						VA_FINC,
						DH_REGT_MESG_SPB,
						NU_ATIV_MERC,
						PU_ATIV_MERC,
						QT_ATIV_MERC,
						DT_LIQU,
						IN_OPER_DEBT_CRED,
						DT_VENC,
						CO_CNPJ_CNPT,
						NO_CNPT,
						CO_CNTA_CUTD_SELIC_VEIC_LEGA,
						TP_INFO_LDL,
						PE_TAXA_NEGO,
						CO_TITL_CUTD,
						TP_NEGO_BMA,
						TP_LIQU_OPER_ATIV,
						CO_OPER_CETIP,
						DE_ATIV_MERC,
						CO_CNTA_CUTD_CNPT,
						CO_ISPB_BANC_LIQU_CNPT,
						CO_AGEN_COTR,
						NU_CC_COTR,
						DT_OPER_ATIV_RETN,
						PZ_DIAS_RETN_OPER_ATIV,
						VA_OPER_ATIV_RETN,
						CO_CNPT_CAMR,
						CO_SISB_COTR,
						CO_CNAL_OPER_INTE)
				SELECT	/*+ RULE */ NU_SEQU_CNTR_REPE,
						NU_CTRL_IF,
						DT_OPER,
						CO_OPER_SELIC,
						NU_COMD_OPER_ORIG,
						VA_FINC,
						DH_REGT_MESG_SPB,
						NU_ATIV_MERC,
						PU_ATIV_MERC,
						QT_ATIV_MERC,
						DT_LIQU,
						IN_OPER_DEBT_CRED,
						DT_VENC,
						CO_CNPJ_CNPT,
						NO_CNPT,
						CO_CNTA_CUTD_SELIC_VEIC_LEGA,
						TP_INFO_LDL,
						PE_TAXA_NEGO,
						CO_TITL_CUTD,
						TP_NEGO_BMA,
						TP_LIQU_OPER_ATIV,
						CO_OPER_CETIP,
						DE_ATIV_MERC,
						CO_CNTA_CUTD_CNPT,
						CO_ISPB_BANC_LIQU_CNPT,
						CO_AGEN_COTR,
						NU_CC_COTR,
						DT_OPER_ATIV_RETN,
						PZ_DIAS_RETN_OPER_ATIV,
						VA_OPER_ATIV_RETN,
						CO_CNPT_CAMR,
						CO_SISB_COTR,
						CO_CNAL_OPER_INTE
				FROM	A8.TB_MESG_RECB_SPB_CNCL
				WHERE	NU_CTRL_IF = rowcncloperativ.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowcncloperativ.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /* clausula para uso do indice primario */;
				--
				DELETE	/*+ RULE */ A8.TB_MESG_RECB_SPB_CNCL
				WHERE	NU_CTRL_IF = rowcncloperativ.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowcncloperativ.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /* clausula para uso do indice primario */;
				--
				--***************************************************
				-- Expurgo da tabela A8.TB_HIST_SITU_ACAO_MESG_SPB.
				--***************************************************
				--
				INSERT	INTO A8HIST.TB_HIST_SITU_ACAO_MESG_SPB
						(NU_SEQU_CNTR_REPE,
						CO_USUA_ULTI_ATLZ,
						CO_ETCA_TRAB_ULTI_ATLZ,
						CO_SITU_PROC,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						TP_ACAO_MESG_SPB,
						TX_CNTD_ANTE_ACAO,
						DH_SITU_ACAO_MESG_SPB)
				SELECT	/*+ RULE */ NU_SEQU_CNTR_REPE,
						CO_USUA_ULTI_ATLZ,
						CO_ETCA_TRAB_ULTI_ATLZ,
						CO_SITU_PROC,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						TP_ACAO_MESG_SPB,
						TX_CNTD_ANTE_ACAO,
						DH_SITU_ACAO_MESG_SPB
				FROM	A8.TB_HIST_SITU_ACAO_MESG_SPB
				WHERE	NU_CTRL_IF = rowcncloperativ.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowcncloperativ.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /*clausula para uso do indice (tunning)*/;
				--
				DELETE	/*+ RULE */ A8.TB_HIST_SITU_ACAO_MESG_SPB
				WHERE	NU_CTRL_IF = rowcncloperativ.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowcncloperativ.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /*clausula para uso do indice (tunning)*/;
				--
				END LOOP;
			--
			--***************************************************************************
			--
			FOR rowmesgrecbenvispb IN curmesgrecbenvispb (rowoperativ.nu_sequ_oper_ativ)
				LOOP
				--
				--**************************
				-- Expurgo do CO_TEXT_XML
				--**************************
				--
				INSERT	INTO A8HIST.TB_TEXT_XML
						(CO_TEXT_XML,
						NU_SEQU_TEXT_XML,
						TX_XML)
				SELECT	CO_TEXT_XML,
						NU_SEQU_TEXT_XML,
						TX_XML
				FROM	A8.TB_TEXT_XML
				WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
										FROM	A8.TB_MESG_RECB_ENVI_SPB
										WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
										AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb);
				--
				DELETE	A8.TB_TEXT_XML
				WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
										FROM	A8.TB_MESG_RECB_ENVI_SPB
										WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
										AND DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb);
				--
				--******************************************
				-- Expurgo da tabela A8.TB_CNCL_OPER_ATIV.
				--******************************************
				--
				INSERT	INTO A8HIST.TB_CNCL_OPER_ATIV
						(NU_SEQU_CNCL_OPER_ATIV_MESG,
						NU_SEQU_OPER_ATIV,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						QT_ATIV_MERC_CNCL,
						NU_SEQU_CNTR_REPE)
				SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG,
						NU_SEQU_OPER_ATIV,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						QT_ATIV_MERC_CNCL,
						NU_SEQU_CNTR_REPE
				FROM	A8.TB_CNCL_OPER_ATIV
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb;
				--
				DELETE	A8.TB_CNCL_OPER_ATIV
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb;
				--
				--***************************************************
				-- Expurgo da tabela A8.TB_HIST_SITU_ACAO_MESG_SPB.
				--***************************************************
				--
				INSERT	INTO A8HIST.TB_HIST_SITU_ACAO_MESG_SPB
						(NU_SEQU_CNTR_REPE,
						CO_USUA_ULTI_ATLZ,
						CO_ETCA_TRAB_ULTI_ATLZ,
						CO_SITU_PROC,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						TP_ACAO_MESG_SPB,
						TX_CNTD_ANTE_ACAO,
						DH_SITU_ACAO_MESG_SPB)
				SELECT /*+ RULE */NU_SEQU_CNTR_REPE,
						CO_USUA_ULTI_ATLZ,
						CO_ETCA_TRAB_ULTI_ATLZ,
						CO_SITU_PROC,
						NU_CTRL_IF,
						DH_REGT_MESG_SPB,
						TP_ACAO_MESG_SPB,
						TX_CNTD_ANTE_ACAO,
						DH_SITU_ACAO_MESG_SPB
				FROM	A8.TB_HIST_SITU_ACAO_MESG_SPB
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /*clausula para uso do indice (tunning)*/;
				--
				DELETE	/*+ RULE */ A8.TB_HIST_SITU_ACAO_MESG_SPB
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /*clausula para uso do indice (tunning)*/;
				--
				--**********************************************
				-- Expurgo da tabela A8.TB_MESG_RECB_SPB_CNCL.
				--**********************************************
				--
				INSERT	INTO A8HIST.TB_MESG_RECB_SPB_CNCL
						(NU_CTRL_IF,
						NU_SEQU_CNTR_REPE,
						DH_REGT_MESG_SPB,
						TP_LIQU_OPER_ATIV,
						DT_OPER,
						CO_OPER_SELIC,
						NU_COMD_OPER_ORIG,
						NU_ATIV_MERC,
						PU_ATIV_MERC,
						QT_ATIV_MERC,
						VA_FINC,
						DT_LIQU,
						IN_OPER_DEBT_CRED,
						DT_VENC,
						CO_CNPJ_CNPT,
						NO_CNPT,
						CO_CNTA_CUTD_SELIC_VEIC_LEGA,
						TP_INFO_LDL,
						PE_TAXA_NEGO,
						CO_TITL_CUTD,
						TP_NEGO_BMA,
						CO_OPER_CETIP,
						DE_ATIV_MERC,
						CO_CNTA_CUTD_CNPT,
						CO_ISPB_BANC_LIQU_CNPT,
						CO_AGEN_COTR,
						NU_CC_COTR,
						DT_OPER_ATIV_RETN,
						PZ_DIAS_RETN_OPER_ATIV,
						VA_OPER_ATIV_RETN,
						CO_CNPT_CAMR)
				SELECT	/*+ RULE */NU_CTRL_IF,
						NU_SEQU_CNTR_REPE,
						DH_REGT_MESG_SPB,
						TP_LIQU_OPER_ATIV,
						DT_OPER,
						CO_OPER_SELIC,
						NU_COMD_OPER_ORIG,
						NU_ATIV_MERC,
						PU_ATIV_MERC,
						QT_ATIV_MERC,
						VA_FINC,
						DT_LIQU,
						IN_OPER_DEBT_CRED,
						DT_VENC,
						CO_CNPJ_CNPT,
						NO_CNPT,
						CO_CNTA_CUTD_SELIC_VEIC_LEGA,
						TP_INFO_LDL,
						PE_TAXA_NEGO,
						CO_TITL_CUTD,
						TP_NEGO_BMA,
						CO_OPER_CETIP,
						DE_ATIV_MERC,
						CO_CNTA_CUTD_CNPT,
						CO_ISPB_BANC_LIQU_CNPT,
						CO_AGEN_COTR,
						NU_CC_COTR,
						DT_OPER_ATIV_RETN,
						PZ_DIAS_RETN_OPER_ATIV,
						VA_OPER_ATIV_RETN,
						CO_CNPT_CAMR
				FROM	A8.TB_MESG_RECB_SPB_CNCL
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /*clausula para uso do indice (tunning)*/;
				--
				DELETE	/*+ RULE */ A8.TB_MESG_RECB_SPB_CNCL
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb
				AND		NU_SEQU_CNTR_REPE > -2 /*clausula para uso do indice (tunning)*/;
				--
				--**********************************************
				-- Expurgo da tabela A8.TB_MESG_RECB_ENVI_SPB.
				--**********************************************
				--
				INSERT	INTO A8HIST.TB_MESG_RECB_ENVI_SPB
						(NU_CTRL_IF,
						NU_SEQU_CNTR_REPE,
						NU_SEQU_OPER_ATIV,
						CO_TEXT_XML,
						CO_VEIC_LEGA,
						CO_USUA_ULTI_ATLZ,
						SG_SIST,
						CO_ETCA_TRAB_ULTI_ATLZ,
						DH_REGT_MESG_SPB,
						CO_ULTI_SITU_PROC,
						CO_MESG_SPB,
						TP_BKOF,
						CO_EMPR,
						NU_COMD_OPER,
						CO_SITU_MESG_SPB,
						DH_RECB_ENVI_MESG_SPB,
						HO_ENVI_MESG_SPB,
						DH_ULTI_ATLZ,
						IN_ENTR_MANU,
						NU_SEQU_CNCL_OPER_ATIV_MESG,
						CO_LOCA_LIQU,
						NU_CTRL_CAMR,
						IN_CONF_MESG_LTR,
						CO_ISPB_PART_CAMR)
				SELECT	NU_CTRL_IF,
						NU_SEQU_CNTR_REPE,
						NU_SEQU_OPER_ATIV,
						CO_TEXT_XML,
						CO_VEIC_LEGA,
						CO_USUA_ULTI_ATLZ,
						SG_SIST,
						CO_ETCA_TRAB_ULTI_ATLZ,
						DH_REGT_MESG_SPB,
						CO_ULTI_SITU_PROC,
						CO_MESG_SPB,
						TP_BKOF,
						CO_EMPR,
						NU_COMD_OPER,
						CO_SITU_MESG_SPB,
						DH_RECB_ENVI_MESG_SPB,
						HO_ENVI_MESG_SPB,
						DH_ULTI_ATLZ,
						IN_ENTR_MANU,
						NU_SEQU_CNCL_OPER_ATIV_MESG,
						CO_LOCA_LIQU,
						NU_CTRL_CAMR,
						IN_CONF_MESG_LTR,
						CO_ISPB_PART_CAMR
				FROM	A8.TB_MESG_RECB_ENVI_SPB
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb;
				--
				DELETE	A8.TB_MESG_RECB_ENVI_SPB
				WHERE	NU_CTRL_IF = rowmesgrecbenvispb.nu_ctrl_if
				AND		DH_REGT_MESG_SPB = rowmesgrecbenvispb.dh_regt_mesg_spb;
				--
				END LOOP;
			--
			--*********************************************
			-- Expurgo da tabela A8.TB_HIST_ENVI_INFO_CC.
			--*********************************************
			--
			INSERT	INTO A8HIST.TB_TEXT_XML
					(CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML)
			SELECT	/*+ RULE */ CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML
			FROM	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
									FROM	A8.TB_HIST_ENVI_INFO_CC
									WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ);
			--
			DELETE	/*+ RULE */ A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML IN (SELECT	CO_TEXT_XML
									FROM	A8.TB_HIST_ENVI_INFO_CC
									WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ);
			--
			INSERT	INTO A8HIST.TB_HIST_ENVI_INFO_CC
					(NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					SG_SIST,
					CO_EMPR,
					DH_ENVI_CC,
					TP_ENVI_RETN,
					CO_TEXT_XML)
			SELECT	/*+ RULE */ NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					SG_SIST,
					CO_EMPR,
					DH_ENVI_CC,
					TP_ENVI_RETN,
					CO_TEXT_XML
			FROM	A8.TB_HIST_ENVI_INFO_CC
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	/*+ RULE */ A8.TB_HIST_ENVI_INFO_CC
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--***********************************************
			-- Expurgo da tabela A8.TB_HIST_ENVI_INFO_CNTB.
			--***********************************************
			--
			INSERT	INTO A8HIST.TB_HIST_ENVI_INFO_CNTB
					(NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					SG_SIST,
					CO_EMPR,
					DH_ENVI_CNTB,
					IN_LANC_DEBT_CRED,
					TX_ITGR_CNTB,
					IN_ITGR_CNTB)
			SELECT	NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					SG_SIST,
					CO_EMPR,
					DH_ENVI_CNTB,
					IN_LANC_DEBT_CRED,
					TX_ITGR_CNTB,
					IN_ITGR_CNTB
			FROM	A8.TB_HIST_ENVI_INFO_CNTB
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_HIST_ENVI_INFO_CNTB
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--**************************************************
			-- Expurgo da tabela A8.TB_HIST_SITU_LANC_CC_CNTB.
			--**************************************************
			--
			INSERT	INTO A8HIST.TB_HIST_SITU_LANC_CC_CNTB
					(DH_SITU_LANC_CC,
					NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					CO_SITU_PROC,
					CO_ERRO,
					TX_MESG_ERRO,
					TX_JUST_CANC,
					CO_USUA_ATLZ,
					CO_ETCA_TRAB_ATLZ)
			SELECT	/*+ RULE */ DH_SITU_LANC_CC,
					NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					CO_SITU_PROC,
					CO_ERRO,
					TX_MESG_ERRO,
					TX_JUST_CANC,
					CO_USUA_ATLZ,
					CO_ETCA_TRAB_ATLZ
			FROM	A8.TB_HIST_SITU_LANC_CC_CNTB
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	/*+ RULE */ A8.TB_HIST_SITU_LANC_CC_CNTB
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--****************************************
			-- Expurgo da tabela A8.TB_LANC_CC_CNTB.
			--****************************************
			--
			INSERT	INTO A8HIST.TB_LANC_CC_CNTB
					(NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					TP_BKOF,
					CO_VEIC_LEGA,
					SG_SIST,
					SG_SIST_CC,
					CO_EMPR,
					TP_OPER,
					CO_LOCA_LIQU,
					IN_LANC_DEBT_CRED,
					DT_OPER,
					CO_BANC,
					CO_AGEN,
					NU_CC,
					NU_COMD_OPER,
					VA_LANC_CC,
					NU_CTRL_ITGR,
					CO_SUB_TIPO_ATIV,
					CO_ULTI_SITU_PROC,
					CO_USUA_ULTI_ATLZ,
					CO_ETCA_TRAB_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			SELECT	NU_SEQU_OPER_ATIV,
					TP_LANC_ITGR,
					TP_BKOF,
					CO_VEIC_LEGA,
					SG_SIST,
					SG_SIST_CC,
					CO_EMPR,
					TP_OPER,
					CO_LOCA_LIQU,
					IN_LANC_DEBT_CRED,
					DT_OPER,
					CO_BANC,
					CO_AGEN,
					NU_CC,
					NU_COMD_OPER,
					VA_LANC_CC,
					NU_CTRL_ITGR,
					CO_SUB_TIPO_ATIV,
					CO_ULTI_SITU_PROC,
					CO_USUA_ULTI_ATLZ,
					CO_ETCA_TRAB_ULTI_ATLZ,
					DH_ULTI_ATLZ
			FROM	A8.TB_LANC_CC_CNTB
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_LANC_CC_CNTB
			WHERE NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--*************************************
			-- Expurgo da tabela A8.TB_OPER_ATIV.
			--*************************************
			--
			INSERT	INTO A8HIST.TB_OPER_ATIV
					(NU_SEQU_OPER_ATIV,
					TP_OPER,
					CO_LOCA_LIQU,
					TP_LIQU_OPER_ATIV,
					CO_EMPR,
					CO_USUA_CADR_OPER,
					HO_ENVI_MESG_SPB,
					CO_OPER_ATIV,
					NU_COMD_OPER,
					NU_COMD_OPER_RETN,
					DT_OPER_ATIV,
					DT_OPER_ATIV_RETN,
					CO_VEIC_LEGA,
					SG_SIST,
					CO_CNTA_CUTD_SELIC_VEIC_LEGA,
					CO_CNPJ_CNPT,
					NO_CNPT,
					CO_CNTA_CUTD_SELIC_CNPT,
					IN_OPER_DEBT_CRED,
					NU_ATIV_MERC,
					DE_ATIV_MERC,
					PU_ATIV_MERC,
					QT_ATIV_MERC,
					IN_ENTR_SAID_RECU_FINC,
					DT_VENC_ATIV,
					VA_OPER_ATIV,
					VA_OPER_ATIV_REAJ,
					DT_LIQU_OPER_ATIV,
					TP_CPRO_OPER_ATIV,
					TP_CPRO_RETN_OPER_ATIV,
					IN_DISP_CONS,
					IN_ENVI_PREV_SIST_PJ,
					IN_ENVI_RELZ_SIST_PJ,
					IN_ENVI_PREV_SIST_A6,
					IN_ENVI_RELZ_SIST_A6,
					CO_ULTI_SITU_PROC,
					TP_ACAO_OPER_ATIV_EXEC,
					NU_COMD_ACAO_EXEC,
					IN_ENTR_MANU,
					NU_PRTC_OPER_LG,
					NU_SEQU_CNCL_OPER_ATIV_MESG,
					NU_CTRL_MESG_SPB_ORIG,
					PE_TAXA_NEGO,
					CO_TITL_CUTD,
					CO_USUA_ULTI_ATLZ,
					CO_ETCA_TRAB_ULTI_ATLZ,
					DH_ULTI_ATLZ,
					CO_OPER_CETIP,
					CO_ISPB_BANC_LIQU_CNPT,
					TP_IF_CRED_DEBT,
					CO_AGEN_COTR,
					NU_CC_COTR,
					PZ_DIAS_RETN_OPER_ATIV,
					VA_OPER_ATIV_RETN,
					TP_CNPT,
					CO_CNPT_CAMR,
					CO_IDEF_LAST,
					CO_SISB_COTR,
					CO_CNAL_OPER_INTE,
					CO_SITU_PROC_MESG_SPB_RECB, 
					NR_CNPJ_CPF)
			SELECT	NU_SEQU_OPER_ATIV,
					TP_OPER,
					CO_LOCA_LIQU,
					TP_LIQU_OPER_ATIV,
					CO_EMPR,
					CO_USUA_CADR_OPER,
					HO_ENVI_MESG_SPB,
					CO_OPER_ATIV,
					NU_COMD_OPER,
					NU_COMD_OPER_RETN,
					DT_OPER_ATIV,
					DT_OPER_ATIV_RETN,
					CO_VEIC_LEGA,
					SG_SIST,
					CO_CNTA_CUTD_SELIC_VEIC_LEGA,
					CO_CNPJ_CNPT,
					NO_CNPT,
					CO_CNTA_CUTD_SELIC_CNPT,
					IN_OPER_DEBT_CRED,
					NU_ATIV_MERC,
					DE_ATIV_MERC,
					PU_ATIV_MERC,
					QT_ATIV_MERC,
					IN_ENTR_SAID_RECU_FINC,
					DT_VENC_ATIV,
					VA_OPER_ATIV,
					VA_OPER_ATIV_REAJ,
					DT_LIQU_OPER_ATIV,
					TP_CPRO_OPER_ATIV,
					TP_CPRO_RETN_OPER_ATIV,
					IN_DISP_CONS,
					IN_ENVI_PREV_SIST_PJ,
					IN_ENVI_RELZ_SIST_PJ,
					IN_ENVI_PREV_SIST_A6,
					IN_ENVI_RELZ_SIST_A6,
					CO_ULTI_SITU_PROC,
					TP_ACAO_OPER_ATIV_EXEC,
					NU_COMD_ACAO_EXEC,
					IN_ENTR_MANU,
					NU_PRTC_OPER_LG,
					NU_SEQU_CNCL_OPER_ATIV_MESG,
					NU_CTRL_MESG_SPB_ORIG,
					PE_TAXA_NEGO,
					CO_TITL_CUTD,
					CO_USUA_ULTI_ATLZ,
					CO_ETCA_TRAB_ULTI_ATLZ,
					DH_ULTI_ATLZ,
					CO_OPER_CETIP,
					CO_ISPB_BANC_LIQU_CNPT,
					TP_IF_CRED_DEBT,
					CO_AGEN_COTR,
					NU_CC_COTR,
					PZ_DIAS_RETN_OPER_ATIV,
					VA_OPER_ATIV_RETN,
					TP_CNPT,
					CO_CNPT_CAMR,
					CO_IDEF_LAST,
					CO_SISB_COTR,
					CO_CNAL_OPER_INTE,
					CO_SITU_PROC_MESG_SPB_RECB, 
					NR_CNPJ_CPF
			FROM	A8.TB_OPER_ATIV
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			DELETE	A8.TB_OPER_ATIV
			WHERE	NU_SEQU_OPER_ATIV = rowoperativ.nu_sequ_oper_ativ;
			--
			--Controle de bloco para COMMIT .
			--
			qt_processada := qt_processada + 1;
			--
			IF qt_processada >= 100 THEN
				COMMIT;
				SYS.dbP_muda_rbs('RBSBIG','ON');
				qt_processada := 0;
			END IF;
			--
			END LOOP;
		--
		-- Commit nos ultimos comandos executados (bloco menor que 100)
		--
		COMMIT;
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		--****************************************************
		-- Expurgo da tabela A8.TB_JUST_CNCL_OPER_ATIV_MESG.
		--****************************************************
		--
		INSERT	INTO A8HIST.TB_JUST_CNCL_OPER_ATIV_MESG
				(NU_SEQU_CNCL_OPER_ATIV_MESG,
				TP_JUST_CNCL,
				TX_JUST,
				CO_USUA_ATLZ,
				CO_ETCA_TRAB_ATLZ,
				DH_JUST_CNCL)
		SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG,
				TP_JUST_CNCL,
				TX_JUST,
				CO_USUA_ATLZ,
				CO_ETCA_TRAB_ATLZ,
				DH_JUST_CNCL
		FROM	A8.TB_JUST_CNCL_OPER_ATIV_MESG
		WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IN (SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG
													FROM	A8.TB_OPER_ATIV
													WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IS NULL
													UNION ALL
													SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG
													FROM	A8.TB_CNCL_OPER_ATIV
													WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IS NULL
													UNION ALL
													SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG
													FROM	A8.TB_MESG_RECB_ENVI_SPB
													WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IS NULL);
		--
		DELETE	A8.TB_JUST_CNCL_OPER_ATIV_MESG
		WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IN (SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG
													FROM	A8.TB_OPER_ATIV
													WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IS NULL
													UNION ALL
													SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG
													FROM	A8.TB_CNCL_OPER_ATIV
													WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IS NULL
													UNION ALL
													SELECT	NU_SEQU_CNCL_OPER_ATIV_MESG
													FROM	A8.TB_MESG_RECB_ENVI_SPB
													WHERE	NOT NU_SEQU_CNCL_OPER_ATIV_MESG IS NULL);
		--
		COMMIT;
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		--*************************************
		-- Expurgo da tabela A8.TB_REME_REJE.
		--*************************************
		--
		FOR reg IN (SELECT CO_TEXT_XML_REJE FROM A8.TB_REME_REJE WHERE DH_REME_REJE < data_corte)
			LOOP
			INSERT	INTO A8HIST.TB_TEXT_XML
					(CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML)
			SELECT	CO_TEXT_XML,
					NU_SEQU_TEXT_XML,
					TX_XML
			FROM	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML = reg.co_text_xml_reje;
			--
			DELETE	A8.TB_TEXT_XML
			WHERE	CO_TEXT_XML = reg.co_text_xml_reje;
			--
			END LOOP;
		--
		INSERT	INTO A8HIST.TB_REME_REJE
				(TX_XML_ERRO,
				CO_TEXT_XML_REJE,
				DH_REME_REJE,
				TP_MESG_INTE,
				CO_EMPR,
				SG_SIST_ORIG_INFO,
				CO_TEXT_XML_RETN_SIST_ORIG)
		SELECT	TX_XML_ERRO,
				CO_TEXT_XML_REJE,
				DH_REME_REJE,
				TP_MESG_INTE,
				CO_EMPR,
				SG_SIST_ORIG_INFO,
				CO_TEXT_XML_RETN_SIST_ORIG
		FROM	A8.TB_REME_REJE
		WHERE	DH_REME_REJE < data_corte;
		--
		DELETE	A8.TB_REME_REJE
		WHERE	DH_REME_REJE < data_corte;
		--
		COMMIT;
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		gravarhistoricoexecucao (5, 1, NULL);
		--
		COMMIT;
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		-- Chamada da procedure para exclusividade de área de Rollback
		--
		SYS.dbP_muda_rbs('RBSBIG','OFF');
		--
	EXCEPTION
		WHEN OTHERS THEN
			ROLLBACK;
			gravarhistoricoexecucao (5, 2, SUBSTR (SQLERRM, 1, 200));
			COMMIT;
			DBMS_OUTPUT.put_line (SQLCODE || ' - ' || SQLERRM);
			-- Chamada da procedure para exclusividade de área de Rollback
			SYS.dbP_muda_rbs('RBSBIG','OFF');
	--
	END processar;
	--
END a8k_expurgo;
/

GRANT EXECUTE ON A8K_EXPURGO TO A7PROC;
GRANT EXECUTE ON A8K_EXPURGO TO A6PROC;
GRANT EXECUTE ON A8K_EXPURGO TO RL01_A8_MANUTENCAO;
GRANT EXECUTE ON A8K_EXPURGO TO RL02_A8_CONSULTA;

