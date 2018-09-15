CREATE OR REPLACE PACKAGE A7K_EXPURGO AS
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
--
/
--
CREATE OR REPLACE PACKAGE BODY A7K_EXPURGO
--
IS
--
	data_corte			DATE;
--
PROCEDURE gravarhistoricoexecucao
(	protina		IN		NUMBER
,	pstatus		IN		NUMBER
,	perro		IN		VARCHAR2
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
		VALUES	(protina,
				SYSDATE,
				pstatus,
				SUBSTR (perro, 1, 200));
	--
	EXCEPTION
		WHEN OTHERS
			THEN
				raise_application_error (SQLCODE,'ERRO NA ROTINA DE GRAVACAO DE HISTORICO DE EXECUCAO: ' || SQLERRM);
	--
	END gravarhistoricoexecucao;
--
PROCEDURE obter_data_corte
--
IS
	--
	vdatacorte			DATE;
	vdiascorte			NUMBER;
	--
	BEGIN
	--
		vdiascorte := a8proc.a8k_expurgo.obterqtddiasexpurgo () - 1;
		--
		SELECT	TRUNC ((SYSDATE - vdiascorte))
		INTO	data_corte
		FROM	dual;
	--
	END obter_data_corte;
--
PROCEDURE msg 
(	mensagem	IN		VARCHAR2
)
--
IS
	--
	BEGIN
		--
		DBMS_OUTPUT.put_line (TO_CHAR (SYSDATE, 'yyyy-mm-dd hh24:mi:ss') || ': ' || mensagem);
		--
	END;
--
PROCEDURE processar
--
IS
	--
	BEGIN
		--
		--Chamada da procedure para exclusividade de área de Rollback
		--
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		--Otem a data de corte para o expurgo
		--
		obter_data_corte;
		--
		msg ('Expurgo das tabelas A7.TB_SITU_MESG e A7.TB_MESG (Filha e Pai).');
		--
		--Expurgo das tabelas A7.TB_SITU_MESG e A7.TB_MESG (Filha e Pai)
		--
		INSERT	INTO A7HIST.TB_MESG
				(CO_MESG,
				DH_MESG,
				TP_MESG,
				DH_INIC_VIGE_REGR_TRAP,
				CO_MESG_MQSE,
				CO_CMPO_ATRB_IDEF_MESG,
				CO_TEXT_XML_ENTR,
				CO_TEXT_XML_SAID,
				SG_SIST_ORIG,
				CO_EMPR_ORIG,
				TP_FORM_MESG_SAID)
		SELECT	CO_MESG,
				DH_MESG,
				TP_MESG,
				DH_INIC_VIGE_REGR_TRAP,
				CO_MESG_MQSE,
				CO_CMPO_ATRB_IDEF_MESG,
				CO_TEXT_XML_ENTR,
				CO_TEXT_XML_SAID,
				SG_SIST_ORIG,
				CO_EMPR_ORIG,
				TP_FORM_MESG_SAID
		FROM	A7.TB_MESG
		WHERE	DH_MESG < data_corte;
		--
		INSERT	INTO A7HIST.TB_TEXT_XML
				(CO_TEXT_XML,
				NU_SEQU_TEXT_XML,
				TX_XML)
		SELECT	A.CO_TEXT_XML,
				A.NU_SEQU_TEXT_XML,
				A.TX_XML
		FROM	A7.TB_TEXT_XML A,
				A7.TB_MESG B
		WHERE	(A.CO_TEXT_XML = B.CO_TEXT_XML_ENTR OR A.CO_TEXT_XML = B.CO_TEXT_XML_SAID)
		AND		B.DH_MESG < data_corte;
		--
		INSERT	INTO A7HIST.TB_SITU_MESG
				(CO_OCOR_MESG,
				CO_MESG,
				NU_SEQU_SITU_MESG,
				DH_OCOR_MESG,
				TX_DTLH_OCOR_ERRO)
		SELECT	CO_OCOR_MESG,
				CO_MESG,
				NU_SEQU_SITU_MESG,
				DH_OCOR_MESG,
				TX_DTLH_OCOR_ERRO
		FROM	A7.TB_SITU_MESG
		WHERE	CO_MESG IN	(SELECT	CO_MESG
							FROM	A7.TB_MESG
							WHERE	DH_MESG < data_corte);
		--
		DELETE	A7.TB_SITU_MESG
		WHERE	CO_MESG IN	(SELECT	CO_MESG
							FROM	A7.TB_MESG
							WHERE	DH_MESG < data_corte);
		--
		DELETE	A7.TB_TEXT_XML
		WHERE	CO_TEXT_XML IN	(SELECT	CO_TEXT_XML_ENTR
								FROM	A7.TB_MESG
								WHERE	DH_MESG < data_corte
								UNION ALL
								SELECT	CO_TEXT_XML_SAID
								FROM	A7.TB_MESG
								WHERE	DH_MESG < data_corte);
		--
		DELETE	A7.TB_MESG
		WHERE	DH_MESG < data_corte;
		--
		COMMIT;
		--
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		msg ('Expurgo da tabela A7.TB_MESG_REJE.');
		--
		--Expurgo da tabela A7.TB_MESG_REJE.
		--
		INSERT	INTO A7HIST.TB_TEXT_XML
				(CO_TEXT_XML,
				NU_SEQU_TEXT_XML,
				TX_XML)
		SELECT	CO_TEXT_XML,
				NU_SEQU_TEXT_XML,
				TX_XML
		FROM	A7.TB_TEXT_XML
		WHERE	CO_TEXT_XML IN	(SELECT	CO_TEXT_XML
								FROM	A7.TB_MESG_REJE
								WHERE	DH_RECB_MESG < data_corte);
		--
		DELETE	A7.TB_TEXT_XML
		WHERE	CO_TEXT_XML IN	(SELECT	CO_TEXT_XML
								FROM	A7.TB_MESG_REJE
								WHERE	DH_RECB_MESG < data_corte);
		--
		INSERT	INTO A7HIST.TB_MESG_REJE
				(CO_MESG_MQSE,
				CO_OCOR_MESG,
				DH_RECB_MESG,
				NO_ARQU_ENTR_FILA_MQSE,
				DH_ENTR_FILA_MQSE,
				CO_TEXT_XML,
				TX_DTLH_OCOR_ERRO)
		SELECT	CO_MESG_MQSE,
				CO_OCOR_MESG,
				DH_RECB_MESG,
				NO_ARQU_ENTR_FILA_MQSE,
				DH_ENTR_FILA_MQSE,
				CO_TEXT_XML,
				TX_DTLH_OCOR_ERRO
		FROM	A7.TB_MESG_REJE
		WHERE	DH_RECB_MESG < data_corte;
		--
		DELETE	A7.TB_MESG_REJE
		WHERE	DH_RECB_MESG < data_corte;
		--
		COMMIT;
		--
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		msg ('Expurgo da tabela A7.TB_HIST_EXEC_ROTI_BATCH.');
		--
		--Expurgo da tabela A7.TB_HIST_EXEC_ROTI_BATCH.
		--
		INSERT	INTO A7HIST.TB_HIST_EXEC_ROTI_BATCH
				(CO_ROTI_BATCH,
				DH_FIM_EXEC,
				IN_EXEC_SUCE,
				DE_ERRO_EXEC)
		SELECT	CO_ROTI_BATCH,
				DH_FIM_EXEC,
				IN_EXEC_SUCE,
				DE_ERRO_EXEC
		FROM	A7.TB_HIST_EXEC_ROTI_BATCH
		WHERE	DH_FIM_EXEC < data_corte;
		--
		DELETE	A7.TB_HIST_EXEC_ROTI_BATCH
		WHERE	DH_FIM_EXEC < data_corte;
		--
		COMMIT;
		--
		SYS.dbP_muda_rbs('RBSBIG','ON');
		--
		msg ('Fim do expurgo do sistema A7.');
		--
		gravarhistoricoexecucao (4, 1, NULL);
		--
		COMMIT;
		--
		-- Chamada da procedure para exclusividade de área de Rollback
		--
		SYS.dbP_muda_rbs('RBSBIG','OFF');
		--
		EXCEPTION
			WHEN OTHERS
				THEN
					ROLLBACK;
					--
					gravarhistoricoexecucao (4, 2, SUBSTR (SQLERRM, 1, 200));
					--
					COMMIT;
					--
					DBMS_OUTPUT.put_line (SQLCODE || ' - ' || SQLERRM);
					--
					--Chamada da procedure para exclusividade de área de Rollback
					--
					SYS.dbP_muda_rbs('RBSBIG','OFF');
	END processar;
	--
END A7K_EXPURGO;
--
/
--
GRANT EXECUTE ON A7K_EXPURGO TO RL01_A7_MANUTENCAO;
GRANT EXECUTE ON A7K_EXPURGO TO RL02_A7_CONSULTA;
