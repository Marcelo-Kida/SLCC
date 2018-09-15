--
--********************************************************************************************
--Sistema 				: A8
--Criaçao de Procedure	: A8PROC.A8K_REPLICACAO
--Descrição				: Esta Package é utilizada para a replicação diária das tabelas
--						  dos sistemas PJ e PK para o sistema A8.
--Data Criaçao			: 19/01/2004
--********************************************************************************************
--Data da alteraçao		: 24/03/2004
--Objetivo da alteraçao	: Inclusão de Rotina para gravação de histórico das execuções
--********************************************************************************************
--Data da alteraçao		: 26/08/2005
--Objetivo da alteraçao	: Retirada da replicação das tabelas do Banespa, pois no Projeto
--						: de Integração, os sistemas PJ e PK estão com um ambiente somente
--********************************************************************************************
--
CREATE OR REPLACE PACKAGE A8K_REPLICACAO AS
	--
	PROCEDURE REPLICA_TABELAS_PJPK;
	--
	PROCEDURE REPLICA_EMPRESA_FUSIONADA;
	PROCEDURE REPLICA_EMPRESA_HO;
	PROCEDURE REPLICA_FERIADO_HO;
	PROCEDURE REPLICA_LOCAL_LIQUIDACAO;
	PROCEDURE REPLICA_PRODUTO;
	PROCEDURE REPLICA_EVENTO_FINANCEIRO;
	PROCEDURE REPLICA_SEGMENTO;
	PROCEDURE REPLICA_TIPO_CONTA;
	PROCEDURE REPLICA_INDEXADOR;
	--
	PROCEDURE REPLICA_ERRO_BACEN;
	PROCEDURE REPLICA_INSTITUICAO_SPB;
	PROCEDURE REPLICA_GRADE_HORARIO;
	PROCEDURE REPLICA_GRUPO;
	PROCEDURE REPLICA_SERVICO;
	PROCEDURE REPLICA_EVENTO;
	PROCEDURE REPLICA_MENSAGEM;
	PROCEDURE REPLICA_TIPO_TAG;
	PROCEDURE REPLICA_TAG;
	PROCEDURE REPLICA_TAG_MENSAGEM;
	PROCEDURE REPLICA_DOMINIO;
	PROCEDURE REPLICA_GRADE_MENSAGEM;
	PROCEDURE REPLICA_SIST_MENSAGEM;
	--
END;
/
--
SHOW ERRORS
--
CREATE OR REPLACE PACKAGE BODY A8K_REPLICACAO AS
--
--*******************************************************************
--
PROCEDURE	GRAVARHISTORICOEXECUCAO (pRotina	IN	number,
									pStatus		IN	number,
									pErro		IN	varchar2)
--
IS
--
	BEGIN
		INSERT	INTO A7.TB_HIST_EXEC_ROTI_BATCH
				(CO_ROTI_BATCH,
				DH_FIM_EXEC,
				IN_EXEC_SUCE,
				DE_ERRO_EXEC)
		VALUES	(pRotina,
				SYSDATE,
				pStatus,
				SUBSTR(pErro, 1, 200));
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(SQLCODE, 'ERRO NA ROTINA DE GRAVACAO DE HISTORICO DE EXECUCAO: ' || SQLERRM);
--
END GRAVARHISTORICOEXECUCAO;
--
--*******************************************************************
--
PROCEDURE REPLICA_TABELAS_PJPK
--
AS
--
	BEGIN
		--
		A8PROC.a8k_replicacao.REPLICA_EMPRESA_FUSIONADA;
		A8PROC.a8k_replicacao.REPLICA_EMPRESA_HO;
		A8PROC.a8k_replicacao.REPLICA_FERIADO_HO;
		A8PROC.a8k_replicacao.REPLICA_LOCAL_LIQUIDACAO;
		A8PROC.a8k_replicacao.REPLICA_PRODUTO;
		A8PROC.a8k_replicacao.REPLICA_EVENTO_FINANCEIRO;
		A8PROC.a8k_replicacao.REPLICA_SEGMENTO;
		A8PROC.a8k_replicacao.REPLICA_TIPO_CONTA;
		A8PROC.a8k_replicacao.REPLICA_INDEXADOR;
		--
		A8PROC.a8k_replicacao.REPLICA_ERRO_BACEN;
		A8PROC.a8k_replicacao.REPLICA_INSTITUICAO_SPB;
		A8PROC.a8k_replicacao.REPLICA_GRADE_HORARIO;
		A8PROC.a8k_replicacao.REPLICA_GRUPO;
		A8PROC.a8k_replicacao.REPLICA_SERVICO;
		A8PROC.a8k_replicacao.REPLICA_EVENTO;
		A8PROC.a8k_replicacao.REPLICA_MENSAGEM;
		A8PROC.a8k_replicacao.REPLICA_TIPO_TAG;
		A8PROC.a8k_replicacao.REPLICA_TAG;
		A8PROC.a8k_replicacao.REPLICA_TAG_MENSAGEM;
		A8PROC.a8k_replicacao.REPLICA_DOMINIO;
		A8PROC.a8k_replicacao.REPLICA_GRADE_MENSAGEM;
		A8PROC.a8k_replicacao.REPLICA_SIST_MENSAGEM;
		--
	COMMIT;
	--
	GRAVARHISTORICOEXECUCAO (1,1,NULL);
	--
	COMMIT;
	--
	EXCEPTION
		WHEN OTHERS THEN
			ROLLBACK;
			GRAVARHISTORICOEXECUCAO (1,2,SUBSTR(SQLERRM, 1, 200));
			COMMIT;
			--
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_TABELAS_PJPK'
			|| SQLCODE || ' - '
			|| SUBSTR(SQLERRM, 1, 200));
--
END ;
--
--*******************************************************************
--
PROCEDURE REPLICA_EMPRESA_FUSIONADA
--
AS
--
	CURSOR cEmpresaFusionada IS
		SELECT	CO_EMPR_FUSI,
				DE_EMPR_FUSI,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_EMPRESA_FUSIONADA;
	--
	BEGIN
	--
		FOR cAux in cEmpresaFusionada LOOP
		--
			UPDATE	A8.TB_EMPRESA_FUSIONADA
			SET		DE_EMPR_FUSI		= cAux.DE_EMPR_FUSI,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'Sistema',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_EMPRESA_FUSIONADA
						(CO_EMPR_FUSI,
						DE_EMPR_FUSI,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.DE_EMPR_FUSI,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'Sistema',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_EMPRESA_FUSIONADA'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END ;
--
--*******************************************************************
--
PROCEDURE REPLICA_FERIADO_HO
--
AS
--
	CURSOR cFeriadoHO IS
		SELECT	DT_FERI,
				TP_FERI,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_FERIADO_HO;
	--
	BEGIN
	--
		FOR cAux IN cFeriadoHO LOOP
		--
			UPDATE	A8.TB_FERIADO_HO
			SET		TP_FERI				= cAux.TP_FERI,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	DT_FERI				= cAux.DT_FERI;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_FERIADO_HO
						(DT_FERI,
						TP_FERI,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.DT_FERI,
						cAux.TP_FERI,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_EMPRESA_FUSIONADA'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END ;
--
--*******************************************************************
--
PROCEDURE REPLICA_EMPRESA_HO
--
AS
--
	CURSOR cEmpresaHO IS
		SELECT	CO_EMPR,
				CO_EMPR_FUSI,
				CO_BANC,
				NU_CNPJ,
				NO_EMPR,
				NO_REDU_EMPR,
				SQ_ISPB,
				CO_GRUP_CPUL,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_EMPRESA_HO;
	--
	BEGIN
	--
		FOR cAux in cEmpresaHO loop
		--
			UPDATE	A8.TB_EMPRESA_HO
			SET		CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI,
					CO_BANC				= cAux.CO_BANC,
					NU_CNPJ				= cAux.NU_CNPJ,
					NO_EMPR				= cAux.NO_EMPR,
					NO_REDU_EMPR		= cAux.NO_REDU_EMPR,
					SQ_ISPB				= cAux.SQ_ISPB,
					CO_GRUP_CPUL		= cAux.CO_GRUP_CPUL ,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR				= cAux.CO_EMPR;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_EMPRESA_HO
						(CO_EMPR,
						CO_EMPR_FUSI,
						CO_BANC,
						NU_CNPJ,
						NO_EMPR,
						NO_REDU_EMPR,
						SQ_ISPB,
						CO_GRUP_CPUL,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ,
						ID_PART_CAMR_CETIP)
				VALUES	(cAux.CO_EMPR,
						cAux.CO_EMPR_FUSI,
						cAux.CO_BANC,
						cAux.NU_CNPJ,
						cAux.NO_EMPR,
						cAux.NO_REDU_EMPR,
						cAux.SQ_ISPB,
						cAux.CO_GRUP_CPUL,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE,
						NULL);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_EMPRESA_HO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END ;
--
--*******************************************************************
--
PROCEDURE REPLICA_LOCAL_LIQUIDACAO
--
AS
--
	CURSOR cLocalLiquidacao IS
		SELECT	CO_EMPR_FUSI,
				CO_LOCA_LIQU,
				SQ_ISPB,
				SG_LOCA_LIQU,
				DE_LOCA_LIQU,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_LOCAL_LIQUIDACAO;
	--
	BEGIN
	--
		FOR cAux IN cLocalLiquidacao LOOP
		--
			UPDATE	A8.TB_LOCAL_LIQUIDACAO
			SET		SQ_ISPB				= SQ_ISPB,
					SG_LOCA_LIQU		= cAux.SG_LOCA_LIQU,
					DE_LOCA_LIQU		= cAux.DE_LOCA_LIQU,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SIATEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI
			AND		CO_LOCA_LIQU		= cAux.CO_LOCA_LIQU;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_LOCAL_LIQUIDACAO
						(CO_EMPR_FUSI,
						CO_LOCA_LIQU,
						SQ_ISPB,
						SG_LOCA_LIQU,
						DE_LOCA_LIQU,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.CO_LOCA_LIQU,
						cAux.SQ_ISPB,
						cAux.SG_LOCA_LIQU,
						cAux.DE_LOCA_LIQU,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_LOCAL_LIQUIDACAO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_PRODUTO
--
AS
--
	CURSOR cProduto IS
		SELECT	CO_EMPR_FUSI,
				CO_PROD,
				SQ_ITEM_CAIX,
				DE_PROD,
				QT_DIAS_MAIR_VALO,
				VA_MINI_MAIR_VALO,
				QT_REGT_MAIR_VALO,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_PRODUTO;
	--
	BEGIN
	--
		FOR cAux IN cProduto LOOP
			UPDATE	A8.TB_PRODUTO
			SET		SQ_ITEM_CAIX		= cAux.SQ_ITEM_CAIX,
					DE_PROD				= cAux.DE_PROD,
					QT_DIAS_MAIR_VALO	= cAux.QT_DIAS_MAIR_VALO,
					VA_MINI_MAIR_VALO	= cAux.VA_MINI_MAIR_VALO,
					QT_REGT_MAIR_VALO	= cAux.QT_REGT_MAIR_VALO,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI
			AND		CO_PROD				= cAux.CO_PROD;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_PRODUTO
						(CO_EMPR_FUSI,
						CO_PROD,
						SQ_ITEM_CAIX,
						DE_PROD,
						QT_DIAS_MAIR_VALO,
						VA_MINI_MAIR_VALO,
						QT_REGT_MAIR_VALO,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.CO_PROD,
						cAux.SQ_ITEM_CAIX,
						cAux.DE_PROD,
						cAux.QT_DIAS_MAIR_VALO,
						cAux.VA_MINI_MAIR_VALO,
						cAux.QT_REGT_MAIR_VALO,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_PRODUTO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_EVENTO_FINANCEIRO
--
AS
--
	CURSOR cEventoFinanceiro IS
		SELECT	CO_EMPR_FUSI,
				CO_EVEN_FINC,
				DE_EVEN_FINC,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_EVENTO_FINANCEIRO;
	--
	BEGIN
	--
		FOR cAux IN cEventoFinanceiro LOOP
		--
			UPDATE	A8.TB_EVENTO_FINANCEIRO
			SET		DE_EVEN_FINC		= cAux.DE_EVEN_FINC,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI
			AND		CO_EVEN_FINC		= cAux.CO_EVEN_FINC;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_EVENTO_FINANCEIRO
						(CO_EMPR_FUSI,
						CO_EVEN_FINC,
						DE_EVEN_FINC,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.CO_EVEN_FINC,
						cAux.DE_EVEN_FINC,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_EVENTO_FINANCEIRO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_SEGMENTO
--
AS
--
	CURSOR cSegmento IS
		SELECT	CO_EMPR_FUSI,
				CO_SEGM,
				DE_SEGM,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM TB_SEGMENTO;
	--
	BEGIN
	--
		FOR cAux IN cSegmento LOOP
		--
			UPDATE	A8.TB_SEGMENTO
			SET		DE_SEGM				= cAux.DE_SEGM,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI
			AND		CO_SEGM				= cAux.CO_SEGM;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_SEGMENTO
						(CO_EMPR_FUSI,
						CO_SEGM,
						DE_SEGM,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.CO_SEGM,
						cAux.DE_SEGM,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_SEGMENTO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_TIPO_CONTA
--
AS
--
	CURSOR cTipoConta IS
		SELECT	CO_EMPR_FUSI,
				CO_TIPO_CNTA,
				DE_TIPO_CNTA,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_TIPO_CONTA;
	--
	BEGIN
	--
		FOR cAux IN cTipoConta LOOP
		--
			UPDATE	A8.TB_TIPO_CONTA
			SET		DE_TIPO_CNTA		= cAux.DE_TIPO_CNTA,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI
			AND		CO_TIPO_CNTA		= cAux.CO_TIPO_CNTA;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_TIPO_CONTA
						(CO_EMPR_FUSI,
						CO_TIPO_CNTA,
						DE_TIPO_CNTA,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.CO_TIPO_CNTA,
						cAux.DE_TIPO_CNTA,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_TIPO_CONTA'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_INDEXADOR
--
AS
--
	CURSOR cIndexador IS
		SELECT	CO_EMPR_FUSI,
				CO_INDX,
				DE_INDX,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_INDEXADOR;
	--
	BEGIN
	--
		FOR cAux IN cIndexador LOOP
		--
			UPDATE	A8.TB_INDEXADOR
			SET		DE_INDX				= cAux.DE_INDX,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	CO_EMPR_FUSI		= cAux.CO_EMPR_FUSI
			AND		CO_INDX				= cAux.CO_INDX;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_INDEXADOR
						(CO_EMPR_FUSI,
						CO_INDX,
						DE_INDX,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.CO_EMPR_FUSI,
						cAux.CO_INDX,
						cAux.DE_INDX,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_INDEXADOR'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_ERRO_BACEN
--
AS
--
	CURSOR cERRO_BACEN IS
		SELECT	CO_ERRO_BACEN,
				DE_ERRO_BACEN,
				DE_TRDZ_ERRO
		FROM	TB_ERRO_BACEN;
	--
	BEGIN
	--	
		DELETE A8.TB_ERRO_BACEN;

		FOR cAux IN cERRO_BACEN LOOP
		--
			INSERT	INTO A8.TB_ERRO_BACEN
					(CO_ERRO_BACEN,
					DE_ERRO_BACEN,
					DE_TRDZ_ERRO)
			VALUES	(cAux.CO_ERRO_BACEN,
					cAux.DE_ERRO_BACEN,
					cAux.DE_TRDZ_ERRO);
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_ERRO_BACEN'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_INSTITUICAO_SPB
--
AS
--
	CURSOR cInstituicaoSPB IS
		SELECT	SQ_ISPB,
				SQ_ISPB_SPER,
				CO_ISPB,
				NO_ISPB,
				NU_CNPJ,
				CO_CPEN,
				IN_TIPO_ISPB,
				IN_ENVI_CRIP,
				DT_INIC_VIGE,
				DT_FIM_VIGE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_INSTITUICAO_SPB;
	--
	BEGIN
	--
		FOR cAux in cInstituicaoSPB loop
		--
			UPDATE	A8.TB_INSTITUICAO_SPB
			SET		SQ_ISPB_SPER		= cAux.SQ_ISPB_SPER,
					CO_ISPB				= cAux.CO_ISPB,
					NO_ISPB				= cAux.NO_ISPB,
					NU_CNPJ				= cAux.NU_CNpj,
					CO_CPEN				= DECODE(cAux.CO_ISPB, 61411633, 0, cAux.CO_CPEN),
					IN_TIPO_ISPB		= cAux.IN_TIPO_ISPB,
					IN_ENVI_CRIP		= cAux.IN_ENVI_CRIP,
					DT_INIC_VIGE		= cAux.DT_INIC_VIGE,
					DT_FIM_VIGE			= cAux.DT_FIM_VIGE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	SQ_ISPB				= cAux.SQ_ISPB;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_INSTITUICAO_SPB
						(SQ_ISPB,
						SQ_ISPB_SPER,
						CO_ISPB,
						NO_ISPB,
						NU_CNPJ,
						CO_CPEN,
						IN_TIPO_ISPB,
						IN_ENVI_CRIP,
						DT_INIC_VIGE,
						DT_FIM_VIGE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.SQ_ISPB,
						cAux.SQ_ISPB_SPER,
						cAux.CO_ISPB,
						cAux.NO_ISPB,
						cAux.NU_CNPJ,
						DECODE(cAux.CO_ISPB, 61411633, 0, cAux.CO_CPEN),
						cAux.IN_TIPO_ISPB,
						cAux.IN_ENVI_CRIP,
						cAux.DT_INIC_VIGE,
						cAux.DT_FIM_VIGE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_INSTITUICAO_SPB'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_GRADE_HORARIO
--
AS
--
	CURSOR cGradeHorario IS
		SELECT	CO_GRAD_HORA,
				SQ_ISPB,
				IN_TIPO_GRAD,
				DT_EMIS_GRAD_BACEN,
				DT_INIC_VIGE_GRAD,
				DT_FIM_VIGE_GRAD,
				HO_ABER,
				HO_ENCE,
				DT_REFE
		FROM	TB_GRADE_HORARIO;
	--
	BEGIN
	--
		FOR cAux IN cGradeHorario LOOP
		--
			UPDATE	A8.TB_GRADE_HORARIO
			SET		IN_TIPO_GRAD		= cAux.IN_TIPO_GRAD,
					DT_EMIS_GRAD_BACEN	= cAux.DT_EMIS_GRAD_BACEN,
					DT_INIC_VIGE_GRAD	= cAux.DT_INIC_VIGE_GRAD,
					DT_FIM_VIGE_GRAD	= cAux.DT_FIM_VIGE_GRAD,
					HO_ABER				= cAux.HO_ABER,
					HO_ENCE				= cAux.HO_ENCE,
					DT_REFE				= cAux.DT_REFE
			WHERE	CO_GRAD_HORA		= cAux.CO_GRAD_HORA
			AND		SQ_ISPB				= cAux.SQ_ISPB
			AND		IN_TIPO_GRAD		= cAux.IN_TIPO_GRAD
			AND		DT_EMIS_GRAD_BACEN	= cAux.DT_EMIS_GRAD_BACEN;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_GRADE_HORARIO
						(CO_GRAD_HORA,
						SQ_ISPB,
						IN_TIPO_GRAD,
						DT_EMIS_GRAD_BACEN,
						DT_INIC_VIGE_GRAD,
						DT_FIM_VIGE_GRAD,
						HO_ABER,
						HO_ENCE,
						DT_REFE)
				VALUES	(cAux.CO_GRAD_HORA,
						cAux.SQ_ISPB,
						cAux.IN_TIPO_GRAD,
						cAux.DT_EMIS_GRAD_BACEN,
						cAux.DT_INIC_VIGE_GRAD,
						cAux.DT_FIM_VIGE_GRAD,
						cAux.HO_ABER,
						cAux.HO_ENCE,
						cAux.DT_REFE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_GRADE_HORARIO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_GRUPO
--
AS
--
	CURSOR cGrupo IS
		SELECT	SQ_GRUP,
				CO_GRUP,
				NO_GRUP,
				TX_GRUP,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_GRUPO;
	--
	BEGIN
	--
		FOR cAux IN cGrupo LOOP
		--
			UPDATE	A8.TB_GRUPO
			SET		CO_GRUP				= cAux.CO_GRUP,
					NO_GRUP				= cAux.NO_GRUP,
					TX_GRUP				= cAux.TX_GRUP,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	SQ_GRUP				= cAux.SQ_GRUP;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_GRUPO
						(SQ_GRUP,
						CO_GRUP,
						NO_GRUP,
						TX_GRUP,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.SQ_GRUP,
						cAux.CO_GRUP,
						cAux.NO_GRUP,
						cAux.TX_GRUP,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_GRUPO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_SERVICO
--
AS
--
	CURSOR cServico IS
		SELECT	SQ_SERV,
				SQ_GRUP,
				NO_SERV,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_SERVICO;
	--
	BEGIN
	--
		FOR cAux IN cServico LOOP
		--
			UPDATE	A8.TB_SERVICO
			SET		NO_SERV				= cAux.NO_SERV,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	SQ_SERV				= cAux.SQ_SERV;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_SERVICO
						(SQ_SERV,
						SQ_GRUP,
						NO_SERV,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.SQ_SERV,
						cAux.SQ_GRUP,
						cAux.NO_SERV,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_SERVICO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_EVENTO
--
AS
--
	CURSOR cEvento IS
		SELECT	SQ_EVEN,
				SQ_ISPB,
				SQ_SERV,
				SQ_TIPO_FLUX,
				CO_EVEN,
				NO_EVEN,
				DE_EVEN,
				TX_OBSE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_EVENTO;
	--
	BEGIN
	--
		FOR cAux IN cEvento LOOP
		--
			UPDATE	A8.TB_EVENTO
			SET		SQ_EVEN				= cAux.SQ_EVEN,
					SQ_ISPB				= cAux.SQ_ISPB,
					SQ_SERV				= cAux.SQ_SERV,
					SQ_TIPO_FLUX		= cAux.SQ_TIPO_FLUX,
					CO_EVEN				= cAux.CO_EVEN,
					NO_EVEN				= cAux.NO_EVEN,
					DE_EVEN				= cAux.DE_EVEN,
					TX_OBSE				= cAux.TX_OBSE,
					ID_USUA_ULTI_ATLZ	= 'SISTEMA',
					DH_ULTI_ATLZ		= SYSDATE
			WHERE	SQ_EVEN				= cAux.SQ_EVEN;
			--
			IF SQL%NOTFOUND THEN
				INSERT	INTO A8.TB_EVENTO
						(SQ_EVEN,
						SQ_ISPB,
						SQ_SERV,
						SQ_TIPO_FLUX,
						CO_EVEN,
						NO_EVEN,
						DE_EVEN,
						TX_OBSE,
						ID_USUA_ULTI_ATLZ,
						DH_ULTI_ATLZ)
				VALUES	(cAux.SQ_EVEN,
						cAux.SQ_ISPB,
						cAux.SQ_SERV,
						cAux.SQ_TIPO_FLUX,
						cAux.CO_EVEN,
						cAux.NO_EVEN,
						cAux.DE_EVEN,
						cAux.TX_OBSE,
						'SISTEMA',
						SYSDATE);
			END IF;
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_EVENTO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_MENSAGEM
--
AS
--
	CURSOR cMENSAGEM IS
		SELECT	SQ_MESG,
				SQ_EVEN,
				CO_MESG,
				NO_MESG,
				NO_TAG_PRIN_MESG,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_MENSAGEM;
	--
	BEGIN
	--
		DELETE A8.TB_MENSAGEM;

		FOR cAux IN cMENSAGEM LOOP
		--
			INSERT	INTO A8.TB_MENSAGEM
					(SQ_MESG,
					SQ_EVEN,
					CO_MESG,
					NO_MESG,
					NO_TAG_PRIN_MESG,
					ID_USUA_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			VALUES	(cAux.SQ_MESG,
					cAux.SQ_EVEN,
					cAux.CO_MESG,
					cAux.NO_MESG,
					cAux.NO_TAG_PRIN_MESG,
					'SISTEMA',
					SYSDATE);
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_MENSAGEM'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_TIPO_TAG
--
AS
--
	CURSOR cTIPO_TAG IS
		SELECT	SQ_TIPO_TAG,
				NO_TIPO_TAG,
				DE_TIPO_TAG,
				NU_TAMA_TAG,
				IN_TIPO_CTER,
				IN_TAG_SITU,
				QT_CASA_DECI,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_TIPO_TAG;
	--
	BEGIN
	--
		DELETE A8.TB_TIPO_TAG;

		FOR cAux IN cTIPO_TAG LOOP
		--
			INSERT	INTO A8.TB_TIPO_TAG
					(SQ_TIPO_TAG,
					NO_TIPO_TAG,
					DE_TIPO_TAG,
					NU_TAMA_TAG,
					IN_TIPO_CTER,
					IN_TAG_SITU,
					QT_CASA_DECI,
					ID_USUA_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			VALUES	(cAux.SQ_TIPO_TAG,
					cAux.NO_TIPO_TAG,
					cAux.DE_TIPO_TAG,
					cAux.NU_TAMA_TAG,
					cAux.IN_TIPO_CTER,
					cAux.IN_TAG_SITU,
					cAux.QT_CASA_DECI,
					'SISTEMA',
					SYSDATE);
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_TIPO_TAG'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_TAG
--
AS
--
	CURSOR cTAG IS
		SELECT	SQ_TAG,
				NO_TAG,
				DE_TAG,
				TX_DEFA,
				SQ_TIPO_TAG,
				IN_CATG_TAG,
				QT_REPE,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_TAG;
	--
	BEGIN
	--
		DELETE A8.TB_TAG;

		FOR cAux IN cTAG LOOP
		--
			INSERT	INTO A8.TB_TAG
					(SQ_TAG,
					NO_TAG,
					DE_TAG,
					TX_DEFA,
					SQ_TIPO_TAG,
					IN_CATG_TAG,
					QT_REPE,
					ID_USUA_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			VALUES	(cAux.SQ_TAG,
					cAux.NO_TAG,
					cAux.DE_TAG,
					cAux.TX_DEFA,
					cAux.SQ_TIPO_TAG,
					cAux.IN_CATG_TAG,
					cAux.QT_REPE,
					'SISTEMA',
					SYSDATE);
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_TAG'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_TAG_MENSAGEM
--
AS
--
	CURSOR cTAG_MENSAGEM IS
		SELECT	SQ_MESG,
				SQ_TAG,
				NU_ORDE_TAG,
				IN_OBRI,
				IN_NIVE_REPE
		FROM	TB_TAG_MENSAGEM;
	--
	BEGIN
	--
		DELETE	A8.TB_TAG_MENSAGEM;
		--
		FOR cAux IN cTAG_MENSAGEM LOOP
		--
			INSERT	INTO A8.TB_TAG_MENSAGEM
					(SQ_MESG,
					SQ_TAG,
					NU_ORDE_TAG,
					IN_OBRI,
					IN_NIVE_REPE)
			VALUES	(cAux.SQ_MESG,
					cAux.SQ_TAG,
					cAux.NU_ORDE_TAG,
					cAux.IN_OBRI,
					cAux.IN_NIVE_REPE);
		--
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_TAG_MENSAGEM'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_DOMINIO
--
AS
--
	CURSOR cDOMINIO IS
		SELECT	CO_DOMI,
				SQ_TIPO_TAG,
				DE_DOMI,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_DOMINIO;
	--
	BEGIN
	--
		DELETE A8.TB_DOMINIO;
		--
		FOR cAux IN cDOMINIO LOOP
		--
			INSERT	INTO A8.TB_DOMINIO
					(CO_DOMI,
					SQ_TIPO_TAG,
					DE_DOMI,
					ID_USUA_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			VALUES	(cAux.CO_DOMI,
					cAux.SQ_TIPO_TAG,
					cAux.DE_DOMI,
					'SISTEMA',
					SYSDATE);
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_DOMINIO'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_GRADE_MENSAGEM
--
AS
--
	CURSOR cGRADE_MENSAGEM IS
		SELECT	CO_GRAD_HORA,
				SQ_ISPB,
				SQ_MESG
		FROM	TB_GRADE_MENSAGEM;
	--
	BEGIN
	--
		DELETE	A8.TB_GRADE_MENSAGEM;
		--
		FOR cAux IN cGRADE_MENSAGEM LOOP
		--
			INSERT	INTO A8.TB_GRADE_MENSAGEM
					(CO_GRAD_HORA,
					SQ_ISPB,
					SQ_MESG)
			VALUES	(cAux.CO_GRAD_HORA,
					cAux.SQ_ISPB,
					cAux.SQ_MESG);
		--
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_GRADE_MENSAGEM'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
PROCEDURE REPLICA_SIST_MENSAGEM
--
AS
--
	CURSOR cSIST_MENSAGEM IS
		SELECT	SQ_MESG,
				CO_EMPR,
				SG_SIST,
				IN_FORM_MESG,
				ID_USUA_ULTI_ATLZ,
				DH_ULTI_ATLZ
		FROM	TB_SIST_MENSAGEM
		WHERE	SG_SIST = 'A8';
	--
	BEGIN
	--  
		DELETE	A8.TB_SIST_MENSAGEM;

		FOR cAux IN cSIST_MENSAGEM LOOP
			INSERT	INTO A8.TB_SIST_MENSAGEM
					(SQ_MESG,
					CO_EMPR,
					SG_SIST,
					IN_FORM_MESG,
					ID_USUA_ULTI_ATLZ,
					DH_ULTI_ATLZ)
			VALUES	(cAux.SQ_MESG,
					cAux.CO_EMPR,
					cAux.SG_SIST,
					cAux.IN_FORM_MESG,
					'SISTEMA',
					SYSDATE);
		END LOOP;
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002,'Erro na execução da Procedure A8K_REPLICACAO.REPLICA_SIST_MENSAGEM'
			|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 200));
--
END;
--
--*******************************************************************
--
END;
--
/
--
SHOW ERRORS
