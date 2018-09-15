CREATE OR REPLACE PACKAGE A7K_INTEGRA_ARQUIVO_CONTABIL		IS

	PROCEDURE	GRAVAR (pDiretorio		IN	VARCHAR2,
				pArquivo		IN	VARCHAR2);

	/* ------------------------------------------------------------------------------------------------
		Premissas:	É necessário ter acesso de leitura e gravação na tabela
							A8.TB_HIST_ENVI_INFO_CNTB.

				O tamanho do registro contábil a ser gerado em cada linha do arquivo
					terá 380 posições, conforme definição da constante
					<< TAM_REGISTRO_CONTABIL >>. Dessa forma, para qualquer alteração
					de layout, será necessário alterar também a constante.
	   ------------------------------------------------------------------------------------------------ */
END;
/

CREATE OR REPLACE PACKAGE BODY A7K_INTEGRA_ARQUIVO_CONTABIL	IS

	-- Declara a estrutura do REGISTRO
	TYPE rec_HA IS RECORD (
		EMPRESA				CHAR(4),
		CLAVE_DE_INTERFASE		CHAR(3)		:= 'A8A',
		FECHA_CONTABLE			CHAR(8),
		FECHA_DE_OPERACION		CHAR(8),
		PRODUCTO			CHAR(2)		:= LPAD('0', 2, '0'),
		SUBPRODUCTO			CHAR(4)		:= '',
		GARANTIA			CHAR(3)		:= LPAD('0', 3, '0'),
		TIPO_DE_PLAZO			CHAR(1)		:= '',
		PLAZO				CHAR(3)		:= LPAD('0', 3, '0'),
		SUBSECTOR			CHAR(1)		:= LPAD('0', 1, '0'),
		SECTOR_B_E			CHAR(2)		:= LPAD('0', 2, '0'),
		CNAE				CHAR(5)		:= LPAD('0', 5, '0'),
		EMPRESA_TUTELADA		CHAR(4)		:= '',
		AMBITO				CHAR(2)		:= LPAD('0', 2, '0'),
		MOROSIDAD			CHAR(1)		:= '',
		INVERSION			CHAR(1)		:= '',
		OPERACION			CHAR(3)		:= LPAD('0', 3, '0'),
		CODIGO_CONTABLE			CHAR(5),
		DIVISA				CHAR(3)		:= '',
		TIPO_DE_DIVISA			CHAR(1)		:= '',
		TIPO_NOMINAL			CHAR(5)		:= LPAD('0', 5, '0'),
		Filler1				CHAR(5)		:= '',
		VARIOS				CHAR(30)	:= RPAD('PZ', 30, ' '),
		CLAVE_DE_AUTORIZACION		CHAR(6)		:= '',
		CENTRO_OPERANTE			CHAR(4),
		CENTRO_ORIGEN			CHAR(4),
		CENTRO_DESTINO			CHAR(4),
		NUM_MOVTOS_AL_DEBE		CHAR(7),
		NUM_MOVTOS_AL_HABER		CHAR(7),
		IMPORTE_DEBE_EN_PESETAS		CHAR(15),
		IMPORTE_HABER_EN_PESETAS	CHAR(15),
		IMPORTE_DEBE_EN_DIVISA		CHAR(15)	:= LPAD('0', 15, '0'),
		IMPORTE_HABER_EN_DIVISA		CHAR(15)	:= LPAD('0', 15, '0'),
		INDICADOR_DE_CORRECCION		CHAR(1)		:= '',
		NUMERO_DE_CONTROL		CHAR(12)	:= '',
		CLAVE_DE_CONCEPTO		CHAR(3),
		DESCRIPCION_DE_CONCEPTO		CHAR(14),
		TIPODE_CONCEPTO			CHAR(1)		:= '',
		OBSERVACIONES			CHAR(30)	:= '',
		SANCTCCC			CHAR(18)	:= '',
		APLICACION_ORIGEN		CHAR(3)		:= '',
		APLICACION_DESTINO		CHAR(3)		:= '',
		OBSERVACIONES3			CHAR(6)		:= '',
		RESERVAT			CHAR(4)		:= '',
		HACTRGEN			CHAR(4)		:= '',
		HAYCOCAI			CHAR(1)		:= '',
		HAYCTORD			CHAR(1)		:= '',
		SATINTER			CHAR(5)		:= LPAD('0', 5, '0'),
		SACCLVOP			CHAR(3)		:= '',
		SACCEGES			CHAR(4)		:= '',
		SACAPLCP			CHAR(2)		:= '',
		SACCDTGT			CHAR(2)		:= '',
		SAYUTILI			CHAR(1)		:= '',
		SAYROTAC			CHAR(2)		:= '',
		FALTPART			CHAR(8)		:= '',
		OBSERV4				CHAR(30),
		NIO				CHAR(24)	:= '',
		Filler2				CHAR(2)		:= ''
	);

	PROCEDURE	GRAVARHISTORICOEXECUCAO (pRotina		IN	number,
						 pStatus		IN	number,
						 pErro			IN	varchar2)	IS
	BEGIN

		INSERT INTO A7.TB_HIST_EXEC_ROTI_BATCH(
			CO_ROTI_BATCH,
			DH_FIM_EXEC,
			IN_EXEC_SUCE,
			DE_ERRO_EXEC)
		VALUES (pRotina,
			SYSDATE,
			pStatus,
			SUBSTR(pErro, 1, 200));

	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(SQLCODE, 'ERRO NA ROTINA DE GRAVACAO DE HISTORICO DE EXECUCAO: ' || SQLERRM);

	END GRAVARHISTORICOEXECUCAO;

	FUNCTION MONTAR_LINHA_ARQUIVO (
		pEMPRESA			NUMBER,
		pFECHA_CONTABLE			NUMBER,
		pFECHA_DE_OPERACION		NUMBER,
		pCODIGO_CONTABLE		NUMBER,
		pCENTRO_OPERANTE		NUMBER,
		pCENTRO_ORIGEN			NUMBER,
		pCENTRO_DESTINO			NUMBER,
		pNUM_MOVTOS_AL_DEBE		NUMBER,
		pNUM_MOVTOS_AL_HABER		NUMBER,
		pIMPORTE_DEBE_EN_PESETAS	NUMBER,
		pIMPORTE_HABER_EN_PESETAS	NUMBER,
		pCLAVE_DE_CONCEPTO		VARCHAR2,
		pDESCRIPCION_DE_CONCEPTO	VARCHAR2,
		pOBSERV4			VARCHAR2) RETURN VARCHAR2 IS

		-- Declara a variável para a execução contendo a estrutura do REGISTRO
		HA	rec_HA;
		v_linha	VARCHAR2(380);

	BEGIN
	/*
		Os campos exibidos com comentário << -- >> são apenas para
		ilustrar a ordem da estrutura e, não necessitam de preenchimento,
		pois a declaração << TYPE RECORD >> já os impôs o DEFAULT.
	*/	
		
  	        HA.EMPRESA                      := LPAD(33, 4, '0');    
		HA.FECHA_CONTABLE		:= LPAD(pFECHA_CONTABLE, 8, '0');
		HA.FECHA_DE_OPERACION		:= LPAD(pFECHA_DE_OPERACION, 8, '0');
		HA.CODIGO_CONTABLE		:= LPAD(pCODIGO_CONTABLE, 5, '0');
		HA.CENTRO_OPERANTE		:= LPAD(pCENTRO_OPERANTE, 4, '0');
		HA.CENTRO_ORIGEN		:= LPAD(pCENTRO_ORIGEN, 4, '0');
		HA.CENTRO_DESTINO		:= LPAD(pCENTRO_DESTINO, 4, '0');
		HA.NUM_MOVTOS_AL_DEBE		:= LPAD(pNUM_MOVTOS_AL_DEBE, 7, '0');
		HA.NUM_MOVTOS_AL_HABER		:= LPAD(pNUM_MOVTOS_AL_HABER, 7, '0');
		HA.IMPORTE_DEBE_EN_PESETAS	:= LPAD(pIMPORTE_DEBE_EN_PESETAS, 15, '0');
		HA.IMPORTE_HABER_EN_PESETAS	:= LPAD(pIMPORTE_HABER_EN_PESETAS, 15, '0');
		HA.CLAVE_DE_CONCEPTO		:= RPAD(pCLAVE_DE_CONCEPTO, 3);
		HA.DESCRIPCION_DE_CONCEPTO	:= RPAD(pDESCRIPCION_DE_CONCEPTO, 14);
		HA.OBSERV4			:= RPAD(RPAD(pDESCRIPCION_DE_CONCEPTO, 17) || 
							RPAD(SUBSTR(pOBSERV4, 2, 5), 6)    || 
							'A8'                               ||
							SUBSTR(pOBSERV4, 1, 1), 30);

		v_linha :=	HA.EMPRESA			||
				HA.CLAVE_DE_INTERFASE		||
				HA.FECHA_CONTABLE		||
				HA.FECHA_DE_OPERACION		||
				HA.PRODUCTO			||
				HA.SUBPRODUCTO			||
				HA.GARANTIA			||
				HA.TIPO_DE_PLAZO		||
				HA.PLAZO			||
				HA.SUBSECTOR			||
				HA.SECTOR_B_E			||
				HA.CNAE				||
				HA.EMPRESA_TUTELADA		||
				HA.AMBITO			||
				HA.MOROSIDAD			||
				HA.INVERSION			||
				HA.OPERACION			||
				HA.CODIGO_CONTABLE		||
				HA.DIVISA			||
				HA.TIPO_DE_DIVISA		||
				HA.TIPO_NOMINAL			||
				HA.Filler1			||
				HA.VARIOS			||
				HA.CLAVE_DE_AUTORIZACION	||
				HA.CENTRO_OPERANTE		||
				HA.CENTRO_ORIGEN		||
				HA.CENTRO_DESTINO		||
				HA.NUM_MOVTOS_AL_DEBE		||
				HA.NUM_MOVTOS_AL_HABER		||
				HA.IMPORTE_DEBE_EN_PESETAS	||
				HA.IMPORTE_HABER_EN_PESETAS	||
				HA.IMPORTE_DEBE_EN_DIVISA	||
				HA.IMPORTE_HABER_EN_DIVISA	||
				HA.INDICADOR_DE_CORRECCION	||
				HA.NUMERO_DE_CONTROL		||
				HA.CLAVE_DE_CONCEPTO		||
				HA.DESCRIPCION_DE_CONCEPTO	||
				HA.TIPODE_CONCEPTO		||
				HA.OBSERVACIONES		||
				HA.SANCTCCC			||
				HA.APLICACION_ORIGEN		||
				HA.APLICACION_DESTINO		||
				HA.OBSERVACIONES3		||
				HA.RESERVAT			||
				HA.HACTRGEN			||
				HA.HAYCOCAI			||
				HA.HAYCTORD			||
				HA.SATINTER			||
				HA.SACCLVOP			||
				HA.SACCEGES			||
				HA.SACAPLCP			||
				HA.SACCDTGT			||
				HA.SAYUTILI			||
				HA.SAYROTAC			||
				HA.FALTPART			||
				HA.OBSERV4			||
				HA.NIO				||
				HA.Filler2;

		RETURN v_linha;

	EXCEPTION
		WHEN OTHERS THEN
			RETURN '';
	END;

	FUNCTION CAPTURAR_CENARIO_CONTABIL (
		pTP_BKOF			A8.TB_PARM_HIST_CNTA_CNTB.TP_BKOF%TYPE,
		pCO_EMPR			A8.TB_PARM_HIST_CNTA_CNTB.CO_EMPR%TYPE,
		pCO_LOCA_LIQU			A8.TB_PARM_HIST_CNTA_CNTB.CO_LOCA_LIQU%TYPE,
		pIN_LANC_DEBT_CRED		A8.TB_PARM_HIST_CNTA_CNTB.IN_LANC_DEBT_CRED%TYPE,
		pCO_CNTA_DEBT		OUT	A8.TB_PARM_HIST_CNTA_CNTB.CO_CNTA_DEBT%TYPE,
		pCO_CNTA_CRED		OUT	A8.TB_PARM_HIST_CNTA_CNTB.CO_CNTA_CRED%TYPE,
		pCO_HIST_CNTA_CNTB	OUT	A8.TB_PARM_HIST_CNTA_CNTB.CO_HIST_CNTA_CNTB%TYPE,
		pDE_HIST_CNTA_CNTB	OUT	A8.TB_PARM_HIST_CNTA_CNTB.DE_HIST_CNTA_CNTB%TYPE,
		pCO_CENT_DEST		OUT	A8.TB_PARM_HIST_CNTA_CNTB.CO_CENT_DEST%TYPE) RETURN BOOLEAN IS
	BEGIN
		SELECT	CO_CNTA_DEBT,
			CO_CNTA_CRED,
			CO_HIST_CNTA_CNTB,
			DE_HIST_CNTA_CNTB,
			CO_CENT_DEST
		INTO	pCO_CNTA_DEBT,
			pCO_CNTA_CRED,
			pCO_HIST_CNTA_CNTB,
			pDE_HIST_CNTA_CNTB,
			pCO_CENT_DEST
		FROM	A8.TB_PARM_HIST_CNTA_CNTB
		WHERE	TP_BKOF			= pTP_BKOF
		AND	CO_EMPR			= pCO_EMPR
		AND	SG_SIST			= 'PZ'
		AND	CO_LOCA_LIQU	        = pCO_LOCA_LIQU
		AND	IN_LANC_DEBT_CRED	= pIN_LANC_DEBT_CRED;

		RETURN TRUE;

	EXCEPTION
		WHEN OTHERS THEN
			RETURN FALSE;
	END;

	PROCEDURE	GRAVAR (pDiretorio		IN	VARCHAR2,
				pArquivo		IN	VARCHAR2)	IS

		TAM_REGISTRO_CONTABIL		CONSTANT	NUMBER := 380;
		CENTRO_OPERANTE			CONSTANT	NUMBER := 6544;
		CENTRO_ORIGEM			CONSTANT	NUMBER := 6544;

		CURSOR cTabela IS
			/* -------------------------------------------------------------
				Alterar o SELECT abaixo de acordo com a necessidade,
			   	mantendo a descrição do ALIAS como << CONTEUDO >>.
			   ------------------------------------------------------------- */
			SELECT	SUBSTR(A.TX_ITGR_CNTB, 1, TAM_REGISTRO_CONTABIL)	CONTEUDO,
				A.SG_SIST,
				A.NU_SEQU_OPER_ATIV,
				B.DT_OPER_ATIV,
				A.CO_EMPR,
				A.TP_LANC_ITGR,
				A.DH_ENVI_CNTB,
				A.IN_LANC_DEBT_CRED,
				C.TP_BKOF
			FROM	A8.TB_HIST_ENVI_INFO_CNTB	A,
				A8.TB_OPER_ATIV			B,
				A8.TB_VEIC_LEGA			C
				-- Liga com a tabela de operações para a capturar o valor para NET STR
				-- ponta contábil com o PZ
			WHERE	A.NU_SEQU_OPER_ATIV     = B.NU_SEQU_OPER_ATIV
			AND	B.CO_VEIC_LEGA		= C.CO_VEIC_LEGA
			AND	B.SG_SIST		= C.SG_SIST
			AND	UPPER(TRIM(A.SG_SIST))	IN ('HA', 'PZ')
			AND	A.IN_ITGR_CNTB		= 0	-- Não Integrado
			ORDER BY
				A.DH_ENVI_CNTB;
			/* ------------------------------------------------------------- */

		CURSOR cTabelaMesg IS
			/* -------------------------------------------------------------
				-- Captura apenas as mensagens de pagamento, pois as mensagens de recebimento
				-- são contabilizadas diretamente pelo PZ
			   ------------------------------------------------------------- */
			/* -------------------------------------------------------------
				-- Mensagens STR / PAG do MV s/ conta corrente
			   ------------------------------------------------------------- */
			SELECT  DISTINCT A.CO_EMPR, A.TP_BKOF, A.CO_LOCA_LIQU, A.NU_COMD_OPER, B.VA_FINC
			FROM    A8.TB_MESG_RECB_ENVI_SPB A,
			        A8.TB_MESG_RECB_SPB_CNCL B,
			        A8.TB_HIST_ENVI_INFO_CNTB C, 
			        A8.TB_OPER_ATIV D,
			        A8.TB_TIPO_OPER E
			WHERE   A.NU_CTRL_IF			= B.NU_CTRL_IF
			AND     A.DH_REGT_MESG_SPB		= B.DH_REGT_MESG_SPB
			AND     A.NU_SEQU_CNTR_REPE		= B.NU_SEQU_CNTR_REPE
			AND     A.NU_SEQU_OPER_ATIV	   NOT IN (SELECT NU_SEQU_OPER_ATIV
			                                           FROM   A8.TB_HIST_ENVI_INFO_CNTB)
			AND     A.NU_SEQU_OPER_ATIV             = D.NU_SEQU_OPER_ATIV
			AND     D.TP_OPER                       = E.TP_OPER 
			AND     E.TP_MESG_RECB_INTE             = 150
			AND     A.CO_MESG_SPB			IN ('STR0007',
			                        		    'PAG0106')
			AND     A.CO_ULTI_SITU_PROC	        = 71
			AND     TRUNC(A.DH_REGT_MESG_SPB)	= TRUNC(SYSDATE)
			AND     TRUNC(C.DH_ENVI_CNTB)           = TRUNC(SYSDATE)

			UNION ALL

			/* -------------------------------------------------------------
				-- Mensagens STR / PAG do MV c/ conta corrente
			   ------------------------------------------------------------- */
			SELECT  DISTINCT A.CO_EMPR, A.TP_BKOF, A.CO_LOCA_LIQU, A.NU_COMD_OPER, B.VA_FINC
			FROM    A8.TB_MESG_RECB_ENVI_SPB A,
			        A8.TB_MESG_RECB_SPB_CNCL B,
			        A8.TB_HIST_ENVI_INFO_CNTB C, 
			        A8.TB_OPER_ATIV D,
			        A8.TB_TIPO_OPER E
			WHERE   A.NU_CTRL_IF			= B.NU_CTRL_IF
			AND     A.DH_REGT_MESG_SPB		= B.DH_REGT_MESG_SPB
			AND     A.NU_SEQU_CNTR_REPE		= B.NU_SEQU_CNTR_REPE
			AND     A.NU_SEQU_OPER_ATIV		= C.NU_SEQU_OPER_ATIV
			AND     A.NU_SEQU_OPER_ATIV             = D.NU_SEQU_OPER_ATIV
			AND     D.TP_OPER                       = E.TP_OPER 
			AND     E.TP_MESG_RECB_INTE             = 150
			AND     A.CO_MESG_SPB			IN ('STR0006',
								    'STR0007',
								    'STR0008',
								    'STR0009',
								    'STR0025',
								    'STR0034',
								    'PAG0105',
								    'PAG0106',
								    'PAG0108',
								    'PAG0109',
								    'PAG0121',
								    'PAG0134')
			AND     A.CO_ULTI_SITU_PROC		= 71
			AND     TRUNC(A.DH_REGT_MESG_SPB)	= TRUNC(SYSDATE)

			UNION ALL

			/* -------------------------------------------------------------
				-- Mensagens Despesa BMC
			   ------------------------------------------------------------- */
			SELECT  DISTINCT A.CO_EMPR, A.TP_BKOF, A.CO_LOCA_LIQU, A.NU_COMD_OPER, B.VA_FINC
			FROM    A8.TB_MESG_RECB_ENVI_SPB A,
			        A8.TB_MESG_RECB_SPB_CNCL B,
			        A8.TB_HIST_ENVI_INFO_CNTB C, 
			        A8.TB_OPER_ATIV D,
			        A8.TB_TIPO_OPER E
			WHERE   A.NU_CTRL_IF			= B.NU_CTRL_IF
			AND     A.DH_REGT_MESG_SPB		= B.DH_REGT_MESG_SPB
			AND     A.NU_SEQU_CNTR_REPE		= B.NU_SEQU_CNTR_REPE
			AND     A.NU_SEQU_OPER_ATIV		NOT IN C.NU_SEQU_OPER_ATIV
			AND     A.NU_SEQU_OPER_ATIV             = D.NU_SEQU_OPER_ATIV
			AND     D.TP_OPER                       = E.TP_OPER 
			AND     E.TP_MESG_RECB_INTE             = 136
			AND     A.CO_MESG_SPB			IN ('STR0007')
			AND     A.CO_ULTI_SITU_PROC		= 71
			AND     TRUNC(A.DH_REGT_MESG_SPB)	= TRUNC(SYSDATE)

			UNION ALL

			/* -------------------------------------------------------------
				-- Net das mensagens de corretoras
			   ------------------------------------------------------------- */
			SELECT  A.CO_EMPR, A.TP_BKOF, A.CO_LOCA_LIQU, A.NU_COMD_OPER, SUM(B.VA_FINC) AS VA_FINC
			FROM    A8.TB_MESG_RECB_ENVI_SPB A,
			        A8.TB_MESG_RECB_SPB_CNCL B,
			        (SELECT DISTINCT(NU_CTRL_IF) AS NU_CTRL_IF FROM  A8.TB_CNCL_OPER_ATIV) C, 
			        (SELECT DISTINCT TP_OPER , NU_SEQU_CNCL_OPER_ATIV_MESG FROM  A8.TB_OPER_ATIV) D,
			        A8.TB_TIPO_OPER E
			WHERE   A.NU_CTRL_IF			= B.NU_CTRL_IF
			AND     A.DH_REGT_MESG_SPB		= B.DH_REGT_MESG_SPB
			AND     A.NU_SEQU_CNTR_REPE		= B.NU_SEQU_CNTR_REPE
			AND     A.NU_CTRL_IF			= C.NU_CTRL_IF
			AND     A.NU_SEQU_CNCL_OPER_ATIV_MESG   = D.NU_SEQU_CNCL_OPER_ATIV_MESG
			AND     D.TP_OPER                       = E.TP_OPER 
			AND     E.TP_MESG_RECB_INTE             = 50
			AND     A.CO_MESG_SPB			IN ('STR0004',
								    'STR0007')
			AND     A.CO_ULTI_SITU_PROC		= 71
			AND     TRUNC(A.DH_REGT_MESG_SPB)	= TRUNC(SYSDATE)
			GROUP BY A.CO_EMPR,
			         A.TP_BKOF,
			         A.CO_LOCA_LIQU,
			         A.NU_COMD_OPER;
			/* ------------------------------------------------------------- */
        

        v_linha			varchar2(380);	-- Tamanho da linha para registro contábil
		v_Cont			NUMBER;
		v_NET_STR		NUMBER    := 0;
		v_OUTPUT_FILE		UTL_FILE.FILE_TYPE;
		v_Valor			NUMBER;
		v_CONTA_CONTABIL	NUMBER;
		v_IndDebito		NUMBER;
		v_IndCredito		NUMBER;
		v_ValorDebito		NUMBER;
		v_ValorCredito		NUMBER;
		v_CO_CNTA_DEBT		A8.TB_PARM_HIST_CNTA_CNTB.CO_CNTA_DEBT%TYPE;
		v_CO_CNTA_CRED		A8.TB_PARM_HIST_CNTA_CNTB.CO_CNTA_CRED%TYPE;
		v_CO_HIST_CNTA_CNTB	A8.TB_PARM_HIST_CNTA_CNTB.CO_HIST_CNTA_CNTB%TYPE;
		v_DE_HIST_CNTA_CNTB	A8.TB_PARM_HIST_CNTA_CNTB.DE_HIST_CNTA_CNTB%TYPE;
		v_CO_CENT_DEST		A8.TB_PARM_HIST_CNTA_CNTB.CO_CENT_DEST%TYPE;

		PARAM_NAO_DEFINIDO	EXCEPTION;
		SELECAO_VAZIA		EXCEPTION;
		HIST_NAO_CADASTRADO	EXCEPTION;
		REGISTRO_PZ_INVALIDO	EXCEPTION;

		BEGIN
		IF pDiretorio IS NULL OR pArquivo IS NULL THEN
			RAISE PARAM_NAO_DEFINIDO;
		END IF;

		-- Verifica se existem registros a serem contabilizados
		SELECT	count(*)
		INTO v_Cont
		FROM	A8.TB_HIST_ENVI_INFO_CNTB
		WHERE   UPPER(TRIM(SG_SIST))	IN ('HA', 'PZ')
		AND	IN_ITGR_CNTB		= 0;

		-- Cria e abre o arquivo para a integração contábil
		v_OUTPUT_FILE	:= UTL_FILE.FOPEN(pDiretorio, pArquivo, 'w');

		-- Inicializa a variável de controle para a tabela temporária
		--  de mensagens PZ (NET de Pagamento) e a própria tabela
		IF v_Cont != 0 THEN
			FOR cAux IN cTabela LOOP
				v_linha := cAux.CONTEUDO;

				UTL_FILE.PUT_LINE(v_OUTPUT_FILE, v_linha);
				UTL_FILE.FFLUSH(v_OUTPUT_FILE);

				UPDATE	A8.TB_HIST_ENVI_INFO_CNTB
				SET	IN_ITGR_CNTB		= 1	-- Integrado
				WHERE	SG_SIST			= cAux.SG_SIST
				AND	NU_SEQU_OPER_ATIV	= cAux.NU_SEQU_OPER_ATIV
				AND	CO_EMPR			= cAux.CO_EMPR
				AND	TP_LANC_ITGR		= cAux.TP_LANC_ITGR
				AND	DH_ENVI_CNTB		= cAux.DH_ENVI_CNTB
				AND	IN_LANC_DEBT_CRED	= cAux.IN_LANC_DEBT_CRED;
			END LOOP;
		END IF;

		FOR cAux IN cTabelaMesg LOOP

			v_CO_CNTA_DEBT		:= 0;
			v_CO_CNTA_CRED		:= 0;
			v_CO_HIST_CNTA_CNTB	:= '';
			v_DE_HIST_CNTA_CNTB	:= '';
			v_CO_CENT_DEST		:= 0;

			-- Monta o lançamento para o HA a partir do NET obtido
			v_Valor := REPLACE(TO_CHAR(cAux.VA_FINC, '0000000000000.00'), '.', '');

			-- Efetua sempre um LOOPING com 2 iterações, pois a 1ª é para o registro de DÉBITO e
			-- a 2ª é para o registro de CRÉDITO
			FOR cLanc IN 1 .. 2 LOOP
				IF NOT CAPTURAR_CENARIO_CONTABIL(
					cAux.TP_BKOF,
					cAux.CO_EMPR,
					cAux.CO_LOCA_LIQU,
					cLanc,
					v_CO_CNTA_DEBT,
					v_CO_CNTA_CRED,
					v_CO_HIST_CNTA_CNTB,
					v_DE_HIST_CNTA_CNTB,
					v_CO_CENT_DEST) THEN

					RAISE HIST_NAO_CADASTRADO;

				END IF;

				v_CONTA_CONTABIL	:= v_CO_CNTA_DEBT;

				IF cLanc = 1 THEN
					v_IndDebito		:= 1;
					v_IndCredito		:= 0;
					v_ValorDebito		:= v_Valor;
					v_ValorCredito		:= 0;
				ELSE
					v_IndDebito		:= 0;
					v_IndCredito		:= 1;
					v_ValorDebito		:= 0;
					v_ValorCredito		:= v_Valor;
				END IF;

				v_linha := MONTAR_LINHA_ARQUIVO(
						cAux.CO_EMPR,
						TO_CHAR(SYSDATE, 'YYYYMMDD'),
						TO_CHAR(SYSDATE, 'YYYYMMDD'),
						v_CONTA_CONTABIL,
						CENTRO_OPERANTE,
						CENTRO_ORIGEM,
						v_CO_CENT_DEST,
						v_IndDebito,
						v_IndCredito,
						v_ValorDebito,
						v_ValorCredito,
						v_CO_HIST_CNTA_CNTB,
						v_DE_HIST_CNTA_CNTB,
						cAux.NU_COMD_OPER);

				IF LENGTH(TRIM(v_linha)) = 0 THEN
					RAISE REGISTRO_PZ_INVALIDO;
				END IF;

				UTL_FILE.PUT_LINE(v_OUTPUT_FILE, v_linha);
				UTL_FILE.FFLUSH(v_OUTPUT_FILE);
			END LOOP;
		END LOOP;

		UTL_FILE.FCLOSE_ALL;

		COMMIT;

		GRAVARHISTORICOEXECUCAO (2,1,null);

		COMMIT;

	EXCEPTION
		WHEN UTL_FILE.invalid_path THEN
			ROLLBACK;
			UTL_FILE.FCLOSE_ALL;

			GRAVARHISTORICOEXECUCAO (2,2,'PATH INVÁLIDO');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20001, 'PATH INVÁLIDO');

		WHEN UTL_FILE.invalid_mode THEN
			ROLLBACK;
			UTL_FILE.FCLOSE_ALL;

			GRAVARHISTORICOEXECUCAO (2,2,'MODO DE ABERTURA INVÁLIDO');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20002, 'MODO DE ABERTURA INVÁLIDO');

		WHEN UTL_FILE.invalid_operation THEN
			ROLLBACK;
			UTL_FILE.FCLOSE_ALL;

			GRAVARHISTORICOEXECUCAO (2,2,'OPERAÇÃO INVÁLIDA');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20003, 'OPERAÇÃO INVÁLIDA');

		WHEN PARAM_NAO_DEFINIDO THEN
			ROLLBACK;

			GRAVARHISTORICOEXECUCAO (2,2,'PARÂMETRO DE SISTEMA NÃO DEFINIDO OU ACESSO NÃO PERMITIDO À TABELA DE SISTEMA');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20004, 'PARÂMETRO DE SISTEMA NÃO DEFINIDO OU ACESSO NÃO PERMITIDO À TABELA DE SISTEMA');

		WHEN SELECAO_VAZIA THEN

			GRAVARHISTORICOEXECUCAO (2,2,'ROTINA EXPORTAÇÃO ARQUIVO CONTÁBIL JÁ EXECUTADA OU NÃO EXISTEM REGISTROS');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20005, 'ROTINA EXPORTAÇÃO ARQUIVO CONTÁBIL JÁ EXECUTADA OU NÃO EXISTEM REGISTROS');

		WHEN HIST_NAO_CADASTRADO THEN
			ROLLBACK;

			GRAVARHISTORICOEXECUCAO (2,2,'CONTAS CONTÁBEIS, CENTRO DESTINO E HISTÓRICO NÃO CADASTRADOS PARA A INTERGRAÇÃO PZ');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20006, 'CONTAS CONTÁBEIS, CENTRO DESTINO E HISTÓRICO NÃO CADASTRADOS PARA A INTERGRAÇÃO PZ');

		WHEN REGISTRO_PZ_INVALIDO THEN
			ROLLBACK;

			GRAVARHISTORICOEXECUCAO (2,2,'ERRO AO MONTAR O REGISTRO CONTÁBIL DE INTERGRAÇÃO PZ');
			COMMIT;

			RAISE_APPLICATION_ERROR(-20007, 'ERRO AO MONTAR O REGISTRO CONTÁBIL DE INTERGRAÇÃO PZ');

		WHEN OTHERS THEN
			ROLLBACK;
			UTL_FILE.FCLOSE_ALL;


			GRAVARHISTORICOEXECUCAO (2, 2, SUBSTR(SQLERRM, 1, 200));
			COMMIT;

			RAISE_APPLICATION_ERROR(SQLCODE, 'ERRO NÃO TRATADO: ' || SQLERRM);
	END GRAVAR;

END;
/
