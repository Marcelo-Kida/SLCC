--
--Criaçao de Trigger do TQ   : A8PROC.TGBR_ATLZ_OPER_ATIV
--Para convivência com o TQ (liquidacao CETIP)
--
CREATE OR REPLACE TRIGGER TGBR_ATLZ_OPER_ATIV
	BEFORE UPDATE OF CO_ULTI_SITU_PROC
	ON A8.TB_OPER_ATIV
	FOR EACH ROW
--
DECLARE
--
PROCEDURE P_HISTORICO IS
	--
	BEGIN
		--
		INSERT	INTO A8.TB_HIST_SITU_ACAO_OPER_ATIV
				(NU_SEQU_OPER_ATIV,
				DH_SITU_ACAO_OPER_ATIV,
				CO_SITU_PROC,
				TP_ACAO_OPER_ATIV,
				TP_JUST_SITU_PROC,
				TX_CNTD_ANTE_ACAO,
				CO_USUA_ATLZ,
				CO_ETCA_USUA_ATLZ)
		VALUES	(:NEW.NU_SEQU_OPER_ATIV,
				SYSDATE,
				:NEW.CO_ULTI_SITU_PROC,
				NULL,
				NULL,
				'Sistema TQ',
				'SIST TQ',
				'SERVIDOR');
	END;
	--
	BEGIN
		--
		IF :NEW.CO_ULTI_SITU_PROC > 1000 THEN 
			-- Quanto o campo for maior que 1000, é o TQ que está alterando o Status
			-- Subtrai 1000 para ficar com o status correto
			:NEW.CO_ULTI_SITU_PROC := :NEW.CO_ULTI_SITU_PROC - 1000;
			--
			:NEW.DH_ULTI_ATLZ := SYSDATE;
			--
			:NEW.CO_ETCA_TRAB_ULTI_ATLZ := 'SERVIDOR';
			--
			:NEW.CO_USUA_ULTI_ATLZ := 'SIST TQ';
			--
			-- grava histórico de status
			P_HISTORICO;
		END IF;
		--
		EXCEPTION
			WHEN OTHERS THEN
				null;
	END;
/

