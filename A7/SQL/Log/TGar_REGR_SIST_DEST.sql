--
--Cria�ao de Trigger de Log  : A7PROC.TGAR_REGR_SIST_DEST
--Tabela associada           : TB_REGR_SIST_DEST
--Data Cria�ao               : 29-08-2003 12:55:39
--
CREATE OR REPLACE TRIGGER TGAR_REGR_SIST_DEST
	AFTER UPDATE OR DELETE OR INSERT
	ON A7.TB_REGR_SIST_DEST
	FOR EACH ROW
--
DECLARE
	--
	--Vari�vel de controle de Opera�ao (1 = Alter ; 2 = Delete, 3 = Insert)
	nIN_TIPO_OPER  NUMBER := 0;
	--
BEGIN
	--
	IF UPDATING THEN
		nIN_TIPO_OPER := 1;
	ELSIF DELETING THEN
		nIN_TIPO_OPER := 2;
	ELSE
		nIN_TIPO_OPER := 3;
	END IF;
	--
	INSERT	INTO A7.TB_LOG_REGR_SIST_DEST
			(TP_MESG,
			TP_FORM_MESG_SAID,
			SG_SIST_ORIG,
			CO_EMPR_ORIG,
			DT_INIC_VIGE_REGR_TRAP,
			SG_SIST_DEST,
			CO_EMPR_DEST,
			IN_TIPO_OPER,
			CO_USUA_OPER,
			CO_ETCA_TRAB_OPER,
			DH_OPER)
	VALUES	(:OLD.TP_MESG,
			:OLD.TP_FORM_MESG_SAID,
			:OLD.SG_SIST_ORIG,
			:OLD.CO_EMPR_ORIG,
			:OLD.DT_INIC_VIGE_REGR_TRAP,
			:OLD.SG_SIST_DEST,
			:OLD.CO_EMPR_DEST,
			nIN_TIPO_OPER,
			NULL,
			NULL,
			SYSDATE);
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20001,'Erro na execu�ao da Trigger TGAR_REGR_SIST_DEST'
									||'A7.TB_LOG_REGR_SIST_DEST'
									|| SQLCODE || ' - '
									|| SUBSTR(SQLERRM, 1, 100));
	--
END;
/
