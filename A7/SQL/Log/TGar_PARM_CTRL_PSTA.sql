--
--Cria�ao de Trigger de Log  : A7PROC.TGAR_PARM_CTRL_PSTA
--Tabela associada           : TB_PARM_CTRL_PSTA
--Data Cria�ao               : 29-08-2003 12:55:38
--
CREATE OR REPLACE TRIGGER TGAR_PARM_CTRL_PSTA
	AFTER UPDATE OR DELETE OR INSERT
	ON A7.TB_PARM_CTRL_PSTA
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
	INSERT	INTO A7.TB_LOG_PARM_CTRL_PSTA
			(DH_PARM,
			QT_TEMP_ENTG_MESG,
			QT_TENT_ENVI_MESG,
			QT_TEMP_RETI_MESG,
			QT_FREQ_VERI,
			CO_USUA_ULTI_ATLZ,
			CO_ETCA_TRAB_ULTI_ATLZ,
			DH_ULTI_ATLZ,
			IN_TIPO_OPER,
			CO_USUA_OPER,
			CO_ETCA_TRAB_OPER,
			DH_OPER)
	VALUES	(:OLD.DH_PARM,
			:OLD.QT_TEMP_ENTG_MESG,
			:OLD.QT_TENT_ENVI_MESG,
			:OLD.QT_TEMP_RETI_MESG,
			:OLD.QT_FREQ_VERI,
			:OLD.CO_USUA_ULTI_ATLZ,
			:OLD.CO_ETCA_TRAB_ULTI_ATLZ,
			:OLD.DH_ULTI_ATLZ,
			nIN_TIPO_OPER,
			NULL,
			NULL,
			SYSDATE);
	--
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20001,'Erro na execu�ao da Trigger TGAR_PARM_CTRL_PSTA'
									||'A7.TB_LOG_PARM_CTRL_PSTA'
									|| SQLCODE || ' - '
									|| SUBSTR(SQLERRM, 1, 100));
	--
END;
/
