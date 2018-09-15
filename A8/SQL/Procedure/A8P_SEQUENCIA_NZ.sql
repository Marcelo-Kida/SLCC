/*---------------------------------------------------------------------------------

OBS: DAR GRANT DE SELECT PARA A8PROC PARA A SEQUENCE A8.SQ_A8_NU_SEQU_REME_PJ

---------------------------------------------------------------------------------*/

CREATE OR REPLACE PROCEDURE A8P_SEQUENCIA_NZ
		(
		 PSISTEMA        IN      VARCHAR2,
		 PSEQUENCIA OUT VARCHAR2
		 )

IS 

BEGIN
	SELECT	A8.SQ_A8_NU_SEQU_REME_PJ.NEXTVAL
	INTO	PSEQUENCIA
	FROM	DUAL
	WHERE	ROWNUM = 1;
	
	PSEQUENCIA := TO_CHAR(SYSDATE, 'YYYYMMDD') || RPAD(PSISTEMA, 3, ' ') || '20' || LPAD(PSEQUENCIA, 7, '0');
	
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20002, 'A8P_SEQUENCIA_NZ - Execução Erro na execução da Procedure PKP_SEQUENCIA_NUNZ ' 
						|| SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 100));
			RETURN;			

END A8P_SEQUENCIA_NZ;
/

Grant Execute on A8P_SEQUENCIA_NZ to SLCCUSER;

