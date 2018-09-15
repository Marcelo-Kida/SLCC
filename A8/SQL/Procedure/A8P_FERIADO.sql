/*---------------------------------------------------------------------------------

OBS: DAR GRANT DE SELECT PARA O USER A8RPOC PARA A8.TB_FERIADO_HO

---------------------------------------------------------------------------------*/

CREATE OR REPLACE PROCEDURE    A8P_FERIADO
(
	P_DT_FERI  	IN	DATE,
	P_RETORNO  	OUT	NUMBER
)

IS

BEGIN
	SELECT COUNT(DT_FERI) INTO P_RETORNO
		FROM A8.TB_FERIADO_HO
		WHERE
			DT_FERI = trunc(P_DT_FERI);
	EXCEPTION
		WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20006, 'Execução - Erro na execução da Procedure A8P_FERIADO '  || SQLCODE || ' - ' || SQLERRM);
	RETURN;

END A8P_FERIADO;

/

Grant Execute on A8P_FERIADO to SLCCUSER;


