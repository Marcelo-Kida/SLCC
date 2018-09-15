
CREATE OR REPLACE
FUNCTION    A8F_PROXIMO_DIA_UTIL
  ( P_DATA_BASE IN DATE)

RETURN  DATE IS

p_data_calc DATE;

BEGIN

  A8Proc.a8p_adiciona_dias_uteis(P_DATA_BASE,1,1,p_data_calc);

    RETURN p_data_calc;

EXCEPTION
WHEN OTHERS THEN
    RAISE_APPLICATION_ERROR(-20006, 'Execução - Erro na execução da Function A8F_PROXIMO_DIA_UTIL '  || SQLCODE || ' - ' || SQLERRM);
END; 
/

Grant Execute on A8F_PROXIMO_DIA_UTIL to SLCCUSER;

