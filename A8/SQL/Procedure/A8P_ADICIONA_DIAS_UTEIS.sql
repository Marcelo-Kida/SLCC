CREATE OR REPLACE PROCEDURE A8P_ADICIONA_DIAS_UTEIS

(
    P_DATA   IN  DATE,
    P_QTDE_DIAS  IN  NUMBER,
    P_MOVIMENTO  IN  NUMBER,
    P_DATA_CALC  OUT  DATE
)
IS

vIncremento         NUMBER;
vQtdeDiasValidos    NUMBER;
vFeriado            NUMBER;

BEGIN
--
    IF P_MOVIMENTO = 2 THEN
        vIncremento := -1;
    ELSE
        vIncremento :=  1;
    END IF;
--
    vQtdeDiasValidos := 0;
    P_DATA_CALC := P_DATA;
--
    WHILE P_QTDE_DIAS > vQtdeDiasValidos LOOP
        P_DATA_CALC := P_DATA_CALC + vIncremento;
--
        IF TO_CHAR(P_DATA_CALC, 'D') NOT IN (1, 7) THEN
            A8P_FERIADO (P_DATA_CALC, vFeriado);
--
            IF vFeriado = 0 THEN
                vQtdeDiasValidos := vQtdeDiasValidos + 1;
            END IF;
        END IF;
    END LOOP;
--
    EXCEPTION
        WHEN OTHERS THEN
            RAISE_APPLICATION_ERROR(-20006, 'Execução - Erro na execução da Procedure A8P_ADICIONA_DIAS_UTEIS '  || SQLCODE || ' - ' || SQLERRM);
--
END A8P_ADICIONA_DIAS_UTEIS;

/

Grant Execute on A8P_ADICIONA_DIAS_UTEIS to SLCCUSER;

