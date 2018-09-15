CREATE OR REPLACE PROCEDURE        A8P_NZO0040 (
                                      co_opca         in    number
                                    , co_empre        in    number
                                    , sg_siste        in    varchar2
				    , dt_movt         in    date
                                    , qt_nu_sequ      in    number
                                    , nu_sequ_inic    out   varchar2
                                    , nu_sequ_fina    out   varchar2
                                    , tabela_oracle   out   varchar2
                                    , funcao_oracle   out   varchar2
				    , sql_code        out   number
                                    , rc_rotina       out   number
                                    , msg_rc_rotina   out   varchar2
 )
IS

BEGIN

 NZP_NZO0040 (co_opca	  , co_empre	 , sg_siste	, dt_movt , qt_nu_sequ, nu_sequ_inic,
	      nu_sequ_fina, tabela_oracle, funcao_oracle, sql_code, rc_rotina , msg_rc_rotina);

 EXCEPTION
  WHEN OTHERS THEN
   RAISE_APPLICATION_ERROR(-20002, 'A8P_NZO0040 - Execução Erro na execução da Procedure A8P_NZO0040 '
      || SQLCODE || ' - ' || SUBSTR(SQLERRM, 1, 100));
   RETURN;

END A8P_NZO0040;
/


Grant Execute on A8P_NZO0040 to SLCCUSER;

