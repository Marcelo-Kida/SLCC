       IDENTIFICATION DIVISION. 					
       PROGRAM-ID.    A8C0001.						
       AUTHOR.	      MARCELO GARCIA.					
       DATE-WRITTEN.  ABRIL/2004.					
       DATE-COMPILED.							
      *---------------------------------------------------------------* 
      * 							      * 
      *     AVALIA O BANCO RECEBIDO E CHAMA AS ROTINAS XPOXP08 OU     * 
      *     DRAT0045.						      * 
      *     CHAMA A ROTINA PASSANDO BCO, AGENCIA E DATA (AAAAMMDD)    * 
      *     DEVOLVE CODIGO IDENTIFICADOR DE DIA UTIL (S/N) E PROXIMO  * 
      *     DIA UTIL.						      * 
      * 							      * 
      *     RETORNA 'S' PARA SIM EH DIA UTIL OU 		      * 
      * 	    'N' PARA NAO EH DIA UTIL			      * 
      * 	     NO CAMPO A8W0001-02-IN-DT-UTIL		      * 
      *     E  PROXIMO DIA UTIL NO A8W0001-02-PROX-UTIL (AAAAMMDD)    * 
      * 							      * 
      *---------------------------------------------------------------* 
       ENVIRONMENT DIVISION.						
       CONFIGURATION SECTION.						
       SPECIAL-NAMES. DECIMAL-POINT IS COMMA.				
       DATA DIVISION.							
									
       WORKING-STORAGE SECTION. 					
									
       01  FILLER.							
									
	   03  FILLER			 PIC  X(034)	   VALUE	
	   'INICIO DA WORK DO PROGRAMA A8C0001'.			
									
      *---------------------------------------------------------------* 
      *    CAMPOS DE USO GERAL					      * 
      *---------------------------------------------------------------* 
									
	   03  WK-TSSTATUS		 PIC S9(009) COMP  VALUE ZEROS. 
	   03  WK-NOMEDATS.						
	       05  WK-NOMETRAN		 PIC  X(004)	   VALUE SPACES.
	       05  WK-NOMETERM		 PIC  X(004)	   VALUE SPACES.
									
	   03  WK-ITEM			 PIC S9(004) COMP  VALUE 0.	
									
	   03  WK-PROGRAMA		 PIC  X(008) VALUE SPACES.	
	   03  WK-LEN-COMM		 PIC S9(004) COMP VALUE ZEROS.	
									
	   03  WK-DT-AAAAMMDD.						
	      05  WK-DT-A8-ANO		 PIC  9(004).			
	      05  WK-DT-A8-MES		 PIC  9(002).			
	      05  WK-DT-A8-DIA		 PIC  9(002).			
									
	   03  WK-DT-RECB-GREG. 					
	      05  WK-DT-GREG-DIA	 PIC  9(002).			
	      05  WK-DT-GREG-MES	 PIC  9(002).			
	      05  WK-DT-GREG-ANO	 PIC  9(004).			
									
	   03 WK-TRACE. 						
	      05  WK-TRACE-LEN		 PIC S9(004) COMP  VALUE +80.	
	      05  WK-TRACE-VAL. 					
		  07  WK-CAMPO		 PIC  X(010).			
		  07  WK-VALOR		 PIC  X(070).			
									
      *--- COMANDO PARA DEPURAR PROGRAMA
      *    MOVE 'ROTINA ' TO WK-CAMPO
      *    MOVE 'BANESPA' TO WK-VALOR
      *    EXEC CICS ENTER TRACENUM (1)
      * 	FROM (WK-TRACE-VAL)
      * 	FROMLENGTH (WK-TRACE-LEN)
      *    END-EXEC

      *---------------------------------------------------------------* 
      *    AREA PARA COPY DOS BOOKS				      * 
      *---------------------------------------------------------------* 
									
      *--- BOOK DO ARQ. A8W0001 					
	   COPY  A8W0001.						
									
      *--- BOOK DO ARQ. XPWXP08 					
	   COPY  XPWXP08.						
									
      *--- BOOK DRCE0045 - ROTINA DRAT0045 (AMBIENTE BANESPA)		
	   COPY  A8WDR45.						
									
      *---------------------------------------------------------------* 
      *    AREA PARA DECLARE DOS CURSORES			      * 
      *---------------------------------------------------------------* 
									
       LINKAGE SECTION. 						
									
       01  DFHCOMMAREA. 						
	   05  LK-COMMAREA		 PIC  X(150).			
									
       PROCEDURE DIVISION.						
									
	   PERFORM  R010-CONSISTE-FISICA  THRU	R010-EXIT.		
									
	   PERFORM  R200-PROCESSA	  THRU	R200-EXIT.		
									
	   GO TO    R900-RETORNA.					
									
      *---------------------------------------------------------------* 
      *    CONSISTENCIA FISICA DA ENTRADA			      * 
      *---------------------------------------------------------------* 
       R010-CONSISTE-FISICA.						
									
      *--- CRITICA COMMAREA						

	   EXEC CICS IGNORE CONDITION INVREQ END-EXEC.			
	   EXEC CICS IGNORE CONDITION LENGERR END-EXEC. 		
									
	   MOVE EIBTRNID TO WK-NOMETRAN
	   MOVE EIBTRMID TO WK-NOMETERM
	   EXEC CICS DELETEQ TS QUEUE (WK-NOMEDATS)
				RESP  (WK-TSSTATUS)
	   END-EXEC.

      *--- MOVIMENTA A COMMAREA DA LINKAGE PARA WORK			
									
	   MOVE  LK-COMMAREA	       TO  A8W0001-COMMAREA.		
									
	   IF EIBCALEN NOT EQUAL A8W0001-LENGTH-COMMAREA		
	      MOVE 05 TO A8W0001-01-CD-RETORNO				
	      GO  TO  R900-FINALIZA-ERRO.				
									
      *--- INITIALIZE A8W0001-RETORNO					
									
	   IF (A8W0001-CD-BANCO NOT NUMERIC) OR 			
	      (A8W0001-CD-BANCO LESS ZEROS)				
	       MOVE    01    TO  A8W0001-01-CD-RETORNO			
	       GO  TO  R900-FINALIZA-ERRO				
	   END-IF.							
									
	   IF (A8W0001-CD-AGENCIA NOT NUMERIC) OR			
	      (A8W0001-CD-AGENCIA LESS ZEROS)				
	       MOVE    02    TO  A8W0001-01-CD-RETORNO			
	       GO  TO  R900-FINALIZA-ERRO				
	   END-IF.							
									
	   IF (A8W0001-DT-INFORMADA NOT NUMERIC) OR			
	      (A8W0001-DT-INFORMADA LESS ZEROS) 			
	       MOVE    03    TO  A8W0001-01-CD-RETORNO			
	       GO  TO  R900-FINALIZA-ERRO				
	   END-IF.							
									
	R010-EXIT.							
	    EXIT.							
									
      *---------------------------------------------------------------* 
      *    PROCESSA OS DADOS					      * 
      *---------------------------------------------------------------* 
       R200-PROCESSA.							
									
	   IF A8W0001-CD-BANCO EQUAL 33

	       PERFORM	R202-PROCESSA-ROTINA-BANESPA			
		  THRU	R202-EXIT					

	   ELSE

	       PERFORM	R201-PROCESSA-ROTINA-SANTANDER			
		  THRU	R201-EXIT					

	   END-IF.
									
       R200-EXIT.							
	   EXIT.							

      *---------------------------------------------------------------* 
      *    EFETUA A CHAMADA PARA A ROTINA SANTANDER(XPOXP08)	      * 
      *---------------------------------------------------------------* 
       R201-PROCESSA-ROTINA-SANTANDER.					

      *--- MONTA O REGISTRO PARA O LINK COM O XPOXP08			
									
	   MOVE A8W0001-CD-BANCO     TO  XPWXP08-NBANCO.		
	   MOVE A8W0001-CD-AGENCIA   TO  XPWXP08-CODAGE.		
	   MOVE A8W0001-DT-INFORMADA TO  XPWXP08-DTINFO.		
									
	   MOVE 'XPOXP08'	     TO  WK-PROGRAMA.			
	   MOVE +25		     TO  WK-LEN-COMM.			
									
	   EXEC CICS LINK PROGRAM (WK-PROGRAMA) 			
			  COMMAREA (XPWXP08-DADOSX)			
			  LENGTH  (WK-LEN-COMM) 			
	   END-EXEC.							
									
	   MOVE 'RET XP ' TO WK-CAMPO.					
	   MOVE XPWXP08-RTCODE	TO WK-VALOR.				
	   EXEC CICS ENTER TRACENUM (2)
		FROM (WK-TRACE-VAL)
		FROMLENGTH (WK-TRACE-LEN)
	   END-EXEC.

	   IF  XPWXP08-RTCODE NOT EQUAL ZEROS				
	       MOVE XPWXP08-RTCODE   TO  A8W0001-01-CD-RETORNO		
	       GO TO R900-FINALIZA-ERRO 				
	   END-IF.							
									
      *--- LIMPA CODIGO DE RETORNO
	   MOVE  ZEROS		     TO A8W0001-01-CD-RETORNO.		
	   MOVE  SPACES 	     TO A8W0001-01-MENSAGEM.		

	   IF  XPWXP08-PERIOD EQUAL +7 OR +8 OR +9			
	       MOVE 'N' 	     TO A8W0001-02-IN-DT-UTIL		
	   ELSE 							
	       MOVE 'S' 	     TO A8W0001-02-IN-DT-UTIL		
	   END-IF.							
									
	   MOVE  XPWXP08-PROXDT      TO A8W0001-02-DT-PROX-UTIL.	

       R201-EXIT.							
	   EXIT.							
      *---------------------------------------------------------------* 
      *    EFETUA A CHAMADA PARA A ROTINA BANESPA (DRAT0045)	      * 
      *---------------------------------------------------------------* 
       R202-PROCESSA-ROTINA-BANESPA.					

      *--- CONVERTE A DATA RECEBIDA PARA FORMATO GREGORIANO
	   MOVE A8W0001-DT-INFORMADA TO  WK-DT-AAAAMMDD.
	   MOVE WK-DT-A8-ANO	     TO  WK-DT-GREG-ANO.
	   MOVE WK-DT-A8-MES	     TO  WK-DT-GREG-MES.
	   MOVE WK-DT-A8-DIA	     TO  WK-DT-GREG-DIA.

	   INITIALIZE	 W045-DRCE0045.

	   MOVE 01		     TO  W045-OPC.
	   MOVE 001		     TO  W045-TPO-UOR.
	   MOVE A8W0001-CD-AGENCIA   TO  W045-COD-UOR.			
	   MOVE ZEROS		     TO  W045-COD-IDN-DPN.		
	   MOVE WK-DT-RECB-GREG      TO  W045-DAT-1.			
	   MOVE 1		     TO  W045-TPO-TRT.

	   MOVE 'DRAT0045'	     TO  WK-PROGRAMA.			
	   MOVE +300		     TO  WK-LEN-COMM.			

	   EXEC CICS LINK PROGRAM (WK-PROGRAMA) 			
			  COMMAREA (W045-DRCE0045)			
			  LENGTH  (WK-LEN-COMM) 			
	   END-EXEC.							

	   IF  W045-COD-RET  NOT EQUAL ZEROS				
	       MOVE W045-COD-RET     TO  A8W0001-01-CD-RETORNO		
	       GO  TO  R900-FINALIZA-ERRO				
	   END-IF.							
									
      *--- LIMPA CODIGO DE RETORNO
	   MOVE  ZEROS		     TO A8W0001-01-CD-RETORNO.		
	   MOVE  SPACES 	     TO A8W0001-01-MENSAGEM.		

	   IF  W045-IND-UTI  EQUAL ZEROS				
	       MOVE 'S' 	     TO A8W0001-02-IN-DT-UTIL		
	   ELSE 							
	       MOVE 'N' 	     TO A8W0001-02-IN-DT-UTIL		
	   END-IF.							
									
      *--- CONVERTE A DATA DE RETORNO PARA FORMATO DO SISTEMA A8
	   MOVE W045-DAT-UTI-POS(1)  TO  WK-DT-RECB-GREG.		
	   MOVE WK-DT-GREG-ANO	     TO  WK-DT-A8-ANO.
	   MOVE WK-DT-GREG-MES	     TO  WK-DT-A8-MES.
	   MOVE WK-DT-GREG-DIA	     TO  WK-DT-A8-DIA.
	   MOVE WK-DT-AAAAMMDD	     TO  A8W0001-02-DT-PROX-UTIL.	

       R202-EXIT.							
	   EXIT.							
      *--------------------------------------------------------------*	
      *   AVALIA ERRO E MONTA MENSAGEM DE RESPOSTA		     *	
      *--------------------------------------------------------------*	
       R900-FINALIZA-ERRO.						
									
	   IF WK-PROGRAMA EQUAL 'DRAT0045'				
      *--- AVALIA ERROS DA ROTINA DRAT0045 - BANESPA			

	     EVALUATE TRUE						
		WHEN A8W0001-01-CD-RETORNO = 01 			
		   MOVE 'OPCAO/DATA/TAMANHO COMMAREA INVALIDA'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 02 			
		   MOVE 'DATA GREGORIANA - MES INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 03 			
		   MOVE 'DATA GREGORIANA - DIA INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 04 			
		   MOVE 'DATA JULIANA	 - DIA INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 05 			
		   MOVE 'QTDE. DIAS NAO NUMERICA (OPC. 03 E 04)'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 06 			
		   MOVE 'IDENTIFICADOR DE DEPENDENCIA INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 07 			
		   MOVE 'TIPO OU CODIGO DA UNIOR NAO NUMERICO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 33 			
		   MOVE 'TIPO DE AREA GEOGRAFICA INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 34 			
		   MOVE 'CODIGO DE AREA GEOGRAFICA INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 35 			
		   MOVE 'AREA GEOGRAFICA NAO CADASTRADA'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 36 			
		   MOVE 'CODIGO MUNICIPIO BACEN INVALIDO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 37 			
		   MOVE 'CODIGO MUNICIPIO BACEN NAO CADASTRADO'
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 99 			
		   MOVE 'ABEND NA ROTINA'
				     TO A8W0001-01-MENSAGEM		
	     END-EVALUATE						
									
	   ELSE 							
      *--- AVALIA ERROS DE CONSISTENCIA FISICA E ERROS XPOXP08		

	     EVALUATE TRUE						
		WHEN A8W0001-01-CD-RETORNO = 01 			
		   MOVE 'CODIGO DE BANCO INVALIDO'			
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 02 			
		   MOVE 'CODIGO DE AGENCIA INVALIDA'			
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 03 			
		   MOVE 'DATA INFORMADA INVALIDA'			
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 05 			
		   MOVE 'ERRO DE COMUNICACAO - TAMANHO' 		
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 20 			
	       MOVE 'BCO/AG NAO NUMERICO, NAO INFORMADO OU INEXISTENTE' 
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 21 			
		   MOVE 'AGENCIA INEXISTENTE'				
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 25 			
	       MOVE 'DATA INFORMADA NAO NUMERICA, ZERADA OU INVALIDA'	
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 29 			
		   MOVE 'DATA MENOR QUE 1974ALIDO'			
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 91 			
		   MOVE 'ERRO ROTINA UGS0090'				
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 94 			
		   MOVE 'ERRO NA ROTINA HOH0097 (AGENCIA 424)'		
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 95 			
		   MOVE 'ERRO NA ROTINA HOH0092 (AGENCIA 353)'		
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 96 			
		   MOVE 'ERRO NA ROTINA HOH0018 (QUALQUER BANCO)'	
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 97 			
		   MOVE 'ERRO NA ROTINA UGS0090'			
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 98 			
		   MOVE 'ERRO NA ROTINA HOH0004 (CARTEIRA)'		
				     TO A8W0001-01-MENSAGEM		
		WHEN A8W0001-01-CD-RETORNO = 99 			
		   MOVE 'ERRO NAO PREVISTO'				
				     TO A8W0001-01-MENSAGEM		
		WHEN OTHER						
		   MOVE 'ERRO DA ROTINA XPHXP08'			
				     TO A8W0001-01-MENSAGEM		
	     END-EVALUATE						
									
	   END-IF.							

	   MOVE  SPACES 	    TO A8W0001-02-IN-DT-UTIL		
	   MOVE  ZEROS		    TO A8W0001-02-DT-PROX-UTIL. 	
									
      *--------------------------------------------------------------*	
      *    RETORNA AO PROGRAMA CHAMADOR 			     *	
      *--------------------------------------------------------------*	
       R900-RETORNA.							

	   MOVE  A8W0001-RETORNO    TO	DFHCOMMAREA.			
									
	   MOVE  1 TO WK-ITEM.

	   EXEC  CICS WRITEQ TS QUEUE (WK-NOMEDATS)
				FROM  (A8W0001-RETORNO)
				LENGTH (A8W0001-LENGTH-TS)
				ITEM (WK-ITEM)
	   END-EXEC.

	   EXEC  CICS SYNCPOINT END-EXEC.				
									
	   EXEC  CICS  RETURN  END-EXEC.				
									
       R900-EXIT.							
	   EXIT.							

