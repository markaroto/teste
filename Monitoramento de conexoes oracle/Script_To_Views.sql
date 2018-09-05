CREATE OR REPLACE FORCE VIEW oraadm.sessoesativas (maquina,
                                                   usuario,
                                                   tempo,
                                                   conexao
                                                  )
AS
   SELECT   userhost AS maquina, username AS usuario, hora.tempo AS tempo,
            'Loggons no periodo ate proxima hora.' AS conexao
       FROM SYS.dba_audit_trail,
            (SELECT     TRUNC (SYSDATE - 30) + (LEVEL / 24) AS tempo      --30
                   FROM DUAL
             CONNECT BY LEVEL <= 744) hora                               --744
      WHERE hora.tempo <= SYSDATE
        AND TIMESTAMP <= hora.tempo
        AND (   (action = '101' AND logoff_time >= hora.tempo)
             OR (action = '100' AND logoff_time IS NULL)
            )
   ORDER BY tempo DESC;

   /* Formatted on 2016/12/28 10:58 (Formatter Plus v4.8.8) */
CREATE OR REPLACE FORCE VIEW oraadm.conexoes_inicializadas (maquina,
                                                            usuario,
                                                            tempo,
                                                            conexao
                                                           )
AS
   SELECT userhost AS maquina, username AS usuario,
          TRUNC (TIMESTAMP, 'hh24') AS tempo         --,count(*) as UTILIZACAO
                                            ,
          'Conexoes criado no periodo' AS conexao
     FROM dba_audit_trail
    WHERE action = '101' OR action = '100';



CREATE OR REPLACE FORCE VIEW oraadm.conexoes_finalizadas (maquina,
                                                          usuario,
                                                          periodo,
                                                          conexoes_periodo
                                                         )
AS
    SELECT   userhost AS maquina, username AS usuario,
             TRUNC (logoff_time, 'hh24') AS periodo,
			 'Conexoes Finalizadas no periodo' AS conexao             
        FROM dba_audit_trail
       WHERE action = '101'
    ;
	
	
CREATE OR REPLACE FORCE VIEW oraadm.all_sessions (maquina,
                                                  usuario,
                                                  tempo,
                                                  conexao
                                                 )
AS
   (SELECT tab.maquina, tab.usuario, tempo AS tempo, conexao
      FROM (SELECT maquina, usuario,
                                    --TO_DATE (tempo, 'DD/MM/YY HH24:MI:ss') AS tempo
                                    tempo
                                         --,utilizacao
                   , conexao
              FROM oraadm.conexoes_inicializadas
            UNION ALL
            SELECT maquina, usuario, tempo AS tempo
                                                   --,utilizacao
                   , conexao
              FROM oraadm.sessoesativas
			UNION ALL
			SELECT maquina, usuario, tempo AS tempo
                                                   --,utilizacao
                   , conexao
              FROM oraadm.conexoes_finalizadas
			 ) tab);