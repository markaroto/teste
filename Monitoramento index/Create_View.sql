/* Formatted on 2016/09/14 08:54 (Formatter Plus v4.8.8) */
CREATE OR REPLACE FORCE VIEW oraadm.index_comparativo (owner,
                                                       index_name,
                                                       "Ultima_utilização",
                                                       "Tamanho_MB"
                                                      )
AS
   SELECT   t.owner, t.index_name,
            CASE
               WHEN MAX (a.TIMESTAMP) IS NULL
                  THEN 'NaoUtilizado'
               ELSE TO_CHAR
                      (MAX (a.TIMESTAMP),
                       'DD-MM-YYYY HH24:MI'
                      )                              --AS "Ultima_utilização",
            END AS "UltimaUtilizacao",
            (t.BYTES / 1024 / 1024) "Tamanho_MB"
       FROM oraadm.index_total t LEFT JOIN oraadm.index_audit a
            ON a.name_index = t.index_name AND t.owner = a.owner
   --where A.TIMESTAMP is not null
   GROUP BY t.owner, t.index_name, t.BYTES;
/;
/* Formatted on 2016/09/14 08:54 (Formatter Plus v4.8.8) */
CREATE OR REPLACE FORCE VIEW oraadm.index_total (owner,
                                                 index_name,
                                                 index_type,
                                                 table_owner,
                                                 table_name,
                                                 BYTES
                                                )
AS
   SELECT i.owner, i.index_name, i.index_type, i.table_owner, i.table_name,
          s.BYTES
     FROM dba_indexes i JOIN dba_segments s
          ON i.owner = s.owner AND s.segment_name = i.index_name
    WHERE i.owner NOT IN ('SYSAUX', 'SYSTEM', 'NEXADB', 'ORAADM', 'SYS')
      AND i.index_type = UPPER ('normal')
      AND i.tablespace_name NOT IN ('SYSTEM', 'SYSAUX');
/;
create OR REPLACE FORCE  view oraadm.index_ultimo_30_dias as
SELECT   t.owner, t.index_name,
            CASE
               WHEN MAX (a.TIMESTAMP) IS NULL
                  THEN 'NaoUtilizado'
               ELSE TO_CHAR
                      (MAX (a.TIMESTAMP),
                       'DD-MM-YYYY HH24:MI'
                      )                              --AS "Ultima_utilização",
            END AS "UltimaUtilizacao",
            count (A.NAME_INDEX ) as "QuantidadeUtilizacao",            
            (t.BYTES / 1024 / 1024) "Tamanho_MB"
       FROM oraadm.index_total t LEFT JOIN oraadm.index_audit a
            ON a.name_index = t.index_name AND t.owner = a.owner and to_date(A.TIMESTAMP,'dd/mm/yy') >= to_date((sysdate -30 ),'dd/mm/yy')
   --where A.TIMESTAMP is not null
   GROUP BY t.owner, t.index_name, t.BYTES;

	  
