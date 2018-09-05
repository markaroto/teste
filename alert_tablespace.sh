#!/bin/bash 
#********************************************************************************************************************
#*                                                                                                                 *
#*                                                                                                                  *
#* SCRIPT: alert_tablespace.sh                                                                                      *
#* DESCRIÇO: Realiza Monitoramento das tablespaces  			                                                    *
#*                                                                                                                  *
#* VERSAO: 1.0                                                                                                      *
#*                                                                                                                  *
#* PARAMETROS:                                                                                        				*
#* ATUALIZACOES:                                                                                                    *
#********************************************************************************************************************
#clear screen
#lista dos servidores 
#funções do codigo
source /home/oracle/.bash_profile
caminho_script=/scripts_backup/
mensagem=`echo -e "Tamanho tablespaces"`;
desc=`printf "%-20s%-15s%-10s%-10s%-15s%-15s%-10s%-15s" TableSpace_Name MB_Allocated MB_Used MB_Free Pct_Used Pct_Free Instancia Data`; 
mensagem=`echo -e "${mensagem} \n ${desc}"`;
enviar=2;
assunto="Tablespace com mais 90%";
function conexao_sql(){
	instancia=$1
	sair=17	
	tnsping ${instancia}
	error=$?
	if [ ${error} != 0 ];then
		arquivo_tem="Instancia  ${instancia} indisponivel "
		mensagem=`echo -e "${mensagem} \n ${arquivo_tem}"`
	else
		
	nome_instance='v$instance'
	arquivo_bruto=`sqlplus -s bkp_copasa/copasabkp@${instancia}<<EOF
SET LINES 1000
SET PAGES 1000
SELECT a.TABLESPACE_NAME "TableSpace_Name",
round(a.BYTES/1024/1024) "MB_Allocated",
round((a.BYTES-nvl(b.BYTES, 0)) / 1024 / 1024) "MB_Used",
nvl(round(b.BYTES / 1024 / 1024), 0) "MB_Free",
round(((a.BYTES-nvl(b.BYTES, 0))/a.BYTES)*100,2) "Pct_Used",
round((1-((a.BYTES-nvl(b.BYTES,0))/a.BYTES))*100,2) "Pct_Free",
(select instance_name from ${nome_instance}) "Instancia",
to_char(sysdate,'dd/mm/yyyy') "Data"
FROM (SELECT TABLESPACE_NAME,
sum(BYTES) BYTES
FROM dba_data_files
GROUP BY TABLESPACE_NAME) a,
(SELECT TABLESPACE_NAME,
sum(BYTES) BYTES
FROM sys.dba_free_space
GROUP BY TABLESPACE_NAME) b
WHERE a.TABLESPACE_NAME = b.TABLESPACE_NAME and round(((a.BYTES-nvl(b.BYTES, 0))/a.BYTES)*100,2) >= 90 ${arq_excl} ${excluir_perso}
ORDER BY ((a.BYTES-b.BYTES)/a.BYTES);
EOF`
	
	while [ ${sair} -ne 0 ]
		do
		valida=`echo ${arquivo_bruto} | cut -d ' ' -f ${sair}`
		arquivo_tem=""
		if [ -z ${valida} ]; then			
			sair=0
			#enviar=0
		else 
		enviar=$(($enviar +1))
		for((i=0; i <= 7; i++));
			do
				p=$((${sair}+${i}))
				temp[$i]=`echo ${arquivo_bruto} | cut -d ' ' -f ${p}`
				temp=`printf "%-30s" ${temp}`
				if [ ${i} -eq 8 ]; then 
					arquivo_tem=`echo -e "${arquivo_tem}  ${temp[$i]}"`
				else
					arquivo_tem=`echo -e "${arquivo_tem}\t  ${temp[$i]}"`
				fi
				
				
		done
		arquivo_tem=`printf "%-20s%-15s%-10s%-10s%-15s%-15s%-10s%-15s" ${temp[0]} ${temp[1]} ${temp[2]} ${temp[3]} ${temp[4]} ${temp[5]} ${temp[6]} ${temp[7]} `
		mensagem=`echo -e "${mensagem} \n ${arquivo_tem}"`
		sair=$(($sair+8));
		fi		
	done
	fi
}
#lista das instancias monitoradas analisado 
instancias=("" )
excluir=("" )
#contador para mudar localizaçao do array da variavel instancias
contador=0
#variavel para mudar localizaçao do array da variaveis tamanho
arquivo=1
##
arq_excl=""
while [ ! -z ${excluir[$contador]} ]
	do
	arq_excl=`echo "${arq_excl} and a.TABLESPACE_NAME !='${excluir[$contador]}'"`
	contador=$(($contador+1));
done

#contador para mudar localizaçao do array da variavel instancias
contador=0

#loop para localizar todos instancias
while [ ! -z ${instancias[$contador]} ]
	do
		#acionamento da função conexao_ssh
		conexao_sql "${instancias[$contador]}"
		contador=$(($contador+1));				
done
echo "${mensagem}"
echo $enviar
if [  ${enviar} -ne 2 ];then
	${caminho_script}mail.sh " ${mensagem}" "$assunto "
fi

