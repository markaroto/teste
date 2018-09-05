
<#PSScriptInfo

.VERSION 2.0

.GUID 415c07d5-6f39-4566-aa73-5a5a839d27f2

.AUTHOR Marcos Flavio Alves da Silva

.COMPANYNAME Nexa/Copasa

.COPYRIGHT 2017/05

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES DLL EPPlus.dll

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>


<# 

.DESCRIPTION 
 Script que realizar a coleta de informações sobre as fitas do oracle no netbackup 

#> 
Param(
#planilha local com excel.
[string]$arquivo_excel="C:\netbackup\politica.xlsx",
#local da copia remota da planilha
[string]$arquivo_excel1="\\copanet04\dvsu\CONFIGURACAO\ORACLE\PROJETOS ORACLE\SEGURANÇA\BACKUP ORACLE\ACOMPANHAMENTO DO BACKUP\Tabela de utilizacao Fitas por Pool.xlsx",
[string]$grafico_excel1="\\copanet04\dvsu\CONFIGURACAO\ORACLE\PROJETOS ORACLE\SEGURANÇA\BACKUP ORACLE\ACOMPANHAMENTO DO BACKUP\Grafico de utilizacao Fitas por Pool.xlsx",
#Planilha de grafico local.
[string]$grafico_excel="C:\netbackup\grafico.xlsx",
#Arquivo de historico.
[string]$historico_caminho="C:\netbackup\historico.csv",
#Arquivo de log.
[string]$logerro="C:\netbackup\log.txt",
#Arquivo com nome da politicas.
[string]$arquivo_servidores="C:\netbackup\politica.txt",
#Caminho do comando vmquery que lista as fitas por politica
[string]$caminho_volumgr="C:\Program Files\Veritas\Volmgr\bin\vmquery.exe",
#Caminho do comando bpmedialist que lista informações das fitas.
[string]$caminho_admincmd="C:\Program Files\Veritas\NetBackup\bin\admincmd\bpmedialist.exe"
)
#função para gravar historico.
function historico{
    # validar se o arquivo existe.
	if(! (Test-Path $historico_caminho)){
        $arquivo | Export-Csv $historico_caminho
    }else{
        # Atualizar o arquivo de historico existente.
		$temp_arquivo= Import-Csv $historico_caminho
        $temp=@()
        $tem=$temp_arquivo + $arquivo
        $tem | Export-Csv $historico_caminho
    }
}
#função para enviar email.
function enviar_email(){
	try{
		$assunto="Monitoria poll fitas $politica_email com ocupaçã ocupação de acima 70%"
		#Criação do body do email.
		$mensagem=$arquivo | where { $_.pool -match $politica_email }   | Out-String
		$mensagem=$mensagem.replace("tamanho","ESPAÇO OCUPADO GB")
		$mensagem=$mensagem.replace("validade","EXPIRAÇÃO")
		$mensagem=$mensagem.replace("Tamanho_maximo","CAPACIDADE MÉDIA GB")
		$mensagem=$mensagem.replace("pc_ocupado","PCT OCUPADO")
		$mensagem=$mensagem.replace("total_dias","TOTAL DIAS UTILIZAÇÃO")
		#Ip e porta para enviar o email.
		$smtp = new-object Net.Mail.SmtpClient("10.1.1.52","587")
		#senha para enviar o email.
		$smtp.Credentials = new-object System.Net.NetworkCredential("dvsu.nexa@copasa.com.br", "nex@2015")
		#O email sera enviado para:
		$user_email="marcos.flavio@nexa.com.br,Mardem.barbosa@copasa.com.br,rogerio.notini@copasa.com.br,roziane.barroso@copasa.com.br"
		#Criação do objecto da mensagem.
		$email = new-object System.Net.Mail.MailMessage("dvsu.nexa@copasa.com.br",$user_email) 
		#Email não tem html.
		$email.IsBodyHtml = $false
		#campo assunto.
		$email.Subject = $assunto
		#padrões de texto.
		$email.BodyEncoding=[System.Text.Encoding]::UTF8
		#prioridade do email.
		$email.Priority=2 
		# body do email.
		$email.Body = $mensagem
		# enivar email.
		$smtp.Send($email)
	}
	catch{
		#erro durante envio.
        ($_ | select-object  @{name="temp";expression={$(Get-Date) -f ("dd/MM/yyy HH:mm:ss") +" "+$_.Exception}}).temp -split "`n" | Select-Object -Index 0  | out-file -FilePath $logerro -Append		
	}
}
#criação da planilha total.
function grafico_crecimento {
	#lista politicas coletadas.
    $list_politica=$arquivo | Select-Object -Unique pool
	#criado objecto para trabalhar com excel.
    $excel= New-Object OfficeOpenXml.ExcelPackage($grafico_excel)
	#Se planilha existe.
    if( Test-Path $grafico_excel){        
        $plan=$excel.Workbook.Worksheets["dados"]        
    }else{
        #Configuração padrão dos arquivos.
		$plan= $excel.Workbook.Worksheets.Add("dados")
        $plan.Cells[1,1].value="DATA DA COLETA"
        $coluna=2
        foreach ( $pool in $list_politica){
            #criar titulo com politica
			$plan.Cells[1,$coluna].value=$pool.pool
            $coluna++
        }        
    }
	#alinhamento ponto buscar no excel.
    $linha=2
    $coluna=2
	#Loop com politicas.
    foreach ($pool in $list_politica){
        $coluna=2
		#loop para localizar a coluna da politica.
        while ($coluna -ne 0){
			#valor da coluna atual.
            $poll_atual=$plan.cells[1,$coluna].text
            #Se localizar a politica sair ou criar uma nova linha para politica.
			if ( $pool.pool -eq $poll_atual ){
                $coluna=0
            }else{
				# Se politica não existir criar ela.
                if ( $poll_atual -eq ""){                    
                    #Gravando nome nova politica.
					$plan.cells[1,$coluna].value=$pool.pool
                    $coluna=0
                }else{
					#Continuar a buscar
                    $coluna++
                }
            }
        }
    }
	#Atualizando a variavel da coluna.
    $coluna=2
	#loop com saida somente quando linha for igual 0
    while($linha -ne 0){
		#coletando a data execução da coleta.
        $data_exc=($arquivo | select -Unique data).data
        #coletando daados da linha x da coluna 1 que seria a data.
		$plandata=$plan.cells[$linha,1].text
		#Se a data for vazia.
		if($plan.Cells[$linha,1].value -eq $null ){
            echo "$linha"
			if($plan.Cells[$linha,1].Formula -eq "" ){
				# Gravar data no formato Americano.
				$plan.cells[${linha},1].Style.Numberformat.Format = "dd/mm/yyyy"
				$plan.cells[${linha},1].Formula="=DATE("+$(get-date $data_exc -f ("yyyy,MM,dd"))+")"
				$data_exc= get-date $data_exc -f ("dd/MM/yyyy")
				#$plan.cells[$linha,1].value=$data_exc
				# Loop referente a coluna.
				while($coluna -ne 0){
					#localizar a posição da politica.
					$plancoluna=$plan.Cells[1,$coluna].Text
					# Se politica estiver vazia sair.
					if( $plancoluna -eq "" ){
						#Sair.
						$coluna=0
					}else{
						#Gravando nome da coluna em variavel.
						$poll_atual=$plancoluna
						#Gravando na variavel somente dados da politica.
						$arquivo_total=$arquivo | Where-Object {$_.pool -like $poll_atual }
						$tamanho_poll=0
						#loop para fazer soma do tamanho.
						foreach($pol_tamanho in $arquivo_total){
							#Calculo do tamanho.
							[float]$tamanho_temp=($pol_tamanho | Select-Object tamanho).tamanho
							[float]$tamanho_poll= $tamanho_poll + $tamanho_temp 
						}
						#Gravando o tamanho na planilha.
						$plan.cells[$linha,$coluna].value=$tamanho_poll
						$coluna++
					}                
				}
            #Update da linha 0
			$linha=0
            }else{
				# Somente atualizar se o registro não existir.
				if ( $plan.Cells[${linha},1].Formula -eq "=DATE("+$(get-date $data_exc -f ("yyyy,MM,dd"))+")"){
					$linha=0				
				}else{		
					$linha++
				}			
			}
        }else{
            #Não atualizar se data já estiver no arquivo.
			if ($plan.cells[$linha,1].text -eq $data_exc ){
                $linha=0 
            }else{
                $linha++            
            }
        }
        
    }
	#Salvar as atualizações.
    $excel.Save()       
        
}

function gravar_excel{ 
	#criado objecto para trabalhar com excel.
    $excel = New-Object OfficeOpenXml.ExcelPackage($arquivo_excel)
	#Se planilha existe.
	if( Test-Path $arquivo_excel){        
        $plan=$excel.Workbook.Worksheets["dados"]        
    }else{
        #Configuração padrão dos arquivos.
		$plan= $excel.Workbook.Worksheets.Add("dados")        
		$plan.cells[1,1].Value="DATA DA COLETA"
        $plan.cells[1,2].Value="HORA COLETA"
        $plan.cells[1,3].Value="POOL"
        $plan.cells[1,4].Value="MÍDIA"
        $plan.cells[1,5].Value="CAPACIDADE MÉDIA DE ARMAZENAMENTO (GB)"
        $plan.cells[1,6].Value="ESPAÇO DISPONÍVEL (GB)"
        $plan.cells[1,7].Value="ESPAÇO CUPADO (GB)"
        $plan.cells[1,8].Value="% OCUPADO"
        $plan.cells[1,9].Value="STATUS DA FITA"
        $plan.cells[1,10].Value="DATA INICIO UTILIZAÇÃO"
        $plan.cells[1,11].Value="DATA EXPIRAÇÃO"
        $plan.cells[1,12].Value="QTDE DIAS UTILIZAÇÃO" 
    }
	$linha=1
	#Data da coleta da informações.
	$data_exc=($arquivo | select -Unique data).data
	$data_exc_f= "=DATE("+$(get-date $data_exc -f ("yyyy,MM,dd"))+")"
    $data_exc= get-date $data_exc -f ("MM/dd/yyyy")
    #Loop para criar arquivo.
	while( $linha -ne 0){		
        # Se o campo data estiver fazio.
		if ( $plan.Cells[$linha,1].value -eq $null){
			if($plan.Cells[$linha,1].Formula -eq "" ){
				#Gravando dados da variavel arquivo na variavel media_loops
                $media_loops=$arquivo
				#loop para gravação de informações no excel.
				foreach ($media_loop in $media_loops){                    
					$plan.cells[${linha},1].Style.Numberformat.Format = "dd/mm/yyyy"
					$plan.cells[${linha},1].Formula="=DATE("+$(get-date $media_loop.data -f ("yyyy,MM,dd"))+")"	               
					$plan.cells[${linha},2].value=$media_loop.hora_atual
					$plan.cells[${linha},3].value=$media_loop.pool
					$plan.cells[${linha},4].value=$media_loop.Midia
					$plan.cells[${linha},5].value=$media_loop.Tamanho_maximo
					$plan.cells[${linha},6].value=$media_loop.Disponivel
					$plan.cells[${linha},7].value=$media_loop.tamanho
					$plan.cells[${linha},8].value=$media_loop.pc_ocupado
					$plan.cells[${linha},9].value=$media_loop.Estatus_fita
					#Se o campo não for vazio.
					if ( $media_loop.incio_midia -ne $null){
						$plan.Cells[${linha},10].Style.Numberformat.Format="m/d/yy h:mm"
						$plan.cells[${linha},10].value=get-date $media_loop.incio_midia 
					}
					#Se o campo não for vazio.
					if ( $media_loop.validade -ne $null){
						$plan.Cells[${linha},11].Style.Numberformat.Format="m/d/yy h:mm"
						$plan.cells[${linha},11].value=get-date $media_loop.validade 
					}
					$plan.cells[${linha},12].value=$media_loop.total_dias             
					$linha++                              
				}
				$linha=0
			}
			else{
				# Somente atualizar se o registro não existir.				
                if ( $plan.Cells[${linha},1].Formula -eq $data_exc_f){
					$linha=0				
				}else{		
					$linha++
				}			
			}
        }else{
            # Somente atualizar se o registro não existir.
			if ( $plan.Cells[${linha},1].value -eq $data_exc){
				$linha=0				
			}else{		
			    $linha++
            }
        }        
    }
	#Grava update no arquivo
    $excel.Save()    
}

######################################################Incio do script#######################################################
$(Get-Date) -f ("dd/MM/yyy HH:mm:ss")+" Script iniciado com sucesso." | out-file -FilePath $logerro -Append
#Acionamento da dll.
Add-Type -Path $PSScriptRoot\EPPlus.dll
#retirar campos vazios do arquivo
try{
    $politicas=Get-Content $arquivo_servidores  -ErrorAction Stop
    $politicas= $politicas -split "`n" | where {$_ -notlike ""}
}catch{
    ($_ | select-object  @{name="temp";expression={$(Get-Date) -f ("dd/MM/yyy HH:mm:ss") +" "+$_.Exception}}).temp -split "`n" | Select-Object -Index 0  | out-file -FilePath $logerro -Append
    $(Get-Date) -f ("dd/MM/yyy HH:mm:ss")+" Script Cancelado." | out-file -FilePath $logerro -Append
    exit
}
#criação da variavel armazenamento
$arquivo=@()
# loop com todas politicas do arquivos.
foreach ($politica in $politicas){	
    try{
        #comando remoto no servidor BKPMASTER para coletar a informações nomes da fitas da politica.
		$dados=Invoke-Command -ComputerName bkpmaster -ScriptBlock { & $args[0] -h bkpmaster -pn $args[1] } -ArgumentList $caminho_volumgr,$politica -ErrorAction stop
        #tratamento das informações recebidas.
		$medias=($dados | where {$_ -match "media ID"} ) -split(" ") | where{$_ -notlike "" -and $_ -notmatch "ID:" -and $_ -notmatch "media" }
    }catch{
		#Se apresenta algum erro durante o processo.
        $medias=$null
    }  
	#Loop da fitas.
    foreach ($media in $medias) {        
        try{            
            #Commando remoto no servidor BKPMASTER
			$media_info= Invoke-Command -ComputerName bkpmaster -ScriptBlock { & $args[0] -ev $args[1] } -ArgumentList $caminho_admincmd,$media -ErrorAction stop
            #Processo de tratamento das informações recebidas pelo comando.
			$incio_midia=[datetime]$(($media_info[6] -split "\s+" | Select-Object -Index 3,4 ) -join " ")
            $tamanho_midia=([float]$($media_info[6] -split "\s+"  | Select-Object -Index 8) * 1kb) /1gb
            $date=[datetime]$(($media_info[7] -split "\s+"  | Select-Object -Index 2,3) -join " ")
            $estatus_midia=$media_info[7] -split "\s+"  | Select-String "FULL"
			#Colocando valor medio das fitas.
            [float]$tamanho_medio='1536'
			#Calculo de tamanho disponivel.
            $disponivel_espaco=$tamanho_medio-$tamanho_midia
			# Porcentagem de tamanho da fita.
            $pc_ocupado=100-(($disponivel_espaco*100)/$tamanho_medio)
            # tempo de utilização da fita.
			$total_tempo=($date - $incio_midia).days
            # Fitas sem infromações estão ativas. Linha sera descontinuada proxima versão.
			if($estatus_midia -eq $null){
                $estatus_midia="Ativa"
            }
			# Configurando padrões data.
            $hora_atual=(get-date).tostring("HH:mm:ss")
            $data_atual= (get-date).tostring("dd/MM/yyyy")
			# Criando um objecto com informações recebidas do netbackup.
            $banco= New-Object psobject -Property @{
                pool=$politica;
                Midia=$media;
                data=(get-date).tostring("dd/MM/yyyy");
                tamanho=$tamanho_midia;
                Estatus_fita=$estatus_midia;           
                incio_midia=$incio_midia;
                validade=$date;
                Tamanho_maximo=$tamanho_medio;
                Disponivel=$disponivel_espaco;
                pc_ocupado=$pc_ocupado;            
                total_dias=$total_tempo;
                hora_atual=(get-date).tostring("HH:mm:ss");
            }
        }
        catch{
			# Criando um objecto com informações recebidas do netbackup.
            $banco= New-Object psobject -Property @{
                pool=$politica;
                Midia=$media;             
                data=(get-date).tostring("dd/MM/yyyy")
                hora_atual=(get-date).tostring("HH:mm:ss") ;            
                tamanho=$null;
                Estatus_fita="Disponivel";            
                incio_midia=$null;
                validade=$null;
                Tamanho_maximo=$null;
                Disponivel=$null;
                pc_ocupado=$null;          
                total_dias=$null;
            }
        }
        finally{
			#Criação do objecto com as informações.
            $arquivo+=$banco
        }          
    }
} 
#Acionamento da função que criar o excel.
gravar_excel ""
#
$polita_email=$arquivo | Select-Object -Unique pool
foreach ( $p_email in $polita_email ){
    $politica_email=${p_email}.pool
    $quantida_p=($arquivo | Where-Object { $_.pool -match $politica_email  -and $_.estatus_fita -notmatch "FULL" -and $_.pc_ocupado -le 70 } )
    if ( $quantida_p -eq $null) { 
       #Acionamento da função enviar_email
	   enviar_email ""       
    }
}
#Acionamento da função grafico_crecimento
grafico_crecimento ""
#Acionamento da função historico 
historico ""

#copia das planilha para servidor.
try{
    copy-item -path $arquivo_excel -Destination $arquivo_excel1  -Force -ErrorAction stop
}catch{
    ($_ | select-object  @{name="temp";expression={$(Get-Date) -f ("dd/MM/yyy HH:mm:ss") +" "+$_.Exception}}).temp -split "`n" | Select-Object -Index 0  | out-file -FilePath $logerro -Append
}
try{
    copy-item -path $grafico_excel -Destination $grafico_excel1  -Force -ErrorAction Stop
}catch{
    ($_ | select-object  @{name="temp";expression={$(Get-Date) -f ("dd/MM/yyy HH:mm:ss") +" "+$_.Exception}}).temp -split "`n" | Select-Object -Index 0  | out-file -FilePath $logerro -Append
}
finally{
    $(Get-Date) -f ("dd/MM/yyy HH:mm:ss")+" Concluida" | out-file -FilePath $logerro -Append
}

