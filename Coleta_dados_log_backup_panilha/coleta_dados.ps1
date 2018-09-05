#Set-ExecutionPolicy bypass
Import-Module SSH-Sessions

function criar_excel(){
    $arquivos=@()
    $arquivo1=Import-Csv .\tamanho_backup.txt
    $arquivos += $arquivo1
    $exc= New-Object -ComObject excel.application
    $exc.DisplayAlerts=$false
    $exc.Visible= $true
    if ((Test-Path  $caminho_excel_mesal )){
        $work= $exc.Workbooks.Open($caminho_excel_mesal)
        $plan=$work.Worksheets | where {$_.name -like "dados"}
        if(!$plan){
            #($work.Worksheets | where {$_.name -like "dados"}).delete()
            $plan=$work.Worksheets.Add()
            $plan.name="dados"
            $plan.Cells.Item(1,1)="SERVIDOR"
            $plan.Cells.Item(1,2)="ALVO"
            $plan.Cells.Item(1,3)="UTILITARIO"
            $plan.Cells.Item(1,4)="TIPO DE BAKCUP"
            $plan.Cells.Item(1,5)="STATUS DA INSTANCIA ALVO"
            $plan.Cells.Item(1,6)="FREQUENCIA"
            $plan.Cells.Item(1,7)="RETENCAO (DIAS)"
            $plan.Cells.Item(1,8)="CALENDARIO"
            $plan.Cells.Item(1,9)="DATA INICIO BACKUP"
            $plan.Cells.Item(1,10)="DATA TEMINO BACKUP"
            $plan.Cells.Item(1,11)="DURACAO BACKUP (minuto)"
            $plan.Cells.Item(1,12)="TAMANHO BACKUP"
            $plan.Cells.Item(1,13)="CODIGO RETORNO"
        }             
                           
    }else{
        $work= $exc.Workbooks.add()
        $plan= $work.ActiveSheet
        $plan.name="dados"
        
        $plan.Cells.Item(1,1)="SERVIDOR"
        $plan.Cells.Item(1,2)="ALVO"
        $plan.Cells.Item(1,3)="UTILITARIO"
        $plan.Cells.Item(1,4)="TIPO DE BAKCUP"
        $plan.Cells.Item(1,5)="STATUS DA INSTANCIA ALVO"
        $plan.Cells.Item(1,6)="FREQUENCIA"
        $plan.Cells.Item(1,7)="RETENCAO (DIAS)"
        $plan.Cells.Item(1,8)="CALENDARIO"
        $plan.Cells.Item(1,9)="DATA INICIO BACKUP"
        $plan.Cells.Item(1,10)="DATA TEMINO BACKUP"
        $plan.Cells.Item(1,11)="DURACAO BACKUP (minuto)"
        $plan.Cells.Item(1,12)="TAMANHO BACKUP"
        $plan.Cells.Item(1,13)="CODIGO RETORNO"                         
    }    
    
    $x=1    
    foreach ($arquivo in $arquivos){
        $x=$x+1
         $plan.Cells.Item($x,1)=$arquivo.servidor
         $plan.Cells.Item($x,2)=$arquivo.instancia
         $utilitario=$arquivo.script -split "_" 
         if ( $utilitario -match "0"){
            $temp_arquivo=$arquivo.Local -split "/" | where {$_ -notlike ""} | select -Last 1
            if ($temp_arquivo -match "dpump"){
                $utilitario="bkpdpump"  -split "_"
            }else{
                $utilitario="RMAN"  -split "_"
            }            
         }          
         if ( $utilitario[0] -match "bkpdpump"  ){
            $utilitario[0]="DATA PUMP"
            $plan.Cells.Item($x,4)="FULL" 
            $plan.Cells.Item($x,5)="OK"           
         }elseif( $utilitario[0] -match "BKPEXP" ){
            $utilitario[0]="EXPORT"
            $plan.Cells.Item($x,4)="FULL" 
            $plan.Cells.Item($x,5)="OK"
         }elseif( $utilitario[0] -match "BKP-" ){
            $utilitario[0]="DATA PUMP"
            $plan.Cells.Item($x,4)="SCHEMA" 
            $plan.Cells.Item($x,5)="OK"
         }else{
            if (  $utilitario[1] -match "full" ){
                $plan.Cells.Item($x,4)="FULL"
                $plan.Cells.Item($x,5)="ON"
            }elseif( $utilitario[1] -match "arch"){
                $plan.Cells.Item($x,4)="ARCHIVE"
                $plan.Cells.Item($x,5)="ON"
            }elseif( $utilitario[1] -match "off"){
                $plan.Cells.Item($x,4)="OFF"
                $plan.Cells.Item($x,5)="OFF"
            }elseif( $utilitario[1] -match "est"){
                $plan.Cells.Item($x,4)="ESTRUTURA"
                $plan.Cells.Item($x,5)="ON"
            }            
         }
         $plan.Cells.Item($x,3)=$utilitario[0].toupper()
         if ( $arquivo.schedule -match "dia" ){
            if($arquivo.dia_semana -match "domingo"){
                $plan.Cells.Item($x,6)="SEMANAL"    
            }else{
                $plan.Cells.Item($x,6)="DIARIA"
            }
            $plan.Cells.Item($x,7)="15"
         }elseif ( $arquivo.schedule -match "sem_"){
            $plan.Cells.Item($x,6)="SEMANAL"
            $plan.Cells.Item($x,7)="90"
         }elseif ( $arquivo.schedule -match "archive"){
            $plan.Cells.Item($x,6)="HORA"
            $plan.Cells.Item($x,7)="15"         
         }elseif ( $arquivo.schedule -match "mens_"){
            $plan.Cells.Item($x,6)="MENSAL"
            $plan.Cells.Item($x,7)="365"
         }elseif ( $arquivo.schedule -match "semestral"){
            $plan.Cells.Item($x,6)="SEMESTRAL"
            $plan.Cells.Item($x,7)="180"
         }
         if ($arquivo.dia_semana -match "sexta-feira" -or $arquivo.dia_semana -match "segunda-feira" -or $arquivo.dia_semana -match "terça-feira" -or $arquivo.dia_semana -match "quarta-feira" -or $arquivo.dia_semana -match "quinta-feira" ){
            $plan.Cells.Item($x,8)="2a a 6a"      
         }else{
            $plan.Cells.Item($x,8)=$arquivo.dia_semana
         }
         $plan.Cells.Item($x,9)=Get-Date $arquivo.data_inicio -f ("MM/dd/yyyy HH:mm:ss")  
         $plan.Cells.Item($x,10)=Get-Date $arquivo.data_final -f ("MM/dd/yyyy HH:mm:ss")
         $plan.Cells.Item($x,11)=$arquivo.minutos_gasto -replace ",","."
         $plan.Cells.Item($x,12)=$arquivo.tamanho -replace ",","."
         $plan.Cells.Item($x,13)=$arquivo.rc


    }
    
    $work.Saveas($caminho_excel_mesal)
    $work.Close() 
    $exc.Quit()
    
   


        
}

function formato_arquivo(){
    $arquivo=@()
    foreach ($text in $texto){
        $tabela=New-Object psobject 
        $data=$text -split " " | where {$_ -notlike "" } | select -Index 0
        $hora=$text -split " " | where {$_ -notlike "" } | select -Index 1
        $servidor=$text -split " " | where {$_ -notlike "" } | select -Index 2
        $instancia=$text -split " " | where {$_ -notlike "" } | select -Index 3
        $script=$text -split " " | where {$_ -notlike "" } | select -Index 4
        $processo=$text -split " " | where {$_ -notlike "" } | select -Index 5
        $descricao=$text -split "- " | where {$_ -notlike "" } | select -Index 1
        $tabela | Add-Member -MemberType noteproperty -Name data -Value $data
        $tabela | Add-Member -MemberType noteproperty -Name hora -Value $hora
        $tabela | Add-Member -MemberType noteproperty -Name servidor -Value $servidor
        $tabela | Add-Member -MemberType noteproperty -Name instancia -Value $instancia
        $tabela | Add-Member -MemberType noteproperty -Name script -Value $script
        $tabela | Add-Member -MemberType noteproperty -Name processo -Value $processo
        $tabela | Add-Member -MemberType noteproperty -Name descricao -Value $descricao
        $arquivo += $tabela
    }
    #$teste=@()
    $nume_processo=$arquivo | where {$_.script -match "faz-backup" }  | select -Unique processo | select -ExpandProperty processo
    if (! ( test-path .\tamanho_backup.txt)){
        $exc_arquivo=@()
    } else {
        $exc_arquivo=@()
        $exc_arquivo1=Import-Csv .\tamanho_backup.txt
        $exc_arquivo += $exc_arquivo1
        
    }    
    foreach ($tex_processo in $nume_processo ){        
        $valida=$arquivo | where {$_.processo -match $tex_processo }
        $data_inicio= get-date ( $valida | select -Index 0 | select @{name="data";expression={$_.data +" "+ $_.hora}}).data
        $data_final= get-date ( $valida | select -last 1 | select @{name="data";expression={$_.data +" "+ $_.hora}}).data
        $minutos_gastos=(($data_final - $data_inicio).TotalMinutes | select @{name="tempo";expression={"{0:N3}" -f $_ }}).tempo
        $rc=(($valida | select -last 1).descricao  -split " " | where {$_ -match "RC="}).split("=") | select -Last 1
        $tabela2=New-Object psobject
        $tabela2 | Add-Member -MemberType noteproperty -Name data_inicio -value $data_inicio
        $tabela2 | Add-Member -MemberType noteproperty -Name data_final -Value $data_final
        $tabela2 | Add-Member -MemberType noteproperty -Name minutos_gasto -Value $minutos_gastos
        $tabela2 | Add-Member -MemberType noteproperty -Name rc -Value $rc
        $campo_1=$valida | select -Index 0
        $tabela2 | Add-Member -MemberType noteproperty -Name servidor -Value $campo_1.servidor
        $tabela2 | Add-Member -MemberType noteproperty -Name instancia -Value $campo_1.instancia
        $tabela2 | Add-Member -MemberType noteproperty -Name processo -Value $campo_1.processo
        $script_executado= ($arquivo | where {$_.processo -match $tex_processo -and $_.descricao -match "iniciando"} | select -index 2 | select script).script
        if (${script_executado}){
            $tabela2 | Add-Member -MemberType noteproperty -Name script -Value $script_executado
        }else{
           $tabela2 | Add-Member -MemberType noteproperty -Name script -Value 0 
        }
        $sdia=(Get-Date ${data_inicio} -f "D").split(",") | select -Index 0
        $tabela2 | Add-Member -MemberType noteproperty -Name dia_semana -Value $sdia
        $campo=$arquivo | where {$_.processo -match $tex_processo -and $_.script -match "netbackup"}| select -Index 0
        if(${campo}){
            $text_descricao=$campo.descricao -split " " | where {$_ -match "/" -or $_ -match "="}
            $local=$text_descricao | select -Index 0
            $tamanho_bkp=($text_descricao | select -Index 1 ).split("=")
            $politica_bkp=($text_descricao | select -Index 2 ).split("=")
            $schedule_bkp=($text_descricao | select -Index 3 ).split("=")
            if($tamanho_bkp[1] | where {$_ -match "G"}){
                [float]$tama_bkp=($tamanho_bkp[1] -split "G" | select -Index 0).replace(",",".")
                $tama_bkp=$tama_bkp * 1kb
            }else{
                [float]$tama_bkp=($tamanho_bkp[1] -split "M" | select -Index 0).replace(",",".")
                $tama_bkp=$tama_bkp * 1
            }
            $tabela2 | Add-Member -MemberType noteproperty -Name Local -Value $local
            $tabela2 | Add-Member -MemberType noteproperty -Name Tamanho -Value $tama_bkp
            $tabela2 | Add-Member -MemberType noteproperty -Name politica -Value $politica_bkp[1]
            $tabela2 | Add-Member -MemberType noteproperty -Name schedule -Value $schedule_bkp[1]
        }else{
            $tabela2 | Add-Member -MemberType noteproperty -Name Local -Value 0
            $tabela2 | Add-Member -MemberType noteproperty -Name Tamanho -Value 0
            $tabela2 | Add-Member -MemberType noteproperty -Name politica -Value 0
            $tabela2 | Add-Member -MemberType noteproperty -Name schedule -Value 0

        }
        
        $exc_arquivo += $tabela2   
    }
    $exc_arquivo | Export-Csv .\tamanho_backup.txt  -Encoding UTF8    
}
$configura =  Get-Content .\servidores.txt
#servidores
$servidores=$configura 
#$caminho_excel_mesal="C:\scripts\tamanho_backups\tamanho_bkp.xlsx"
$caminho_excel_mesal="C:\spool\tamanho_bkp.xlsx"
foreach ($servidor in $servidores ){
    #data    
    $data1=$servidor -split(" ") | select -Index 1 
    $servidor=$servidor -split(" ") | select -Index 0   
    New-SshSession -ComputerName ${servidor} -Username oracle -Password ora@2011
    $excluir_date= Get-Date -f ("yyyyMMdd") 
    $arquivo_servidor=Invoke-SshCommand -ComputerName ${servidor} -Command  "ls /bkp/backup-2016* | grep -v $excluir_date"
    $arquivo_servidor = $arquivo_servidor -split "`n"
    if (! ( test-path .\lista_arquivos_servidor.txt)){
        $ar_lista_servidor=@()
    } else {
        $ar_lista_servidor=@()
        $ar_lista_servidor1=Import-Csv .\lista_arquivos_servidor.txt
        $ar_lista_servidor +=$ar_lista_servidor1
    }
    foreach ($arquivo_lista in $arquivo_servidor){
        if(!($ar_lista_servidor| where {$_.servidor -match $arquivo_lista })){
            $texto=Invoke-SshCommand -ComputerName ${servidor} -Command "less ${arquivo_lista}"
            $texto=$texto -split "`n"
            formato_arquivo "" 
            $atualizar= New-Object psobject
            $atualizar | Add-Member -MemberType noteproperty -Name servidor -Value ${arquivo_lista}
            $ar_lista_servidor += $atualizar
        }        
    }
    $ar_lista_servidor | Export-Csv .\lista_arquivos_servidor.txt    
}


Remove-SshSession -RemoveAll
criar_excel ""