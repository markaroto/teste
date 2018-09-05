#funçao para criação da tabela mensal
function excel_mesal(){
	#Variavel referente linha do excel
    $line=1
	#Variavel referente teste existencia da planilha.
    $plan_existe=0
	#Verificar se o excel já existe
    if ((Test-Path  $caminho_excel_mesal )){
		#abertura do arquivo excel
        $objWorkbook=$objExcel.Workbooks.Open($caminho_excel_mesal)      
        if ($objWorkbook.Worksheets | where{$_.name -like $dia}){              
            #($objWorkbook.Worksheets | where {$_.name -like $dia}).delete()
            #modificar o valor plan 1, pois planilha já foi feita.
			$plan_existe=1
        }else{
		   #criar nova planilha no final de todas
           $plan=$objWorkbook.Worksheets.add([System.Reflection.Missing]::Value,($objExcel.Worksheets | where {$_.name -lt $dia} | select -Last 1))
           #$plan = $objWorkbook.Worksheets.add() 
        }               
    }else{
		#criar o arquivo de excel
        $objWorkbook = $objExcel.Workbooks.Add()
        #seleciona a planilha ativa
		$plan=$objWorkbook.ActiveSheet                 
    }	
    if ($plan_existe -eq 0 ){
        #altera nome planilha para o dia
		$plan.name=$dia
		#ponteiro das cells da planilha
        $cells1 = $plan.Cells
        $cells2 = $plan.Cells
		#loop para a leitura dos arquivos do linux
        foreach ($v_arquivo in  $caminho_txt){      
            #solicitação da informação texto do arquivo no linux para arquivo para uma variavel windows.
			$arquivo_full=Invoke-SshCommand -ComputerName lxornfs -Command "cat $v_arquivo "
			#limpeza da tela
            Clear-Host
			#Organização em linha para texto recebido do linux
            $arquivo_full=$arquivo_full -split "`n"  | where {$_ -notlike ""}
            if ($line -eq 1){
				#criação do titulo do arquivo
                $titulo_base=$arquivo_full[1]
				#loop criar o titulo arquivo
                for($x=2;$x -le 7;$x++){
                    $cells1.item($line,$x)=(($titulo_base ) -split " " | where {$_ -notlike ""} | select -Index $x)
                }
            }
			#criar um campo com hora dos arquivo
            $hora_excel=(($arquivo_full[1] -split " " | where {$_ -notlike ""} | select -Index 0) -split ":" | select -Index 0)
            #formatação do texto da hora
			$hora_excel=echo "${hora_excel}h"
            #teste do diretorio
			if($nfs5 -eq 1 ){
                $var_arquivo=$arquivo_full | where {$_ -match "Average"}
            }else{       
                $var_arquivo=$arquivo_full | where {$_ -match "Mé"}
            }
			#adiciona uma  mais na variavel linha
            $line++
			#preenchimento do excel
            $cells2.item(${line},1)=($hora_excel)
            for($x=2;$x -le 7;$x++){
                $cells2.item(${line},$x)=(($var_arquivo ) -split " " | where {$_ -notlike ""} | select -Index $x).replace(".",",") 
            }    
        }
		#criação do grafico
        $chart1 = $plan.Shapes.AddChart().Chart
        #modelo do grafico
		$chart1.chartType = $xlChart::xlColumnStacked100
        #cor da linha grafico
		$chart1.FullSeriesCollection(1).format.fill.ForeColor.RGB=12611584
        $chart1.FullSeriesCollection(2).format.fill.ForeColor.RGB=10498160
        $chart1.FullSeriesCollection(3).format.fill.ForeColor.RGB=5287936
        $chart1.FullSeriesCollection(4).format.fill.ForeColor.RGB=49407
        $chart1.FullSeriesCollection(5).format.fill.ForeColor.RGB=255
        $chart1.FullSeriesCollection(6).format.fill.ForeColor.RGB=15652797
        #tamanho do grafico
		$chart1.ChartArea.Height=368.875118110236
        $chart1.ChartArea.Left=382.5
        $chart1.ChartArea.Top=52.874881744384
        $chart1.ChartArea.Width=761.5
        #habilitar o titulo no grafico
		$chart1.SetElement(2)
		#Texto do titulo
        $chart1.ChartTitle.Text="Consumo médio cpu ${guest}"
        #end grafico
		#adiciona uma  mais na variavel linha
        $line++
		#Na ultima linha criando a media
        $cells2.item(${line},1)="Media" 
        $cells2.item(${line},2)='=MÉDIA(B2:B25)'
        $cells2.item(${line},3)='=MÉDIA(c2:c25)'
        $cells2.item(${line},4)='=MÉDIA(d2:d25)'
        $cells2.item(${line},5)='=MÉDIA(e2:e25)'
        $cells2.item(${line},6)='=MÉDIA(f2:f25)'
        $cells2.item(${line},7)='=MÉDIA(g2:g25)'
        #save arquivo
        
    }
	#salvar o arquivo 
    $objWorkbook.Saveas($caminho_excel_mesal)
	#fecha o arquivo
    $objWorkbook.Close()
}
#criar a planilha por hora planilha
function db_arquivo(){
	#criar variavel teste planilha
    $plan_existe=0
	#criar variavel com hora
    $name_plan_hora=($arquivo_full[4] -split ":" | select -Index 0)
    #verficar a existencia do arquivo
	if ((Test-Path  $caminho_excel )){
		#abertura do arquivo
         $objWorkbook=$objExcel.Workbooks.Open($caminho_excel)        
         if ($objWorkbook.Worksheets | where{$_.name -like $name_plan_hora}){
            #($objWorkbook.Worksheets | where {$_.name -like $name_plan_hora}).delete()
            #modificar o valor plan 1, pois planilha já foi feita.
			$plan_existe=1            
         }else{
			#criar uma nova planilha
            $plan = $objWorkbook.Worksheets.add()
         }                 
    }else{
		#criar um novo arquivo
        $objWorkbook = $objExcel.Workbooks.Add()
		#utilizar a planilha ativa 
        $plan = $objWorkbook.ActiveSheet        
    }
    
    if($plan_existe -eq 0 ){
	#altera o nome da planilha
    $plan.name=$name_plan_hora    
    #$plan.name=($arquivo_full_base[1] -split ":" | select -Index 0)
    #criar o titulo
	$titulo_base=$arquivo_full[1]       
    #configurar texto nas variaveis 
	$db_txt=$arquivo_full -split "`n" | where {$_ -notlike ""}
    $db_txt=$db_txt | select -Skip 1
    $db_titulo=$db_txt | select -Index 0
    $db_titulo=($db_titulo -split " ") -split "%" | where {$_ -notlike "" -and $_ -notmatch "CPU"}    
    #loop para criação base dados
	foreach ($db_variavel in  $db_txt ){
        #criaçao de objecto ps
		$tabela=New-Object psobject
        $tem_var=$db_variavel -split " "| where {$_ -notlike "" -and $_ -notlike "all"}
        $tabela | Add-Member -MemberType noteproperty -Name hora -Value $tem_var[0]
        for($x=1;$x -le 6;$x++){        
            $tabela | Add-Member -MemberType noteproperty -Name $db_titulo[$x] -Value $tem_var[$x]
        }
        $arquivo_db = $arquivo_db +$tabela         
        
    }
    $hora_media=$arquivo_db | where {$_.hora -like "*:01"}| select hora
    
    $b=1
    
    for($x=0;$x -le 6;$x++){
        $y=$x+1
        $plan.Cells.item($b,$y)=$db_titulo | select -Index $x
    }    
    $b++
    $arquivo_db_media=$arquivo_db | select -Skip 0  
    foreach ($v_minuto in $hora_media){
        $tabela=New-Object psobject
        $v_minuto=($v_minuto.hora -split ":" | select -Index 0,1 ) -join ":"
        $v_segundo=$arquivo_db_media | where {$_.hora -like "${v_minuto}:*"}
        [float]$tota_user=0
        [float]$tota_nice=0
        [float]$tota_system=0
        [float]$tota_iowait=0
        [float]$tota_steal=0
        [float]$tota_idle=0
        $count=0        
        foreach ($v_hora in $v_segundo){
             $tota_user+=[float]($v_hora.user -replace(",","."))
             $tota_nice+=[float]($v_hora.nice -replace(",",".")) 
             $tota_system+=[float]($v_hora.system -replace(",",".")) 
             $tota_iowait+=[float]($v_hora.iowait -replace(",",".")) 
             $tota_steal+=[float]($v_hora.steal -replace(",","."))
             $tota_idle+=[float]($v_hora.idle -replace(",","."))
             $count++;          
        }         
        
        $tota_user=$tota_user/$count
        [float]$tota_nice=$tota_nice/[float]$count 
        [float]$tota_system=[float]$tota_system/[float]$count
        [float]$tota_iowait=[float]$tota_iowait/[float]$count
        [float]$tota_steal=[float]$tota_steal/[float]$count
        [float]$tota_idle=[float]$tota_idle/[float]$count
        
        
        $plan.Cells.item($b,1)=$v_minuto
        $plan.Cells.item($b,2)=$tota_user #  -replace(",",".")
        $plan.Cells.item($b,3)=$tota_nice # -replace(",",".")
        $plan.Cells.item($b,4)=$tota_system # -replace(",",".")
        $plan.Cells.item($b,5)=$tota_iowait # -replace(",",".")
        $plan.Cells.item($b,6)=$tota_steal #-replace(",",".")
        $plan.Cells.item($b,7)=$tota_idle # -replace(",",".")
        
        $b++      
        
    }
    #grafico
    $chart1 = $plan.Shapes.AddChart().Chart
    $chart1.chartType = $xlChart::xlColumnStacked100
    $chart1.FullSeriesCollection(1).format.fill.ForeColor.RGB=12611584
    $chart1.FullSeriesCollection(2).format.fill.ForeColor.RGB=10498160
    $chart1.FullSeriesCollection(3).format.fill.ForeColor.RGB=5287936
    $chart1.FullSeriesCollection(4).format.fill.ForeColor.RGB=49407
    $chart1.FullSeriesCollection(5).format.fill.ForeColor.RGB=255
    $chart1.FullSeriesCollection(6).format.fill.ForeColor.RGB=15652797
    $chart1.ChartArea.Height=368.875118110236
    $chart1.ChartArea.Left=382.5
    $chart1.ChartArea.Top=52.874881744384
    $chart1.ChartArea.Width=761.5
    $chart1.SetElement(2)
    $chart1.ChartTitle.Text="Consumo médio cpu ${guest}"
        
    
   
    $arquivo_media_final=$arquivo_db | select -Last 1  
    $media_mensal = $media_mensal + $arquivo_media_final
    
    
    
    
    
    $plan.Cells.item($b,1)="Media"
    $plan.Cells.item($b,2)=$arquivo_media_final.user  #-replace(",",".")
    $plan.Cells.item($b,3)=$arquivo_media_final.nice  #-replace(",",".")
    $plan.Cells.item($b,4)=$arquivo_media_final.system # -replace(",",".")
    $plan.Cells.item($b,5)=$arquivo_media_final.iowait #-replace(",",".")
    $plan.Cells.item($b,6)=$arquivo_media_final.steal #-replace(",",".")
    $plan.Cells.item($b,7)=$arquivo_media_final.idle  #-replace(",",".")
    
    
    }
    $objWorkbook.Saveas($caminho_excel)
    $objWorkbook.Close()
       
}
$objExcel = New-object -COMobject excel.application
$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
$objExcel.Visible = $True
$objExcel.DisplayAlerts = $false
$diretorio="c:\excel_sar"
$arquivo_servidores="$diretorio\servidores.txt"
$diretorio_arquivo_servidor="\\copanet04-nas\spin\dvsu\CONFIGURACAO\ORACLE\MONCPUIO\historico"
#$diretorio_arquivo_servidor=$diretorio
Import-Module SSH-Sessions
$codigo="67;97;115;97;64;49;57;57;48;0"
$codigo= $codigo -split ";"
$codigo2="109;97;114;99;111;115;46;110;101;120;97;0"
$codigo2= $codigo2 -split ";"
for($x=0;$x -le 9;$x++){ 
    $resultado+=[char][byte]$codigo[$x]
}
for($x=0;$x -le 11;$x++){ 
    $resultado2+=[char][byte]$codigo2[$x]
}
New-SshSession -ComputerName lxornfs -Username $resultado2 -Password $resultado
if( !(Test-Path  $arquivo_servidores )){
    $r_server=Read-Host "Digite nome do servidores"
    mkdir $diretorio
    ($r_server).ToUpper() -split " " | Out-File "$arquivo_servidores"
}
for ($r=1;$r -ge 1;$r--){
    $r_server=Get-Content $arquivo_servidores
    $r_server=$r_server -split "`n" | where { $_ -notlike ""}    
    $tempo=get-date 
    $dia=(($tempo).adddays(-${r})).tostring('dd') 
    $mes=(($tempo).adddays(-${r})).tostring('MM') 
    $ano=(($tempo).adddays(-${r})).tostring('yy')
    $semana=(($tempo).adddays(-${r}) | select -Property datetime).datetime -split  "," | select -Index 0
    $arquivo_padrao="sar.d${ano}${mes}${dia}.h"
    foreach ($guest in $r_server ){
        if(!(Test-Path "${diretorio_arquivo_servidor}\${guest}")){
            mkdir "${diretorio_arquivo_servidor}\${guest}"
        }    
        $caminho_excel="${diretorio_arquivo_servidor}\${guest}\d${ano}${mes}${dia}-${semana}.xlsx"    
        $caminho_linux="/NFS6/${guest}/${arquivo_padrao}*"
        $caminho_txt=Invoke-SshCommand -ComputerName lxornfs -Command "ls -lr $caminho_linux" 
        $nfs5=0       
        if (($caminho_txt | where {$_ -match "No such file"})){
            $caminho_linux="/NFS5/${guest}/${arquivo_padrao}*"
            $caminho_txt=Invoke-SshCommand -ComputerName lxornfs -Command "ls -lr $caminho_linux"
            $nfs5=1 
        }
        if (!($caminho_txt | where {$_ -match "No such file"})){
            $caminho_txt=(($caminho_txt -split " ") -split "-" | where {$_ -notlike ""  -and $_ -notmatch "No such file"} | where {$_ -match "NFS6" -or $_ -match "NFS5"})
            $media_mensal=@()
            foreach ( $v_arquivo in  $caminho_txt){
                $caminho_linux="${v_arquivo}"                               
                $arquivo_full=Invoke-SshCommand -ComputerName lxornfs -Command "cat $caminho_linux "
                
                Clear-Host 
                $arquivo_full=$arquivo_full -split "`n"  | where {$_ -notlike ""} 
                $arquivo_db=@()
                db_arquivo "" 
                Clear-Host
                echo $media_mensal
           
            }        
            if($nfs5 -eq 1 ){
                $caminho_linux="/NFS5/${guest}/${arquivo_padrao}*"
            }else{
                $caminho_linux="/NFS6/${guest}/${arquivo_padrao}*"
            }
            $caminho_txt=Invoke-SshCommand -ComputerName lxornfs -Command "ls -l $caminho_linux" 
            $caminho_txt=(($caminho_txt -split " ") -split "-" | where {$_ -notlike ""  -and $_ -notmatch "No such file"} | where {$_ -match "NFS6" -or $_ -match "NFS5"})
            $caminho_excel_mesal="${diretorio_arquivo_servidor}\${guest}\${ano}${mes}-${guest}.xlsx"
            $todas_media=$todas_media -split "`n"| where {$_ -notlike ""}
            excel_mesal ""    
              
            
        }  
    }
}
$objExcel.Quit()