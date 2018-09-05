function historico(){
    $historico_caminho="C:\netbackup\historico.csv"
    if(! (Test-Path $historico_caminho)){
        $arquivo | Export-Csv $historico_caminho
    }else{
        $temp_arquivo= Import-Csv $historico_caminho
        $temp=@()
        $tem=$temp_arquivo + $arquivo
        $tem | Export-Csv $historico_caminho
    }
}


function enviar_email(){
    $assunto="Monitoria poll fitas $politica_email com ocupaçã ocupação de acima 70%"
    $mensagem=$arquivo | where { $_.pool -match $politica_email }   | Out-String
    $mensagem=$mensagem.replace("tamanho","ESPAÇO OCUPADO GB")
    $mensagem=$mensagem.replace("validade","EXPIRAÇÃO")
    $mensagem=$mensagem.replace("Tamanho_maximo","CAPACIDADE MÉDIA GB")
    $mensagem=$mensagem.replace("pc_ocupado","PCT OCUPADO")
    $mensagem=$mensagem.replace("total_dias","TOTAL DIAS UTILIZAÇÃO")
    $smtp = new-object Net.Mail.SmtpClient("10.1.1.52","587") 
    $smtp.Credentials = new-object System.Net.NetworkCredential("dvsu.nexa@copasa.com.br", "nex@2015")
    $user_email="marcos.flavio@nexa.com.br,Mardem.barbosa@copasa.com.br,rogerio.notini@copasa.com.br,roziane.barroso@copasa.com.br"
    $email = new-object System.Net.Mail.MailMessage("dvsu.nexa@copasa.com.br",$user_email) 
    $email.IsBodyHtml = $false
    $email.Subject = $assunto
    $email.BodyEncoding=[System.Text.Encoding]::UTF8
    $email.Priority=2 
    $email.Body = $mensagem
    $smtp.Send($email) 
         
    #$out= New-Object -ComObject outlook.application
    #$email=$out.CreateItem(0)
    #$email.To="marcos.flavio@nexa.com.br" #Mardem.barbosa@copasa.com.br"
    #$email.Subject="Monitoria poll fitas $politica_email com ocupaçã ocupação de acima 70%"
    #$email.Body=$arquivo | where { $_.Politica -match $politica_email }   | Out-String
    #$email.Importance=2
    #$email.Send()

}

function grafico_crecimento (){
    $excel= New-Object -com excel.application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $list_politica=$arquivo | select -Unique pool
    if( Test-Path $grafico_excel){
        $file= $excel.Workbooks.Open($grafico_excel)
        $plan= $file.ActiveSheet
    }else{
        $file= $excel.Workbooks.Add()
        $plan= $file.ActiveSheet
        $coluna=2
        $plan.cells.item(1,1)="DATA DA COLETA"
        foreach ($pool in $list_politica){
            $plan.cells.item(1,$coluna)=$pool.pool 
            $coluna++           
        }
    }
    $linha=2
    $coluna=2
    foreach ($pool in $list_politica){
        $coluna=2
        while ($coluna -ne 0){
            $poll_atual=$plan.cells.item(1,$coluna).text
            if ( $pool.pool -eq $poll_atual ){
                $coluna=0
            }else{
                if ( $plan.cells.item(1,$coluna).text -eq ""){
                    $plan.cells.item(1,$coluna)=$pool.pool
                    $coluna=0
                }else{
                    $coluna++
                }
            }
        }
    }
    $coluna=2
    while($linha -ne 0){
        $data_exc=($arquivo | select -Unique data).data
        if ($plan.cells.item($linha,1).text -eq "" ){
            $plan.cells.item($linha,1)=get-date $data_exc -f ("MM/dd/yyyy")
            while($coluna -ne 0){
                if( $plan.cells.item(1,$coluna).text -eq "" ){
                    $coluna=0
                }else{
                    $poll_atual=$plan.cells.item(1,$coluna).text
                    $arquivo_total=$arquivo | where {$_.pool -like $poll_atual }
                    $tamanho_poll=0
                    foreach($pol_tamanho in $arquivo_total){
                        [float]$tamanho_temp=($pol_tamanho | select tamanho).tamanho
                        [float]$tamanho_poll= $tamanho_poll + $tamanho_temp 
                    }
                    $plan.cells.item($linha,$coluna)=$tamanho_poll
                    $coluna++
                }                
            }
            $linha=0
        }else{
            if ($plan.cells.item($linha,1).text -eq $data_exc ){
                $linha=0
            }else{
                $linha++            
            }
        }
        
    }
    $file.Saveas($grafico_excel)
    $file.Close()
    $excel.quit()       
        
}

function gravar_excel(){
    $excel= New-Object -com excel.application
    #$excel.Visible = $True
    $excel.DisplayAlerts = $false
    if( Test-Path $arquivo_excel){
        $file= $excel.Workbooks.Open($arquivo_excel)
		$plan=$file.Worksheets | Where-Object {$_.name -like "dados"}
        #$plan= $file.ActiveSheet
    }else{
        $file= $excel.Workbooks.Add()
        $plan= $file.ActiveSheet
        $plan.name="dados"
		$plan.cells.item(1,1)="DATA DA COLETA"
        $plan.cells.item(1,2)="HORA COLETA"
        $plan.cells.item(1,3)="POOL"
        $plan.cells.item(1,4)="MÍDIA"
        $plan.cells.item(1,5)="CAPACIDADE MÉDIA DE ARMAZENAMENTO (GB)"
        $plan.cells.item(1,6)="ESPAÇO DISPONÍVEL (GB)"
        $plan.cells.item(1,7)="ESPAÇO CUPADO (GB)"
        $plan.cells.item(1,8)="% OCUPADO"
        $plan.cells.item(1,9)="STATUS DA FITA"
        $plan.cells.item(1,10)="DATA INICIO UTILIZAÇÃO"
        $plan.cells.item(1,11)="DATA EXPIRAÇÃO"
        $plan.cells.item(1,12)="QTDE DIAS UTILIZAÇÃO" 
    }
    $sair=1
    while( $sair -ne 0){
        if (($plan.cells.item($sair,1).text) -eq ""){
            $media_loops=$arquivo 
            foreach ($media_loop in $media_loops){  
                $plan.cells.item(${sair},1)=get-date $media_loop.data -f ("MM/dd/yyyy")
                $plan.cells.item(${sair},2)=$media_loop.hora_atual
                $plan.cells.item(${sair},3)=$media_loop.pool
                $plan.cells.item(${sair},4)=$media_loop.Midia
                $plan.cells.item(${sair},5)=$media_loop.Tamanho_maximo
                $plan.cells.item(${sair},6)=$media_loop.Disponivel
                $plan.cells.item(${sair},7)=$media_loop.tamanho
                $plan.cells.item(${sair},8)=$media_loop.pc_ocupado
                $plan.cells.item(${sair},9)=$media_loop.Estatus_fita
                if ( $media_loop.incio_midia -ne $null){
                    $plan.cells.item(${sair},10)=get-date $media_loop.incio_midia -f ("MM/dd/yyyy HH:mm:ss")
                }
                if ( $media_loop.validade -ne $null){
                    $plan.cells.item(${sair},11)=get-date $media_loop.validade -f ("MM/dd/yyyy HH:mm:ss")
                }
                $plan.cells.item(${sair},12)=$media_loop.total_dias             
                $sair= $sair +1                                
            }
            $sair=0
        }else{
            $sair= $sair +1
        }        
    }
    $file.Saveas($arquivo_excel)
    $file.Close()
    $excel.quit()
    
    
}



#coleta dados netbackup
$arquivo_excel="C:\netbackup\politica.xlsx"
$arquivo_excel1="\\copanet04\dvsu\CONFIGURACAO\ORACLE\PROJETOS ORACLE\SEGURANÇA\ACOMPANHAMENTO DO BACKUP\Tabela de utilizacao Fitas por Pool.xlsx"
$grafico_excel="C:\netbackup\grafico.xlsx"
$grafico_excel1="\\copanet04\dvsu\CONFIGURACAO\ORACLE\PROJETOS ORACLE\SEGURANÇA\ACOMPANHAMENTO DO BACKUP\Grafico de utilizacao Fitas por Pool.xlsx"
$politicas=Get-Content C:\netbackup\politica.txt
$politicas= $politicas -split "`n" | where {$_ -notlike ""}
$arquivo=@()
foreach ($politica in $politicas){
    $caminho_volumgr="C:\Program Files\Veritas\Volmgr\bin\vmquery.exe"
    $dados=Invoke-Command -ComputerName bkpmaster -ScriptBlock { & $args[0] -h bkpmaster -pn $args[1] } -ArgumentList $caminho_volumgr,$politica
    #cd "C:\Program Files\Veritas\Volmgr\bin\"
    #$dados=.\vmquery.exe -h bkpmaster -pn $politica
    $medias=($dados | where {$_ -match "media ID"} ) -split(" ") | where{$_ -notlike "" -and $_ -notmatch "ID:" -and $_ -notmatch "media" }
    foreach ($media in $medias){
        $caminho_admincmd="C:\Program Files\Veritas\NetBackup\bin\admincmd\bpmedialist.exe"
        $media_info= Invoke-Command -ComputerName bkpmaster -ScriptBlock { & $args[0] -ev $args[1] } -ArgumentList $caminho_admincmd,$media
        #cd "C:\Program Files\Veritas\NetBackup\bin\admincmd\"
        #$media_info=.\bpmedialist.exe -ev $media             
        if ($? -eq $false){
            $banco= New-Object psobject
            $banco | Add-Member -MemberType noteproperty -Name pool -Value $politica
            $banco | Add-Member -MemberType noteproperty -Name Midia -Value $media
            $data_atual= (get-date).tostring("dd/MM/yyyy")
            $hora_atual=(get-date).tostring("HH:mm:ss")  
            $banco | Add-Member -MemberType noteproperty -Name data -value  $data_atual 
            $banco | Add-Member -MemberType noteproperty -Name hora_atual -value $hora_atual            
            $banco | Add-Member -MemberType noteproperty -Name tamanho -value  $null
            $banco | Add-Member -MemberType noteproperty -Name Estatus_fita -value  "Disponivel"            
            $banco | Add-Member -MemberType noteproperty -Name incio_midia -value  $null
            $banco | Add-Member -MemberType noteproperty -Name validade -value  $null
            $banco | Add-Member -MemberType noteproperty -Name Tamanho_maximo -value  $null
            $banco | Add-Member -MemberType noteproperty -Name Disponivel -value  $null
            $banco | Add-Member -MemberType noteproperty -Name pc_ocupado -value  $null            
            $banco | Add-Member -MemberType noteproperty -name total_dias -Value $null
            
                     
        }else{
            $incio_midia=($media_info[6] -split " " | where {$_ -notlike ""} | select -Index 3,4) -join " "
            $incio_midia=[datetime]$incio_midia
            $tamanho_midia=($media_info[6] -split " " | where {$_ -notlike ""} | select -Index 8)
            $tamanho_midia=([float]$tamanho_midia * 1kb) / 1gb
            $date=($media_info[7] -split " " | where {$_ -notlike ""} | select -Index 1,2)
            $estatus_midia=($media_info[7] -split " " | where {$_ -notlike ""} | select -Index 4) | where {$_ -match "FULL"}
            if ( $estatus_midia -eq $null){
                $estatus_midia=($media_info[7] -split " " | where {$_ -notlike ""} | select -Index 5) | where {$_ -match "FULL"}
            }
            $date=$date -join " "
            $date=[datetime]$date
            $incio_midia=[datetime]$incio_midia
            $data_atual= (get-date).tostring("dd/MM/yyyy")
            [float]$tamanho_medio='1536'
            $disponivel_espaco=$tamanho_medio-$tamanho_midia
            $pc_ocupado=100-(($disponivel_espaco*100)/$tamanho_medio)
            $total_tempo=($date - $incio_midia).days
            if($estatus_midia -eq $null){
                $estatus_midia="Ativa"
            }
            $hora_atual=(get-date).tostring("HH:mm:ss")
            $banco= New-Object psobject
            $banco | Add-Member -MemberType noteproperty -Name pool -Value $politica
            $banco | Add-Member -MemberType noteproperty -Name Midia -Value $media
            $banco | Add-Member -MemberType noteproperty -Name data -value  $data_atual
            $banco | Add-Member -MemberType noteproperty -Name tamanho -value  $tamanho_midia
            $banco | Add-Member -MemberType noteproperty -Name Estatus_fita -value  $estatus_midia            
            $banco | Add-Member -MemberType noteproperty -Name incio_midia -value  $incio_midia
            $banco | Add-Member -MemberType noteproperty -Name validade -value  $date
            $banco | Add-Member -MemberType noteproperty -Name Tamanho_maximo -value  $tamanho_medio
            $banco | Add-Member -MemberType noteproperty -Name Disponivel -value  $disponivel_espaco
            $banco | Add-Member -MemberType noteproperty -Name pc_ocupado -value  $pc_ocupado            
            $banco | Add-Member -MemberType noteproperty -name total_dias -Value $total_tempo
            $banco | Add-Member -MemberType noteproperty -Name hora_atual -value $hora_atual
            
            
        }
        $arquivo+=$banco           
    }
} 

gravar_excel ""

$polita_email=$arquivo | select -Unique pool
foreach ( $p_email in $polita_email ){
    $politica_email=${p_email}.pool
    $quantida_p=($arquivo | where { $_.pool -match $politica_email  -and $_.estatus_fita -notmatch "FULL" -and $_.pc_ocupado -le 70 } )
    if ( $quantida_p -eq $null) { 
       enviar_email ""
       echo ok 
    }
}

grafico_crecimento ""

historico ""




copy-item -path $arquivo_excel -Destination $arquivo_excel1  -Force
copy-item -path $grafico_excel -Destination $grafico_excel1  -Force