function main{
param ($directory="~\Downloads")
    $files=ls $directory\*.bin|sort
    foreach($file in $files){
        parse-loggerfile $file
        Write-Progress -Activity "parsing" -PercentComplete ($files.IndexOf($file) / $files.Count *100) -Status $file.Name 
    }
}

function parse-loggerfile{
param ($file)
    $b=[system.io.file]::ReadAllBytes($file)
    if(check-logfile $b){
        parse-data $b
    }
}

function check-logfile{
param ([byte[]]$b)
    if($b[0] -eq 0xe0 -and $b[1] -eq 0xc5 -and $b[2] -eq 0xea){
        return $true
    }
    return $false
}

function parse-data{
param ([byte[]]$b, $i=3)
    $mon,$d,$y,$h,$min=$b[$i..($i+4)]
    $timestamp=Get-Date -Year ($y+2000) -Month $mon -Day $d -Hour $h -Minute $min -Second 0
    $prevtimestamp=$timestamp
    $prevpower=0.0
    while (($i -lt $b.Length) -and ($b[$i] -ne 0xff)) {
        if (-not (check-logfile $b[($i+5)..($i+7)])){
            $volt1,$volt2,$amper1,$amper2,$factor=$b[($i+5)..($i+9)]
            $volt=($volt1*256+$volt2)/10
            $amper=($amper1*256+$amper2)/1000
            $factor/=100
            $power_real=$volt*$amper*[math]::Cos($factor)
            $power_apparent=$volt*$amper*$factor
            $power_reactive=$volt*$amper*[math]::Sin($factor)
            $energy=($power_real+$prevpower)/2*(($timestamp-$prevtimestamp).totalhours)

            if(($volt -lt 120) -or ($volt -gt 260) -or ($power_real -lt 0)){Write-Host "u=$volt p_real=$power_real $timestamp $file"}
            
            [pscustomobject]@{time=$timestamp;u=$volt;i=$amper;phi=$factor;power_real=$power_real;power_apparent=$power_apparent;power_reactive=$power_reactive;energy=$energy}
            $prevpower=$power_real
            $prevtimestamp=$timestamp
            $i+=10
            $timestamp=$timestamp.AddMinutes(1)
        }else{
            #Write-Host "broken at $i / $($b.Count) $file"
            
            $mon,$d,$y,$h,$min=$b[($i+8)..($i+12)]
            $timestamp=Get-Date -Year ($y+2000) -Month $mon -Day $d -Hour $h -Minute $min -Second 0
            $i+=8
            <# parse-data $b ($i+8); #>
            #break
        }
    }
}