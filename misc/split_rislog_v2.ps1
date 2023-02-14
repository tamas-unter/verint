New-Item -ErrorAction SilentlyContinue -ItemType Directory "_processes"
$processes="NGAConnect","Call","CallTracke","CiscoICMCi","Contact","Content","DataSource","ExtendedHi","ICallManag","IEMessage","Line","PortIPReco","ProxyRecor","ProxySipLi","Recordable","Recording","RequestedC","Session","SipCall","SipCallMan","SipObject","StaticMode","Statistics","StreamMana","StreamMatch","Workspace","RecorderSe"
$files=gci ".\int*.log"
$i=0 ;
Foreach ($file in $files){
	$i++;
	Write-Progress -Activity "Parsing $($file.Name) ($i of $($files.count))" -PercentComplete ($i/$($files.Count)*100)
	foreach ($process in $processes){
		gc $file|sls "^\[$process" | out-file -append -width 2000 -filepath "_processes\$process.log"
	}
}
