# Split all RIS log files in the current folder along the process type and save it into a new subdirectory

New-Item -ErrorAction SilentlyContinue -ItemType Directory "_processes"
$c=gc ".\int*.log"
$processes="NGAConnect","Call","CallTracke","CiscoICMCi","Contact","Content","DataSource","ExtendedHi","ICallManag","IEMessage","Line","PortIPReco","ProxyRecor","ProxySipLi","Recordable","Recording","RequestedC","Session","SipCall","SipCallMan","SipObject","StaticMode","Statistics","StreamMana","StreamMatch","Workspace","RecorderSe"

Foreach($proc in $processes){
	$c|sls "^\[$proc" | out-file -width 2000 -filepath "_processes\$proc.log"
}
