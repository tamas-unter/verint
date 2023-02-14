$extension="62735"
gc .\SIP.txt| 
foreach{
	if($_ -match "[:=]$extension"){
		if($_ -match "Request"){$_|sls "\[.*(.)\] (.{29})  Request<([^>]+)> URI<([^>]+)> Message<(.*)>$"|
		foreach{"_"*100+"`n"+$_.matches[0].groups[2].Value+" "+$_.matches[0].groups[4]+"`n"+"_"*100+"`n"+$_.matches[0].groups[5].Value.Replace("??","`n")}
		}else {if($_ -match "SIP Response"){
			$_|sls "\[.*(.)\] (.{29})  SIP Response<(.*)>$"|
			foreach{$_.matches[0].groups[2].Value+"`n"+"="*100+"`n"+$_.matches[0].groups[3].Value.Replace("??","`n")}
		}}
	}
} |
Out-File ".\SIP_$extension.txt"
