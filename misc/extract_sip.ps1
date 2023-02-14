function Extract-SipMessages{
begin{
<#
    states
    0: not needed
    1: first line of the message
    2: message content
#>
    $state=0
    $result=@()
    $r=@{}
}
process{
    if($state -eq 0){
        if($_ -match "\[ProxySipLi"){
            $state=1
            $r=@{}
            $r=@{time=Get-TimeStamp $_;
                msg=$_}
        }
    }else{
        if($_ -notmatch "^>$"){
            $state=2
            $r.msg+="`n"+$_
            if($_ -match "Call-ID: (.+)$"){
                $r.callid=$Matches[1]
            }
        } else { 
            $state=0
            $result+=[pscustomobject]$r
        }
        
    }
}
end{
    $result
}
}

function Get-Timestamp{
param($line)
    get-date $line.Substring(20,23)
}
<#usage:

$sip=gc int*.log|Extract-SipMessages
$sip | where callid -EQ ...
#>