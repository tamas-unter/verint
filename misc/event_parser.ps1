<#
TESTED ONLY WITH AVAYA events
for Cisco ICM, there are a lot of errors in handle.array.. possibly not having the count
#>

function Get-RISProcess{
param($process)
begin{
#    New-Item -ErrorAction SilentlyContinue -ItemType Directory "_processes"
#    $files=gci ".\int*.log"
}
process{
		gc $_|sls "^\[$process" |select -expandproperty Line #| out-file -append -width 2000 -filepath "_processes\$process.log"
}
}
function Load-RISLogs{
    param($files)
    $global:s=$files|Get-RISProcess "IEMessage"
    $global:i=0
    $global:ident=New-Object System.Collections.Stack
    $global:ident.Push(0)
}
function Load-LogFile {
    Param($file)
    $global:s=gc $file
    $global:i=0
    $global:ident=New-Object System.Collections.Stack
    $global:ident.Push(0)
}  
function EOF{($global:i -ge ($global:s).count)}
function Ok-ToContinue{
	-not (EOF) -and ((get-ident($global:s[$global:i+1])) -gt $global:ident.Peek())
}
Function Get-Line{
	$global:s[$global:i]
}
Function Get-NextLine{
	$global:s[++($global:i)]
}
Function Proceed-Line{$global:i++}

Function Get-Ident{
	Param($l)
	If($l -eq $null) {return -1} else{
        if($l.length -eq 0) {return 1}
    }
	$ident=0;
	[int]$arrayident=0
	If(($l) -match "\[\d+\]:"){
		$arraystart=($l).indexof("]:")
		$arrayident=($l).substring($arraystart)|sls "\s(\w)" |%{$_.matches[0].groups[1].index}
	}
	$i=6;while(($i -ge 1) -and ($l.IndexOf(" "*2*$i) -ne 49)) {$i--}
	$arrayident+$i
}

Function Seek-NextEvent{
	if((get-line) -match "Dispatching Event"){return}
	while(-not (EOF) -and ((get-nextline) -notmatch "Dispatching Event")){}
}

Function Handle-Event{
    $ident=get-ident(get-line)
    $global:ident.Push($ident)
	$event= get-line|sls "<([^>]+)> --> <([^>]+)>"|%{@{"src"=$_.matches[0].groups[1].value;"dst"=$_.matches[0].groups[2].value}}
    $event+=@{"timestamp"=((Get-Line).substring(20,23)|Get-Date)}
	while(Ok-ToContinue) {Proceed-Line; $event+=handle-line}
    $next=$global:ident.Pop()
	$event
}

Function Handle-Line{
	$a=@{}
    $ident=get-ident(get-line)
    $global:ident.Push($ident) 
	If((get-line) -match "array"){$a+=Handle-Array}Else {
		if ((get-line) -match "folder"){$a+=Handle-Folder} Else{
			get-line|sls -AllMatches "<([^>]+)> = ([^;^$]+)"|select -ExpandProperty Matches|%{$a+=@{($_.groups[1].value)=($_.groups[2].value.trim())}}
			while(Ok-ToContinue){Proceed-Line; $a+=Handle-Line}
		}
	}
    $next=$global:ident.Pop()
	$a
}
Function Handle-Array{
	$name,$count=get-line|sls "Array<([^>]+)>.*\[(\d+)\]"|%{$_.matches[0].groups[1].value, $_.matches[0].groups[2].value}
	$a=new-object System.Object[] $count
	while(Ok-ToContinue){
		[int]$index=get-nextline|sls "\s\[(\d+)\]:"|%{$_.matches[0].groups[1].value}
		$a[$index]+=(handle-line)
	} 
	@{$name=$a}
}
Function Handle-folder{
	$o=@{}
	$name=get-line|sls "folder<([^>]+)>"|%{$_.matches[0].groups[1].value}
    do{
        Proceed-Line
		$o+=(handle-line)
	} while(Ok-ToContinue)
	@{$name=$o}
}

#Load-LogFile "C:\Users\tpasztor\Downloads\1081047\New folder (2)\_processes\IEMessage.log"

#$e=New-Object 'System.Collections.Generic.List[hashtable]'
#while(Ok-ToContinue ) { 
#Seek-NextEvent;$evt=(Handle-Event); $e.Add($evt);Proceed-Line}
#

if($false){
$e=New-Object 'System.Collections.Generic.List[hashtable]'
while(Ok-ToContinue ) {
    Write-Progress -Activity "${$evt.timestamp} ${$evt.src}-->${$evt.dst}" -PercentComplete ($Global:i*100/$Global:s.Count); 
    Seek-NextEvent; 
    $evt=(Handle-Event);$e.Add($evt);
    Proceed-Line
}
}

Function Dump-Object{
param($o,$prefix="")
    if ($o -is [Hashtable]){ Dump-Hashtable -prefix $prefix $o}
    else {if ($o -is [Array]){Dump-Array -prefix $prefix $o}
        else {'{0}' -f $o}
    }
}
Function Dump-Array{
param($a,$prefix)
    $r=$a|%{Dump-Object -prefix ("$prefix({0})" -f $a.indexOf($_) ) $_}
    "`n"+($r -join "`n")
}
Function Dump-HashTable{
param($ht,$prefix)
    $a=$ht.Keys|Sort|%{"$prefix."+'{0}={1}' -f $_, (Dump-Object -prefix "$prefix.$_" $ht[$_])}
    "`n"+($a -join "`n") 
}

<#
TODO- find an event having a specific string in ANY value.. roam the tree of hashtables and arrays

#>

Function Select-CallId{
param ($callid)
process{$_|
    where{
        ($_.event.transferredConnections -ne $null) -and ($_.event.transferredConnections[0].party.callid -eq $callid) -or
        ($_.event.conferenceConnections -ne $null) -and ($_.event.conferenceConnections[0].party.callid -eq $callid) -or
        ($_.event.establishedConnection.callid -eq $callid) -or
        ($_.event.connection.callid -eq $callid) -or
        ($_.event.droppedConnection.callid -eq $callid) -or
        ($_.event.initiatedConnection.callid -eq $callid) -or
        ($_.event.originatedConnection.callid -eq $callid) -or
        ($_.event.heldConnection.callid -eq $callid) -or
        ($_.event.retrievedConnection.callid -eq $callid) -or
        ($_.event.failedConnection.callid -eq $callid) -or
        ($_.event.queuedConnection.callid -eq $callid)
    }
}}

function Get-DeviceEvents{
param ($device)
process{ $_|
    sls "^.{31}(.{12}).*Str<deviceID> = $device ;" -context 10,22|%{
        [pscustomobject]@{
            time=$_.matches[0].groups[1].value;
            event=$_.context.precontext|
                sls "description> = ([A-Z_\-]+)"|%{$_.matches[0].groups[1].value}; 
            connection=($_.context.postcontext | 
                sls "Int<callID> = (\d+) ; Int<devIDType> = (\d+) ; Str<deviceID> = (\S+)"|
                select -first 1|%{[pscustomobject]@{
                    callid=$_.matches[0].groups[1].value;
                    devicetype=$_.matches[0].groups[2].value;
                    device=$_.matches[0].groups[3].value
                }}
            )
        }
    }
}}

function uzi {
    cd (Get-Clipboard)
    Unzip-AllRecursively
}
