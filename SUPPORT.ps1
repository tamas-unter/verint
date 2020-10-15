Function Unzip-AllRecursively{
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	$zips=gci *.zip;
	$i=0
	$zips|%{
		$i++
		Write-Progress -Activity "Unzipping" -Status "$($_.Name) ($i/$($zips.Count))" -PercentComplete ($i/$zips.Count*100);
		try{
			[System.IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, (pwd).Path);
			Remove-Item $_.FullName
		}
		catch{Write-Host "error unzipping $($_.Name)" }
	}
	Gci -Directory | %{pushd; cd $_; Unzip-AllRecursively; popd}
}
Function Unzip-Silos{
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	$zips=gci *.zip;
	$i=0
	$zips|%{
		$i++
		Write-Progress -Activity "Unzipping" -Status "$($_.Name) ($i/$($zips.Count))" -PercentComplete ($i/$zips.Count*100);
        $zipdir=$_.Name.Substring(0,$_.Name.LastIndexOf("."))
        mkdir -ErrorAction SilentlyContinue $zipdir
		try{
			[System.IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, ((pwd).Path+"\"+$zipdir));
			Remove-Item $_.FullName
		}
		catch{Write-Host "error unzipping $($_.Name)" }
	}
	Gci -Directory | %{pushd; cd $_; Unzip-Silos; popd}
}
Function Filter-ErrorsInLog{
    param($entry="WE")
    process{
        $_|sls "\|[$entry]\]"
    }
}
Function Collect-Hits{
    param($regex,$captions=@('value'))
    process{
        $_|sls $regex|%{
        # [ordered]$r
            $r=@{}
            $r.Add('time',($_.line.substring(20,30)|get-date))
            for($i=1;$i -le $captions.count;$i++){
                $r.Add($captions[$i-1],$_.matches[0].groups[$i].value)
            }
            [pscustomobject]$r
        }
    }
}
Function Get-WFOVersion{
	[xml]$x=gc (ls -recurse serverversioninformation.xml| select -first 1)
	$x.ServerVersionInformation|select majorVersion, minorVersion, featurePack, hfr|foreach{"$($_.majorversion).$($_.minorversion) FP$($_.featurepack) HFR$($_.HFR)"}
}
Function Get-ComponentVersion{
    param ($pattern="Integration Service|IPCapture|Archiver")
	[xml]$x=gc (ls -recurse serverversioninformation.xml| select -first 1)
    $x.ServerVersionInformation.Components.Component | Where Name -Match $pattern
}

Function Select-TimeRange{
    param($pattern="*.log", $start,$end="")
    if($end -EQ ""){
        $dates=ls $pattern|select LastWriteTime|get-date
        $h,$m,$s=$start -split ":"
        if($s -EQ $null){
            $time=get-date $dates[[int]($dates.Count /2)] -Hour $h -Minute $m -Second 0
        } else{
            $time=get-date $dates[[int]($dates.Count /2)] -Hour $h -Minute $m -Second $s
        }
        ls $pattern|where LastWriteTime -gt $time |sort LastWriteTime |select -first 1
    }
    <# what about end???#>
} 
function Get-Timestamp{
param($line)
    get-date $line.Substring(20,23)
}
function Get-Alarms{
    process{
        $_|Collect-Hits "raisealarm<([^>]+)> instance<([^>]+)>+ params<([^>]+)>" @("alarm","instance","params")
    }
}
function Get-RISHostname{
# from HFR5 of 15.2 this is different..  Recorder Integration Service <VERINAPP19> Compon
    process{
        $_|sls "staticmode.*hostname - (.+)$"|select -first 1|%{$_.matches[0].groups[1].value}
    }
}
<#
locate a config file
ls -Recurse -File Servers.xml
 necessary for select-xml
 $ns=@{srv=$x.ChildNodes[1].Attributes[0].'#text'}

#>
function Get-AssociatedRecorderIdsByRISHostName{
param($hostname)
    [xml]$x=gc (ls -Recurse -File Servers.xml)
    $ns=@{srv=$x.ChildNodes[1].Attributes[0].'#text'}
    $x|Select-Xml -Namespace $ns "//srv:Server[@Hostname='$hostname']/srv:ServerRole[@Name='INTEGRATION_FRAMEWORK']/srv:RoleAssociation/@Identifier"|select -ExpandProperty node|select -ExpandProperty '#text'
}
function Get-ServerByRoleId{
# bad xmpl found in 1093495!!!
begin{
    [xml]$x=gc (ls -Recurse -File Servers.xml)
    $ns=@{srv=$x.ChildNodes[1].Attributes[0].'#text'}
}
process{
    $x|Select-Xml -Namespace $ns "//srv:ServerRole[@Identifier=$_ and @Name='IP_RECORDER']/.."|select -ExpandProperty node    
    }
}

function Get-HostnameBySerial{
	param($serial)[xml]$s=gc $($env:impact360softwaredir)Conf\Cache\Servers.xml$ns=@{srv=$s.ChildNodes[1].Attributes[0].'#text'}$s|Select-Xml -Namespace $ns "//srv:Server[@SerialNumber=$serial]"|select -ExpandProperty Node|select -ExpandProperty HostName
}
function Get-HostnameByInum{
	param([string]$inum)
	Get-HostnameBySerial $inum.Substring(0,6)
}
function Get-PathFromInum{
	param([string]$inum,[string]$bufferdrive="g")
	"\\$(Get-HostnameByInum $inum)\$bufferdrive$\calls\$($inum.substring(0,6))\$($inum.substring(6,3))\$($inum.substring(9,2))\$($inum.substring(11,2))\$inum"
}

function Get-ProcessNames{
param ($file)
    cat $file|sls '^\[([^\|]+)\|'|%{$_.matches[0].groups[1].Value.Trim()}|select -Unique
}

function Extract-RISProcesses{
param($processes)
begin{
    New-Item -ErrorAction SilentlyContinue -ItemType Directory "_processes"
    if($processes.count -eq 0){$processes=Get-ProcessNames (ls *.log | select -first 1)}
#    $files=gci ".\int*.log"
}
process{
    foreach ($process in $processes){
        Write-Progress "Parsing RIS log file" -Status "exporting $process ($($processes.indexof($process)+1)/$($processes.count))" -PercentComplete (($processes.indexof($process)+1)/$processes.count * 100)
		gc $_|sls "^\[$process" | out-file -append -width 2000 -filepath "_processes\$process.log"
    }
}
}

function Extract-LineRange{
<#
works only with full paths..
#>
    param ($file, $start, $end)
    $output="{0}_$start-$end.{1}" -f $file.Split(".")
    $reader=[io.file]::OpenText($file)
    $writer=[io.file]::CreateText($output)
    $count=0;
    while ($reader.EndOfStream -ne $true){
        $line=$reader.ReadLine()
        $count++;
        if($count -lt $start){continue}
        else {
            if ($count -le $end){
                $writer.writeLine($line)
            }else{break}
        }
    }
    $reader.Close()
    $reader.Dispose()
    $writer.Close()
    $writer.Dispose()
}

function Get-RISRedundancyStatus{
# only works with primary RIS log
# useful filter: |where {$_.sender -eq "PrimaryIS" -and $_.state -ne "Master"}
process{
    $_| sls "heartbeat" -context 0,1 |% {@{
        'l1'=$_.line;'l2'=$_.context.postcontext
    }}|%{
        ($_.l1+$_.l2)|
        sls "^.{20}.(.{22}).*heartbeat<(\w+)#(\d+)>.*str<state> = (\w+)"|%{
            [pscustomobject]@{
                timestamp=(get-date ($_.matches[0].groups[1].value)); 
                sender=$_.matches[0].groups[2].value; 
                ds=$_.matches[0].groups[3].value;
                state=$_.matches[0].groups[4].value
            }
        }
    }
}
}

function List-SessionsForContact{
param ($cid)
process{ $_|
    Collect-Hits "\[session.*  (\w+)((?: final)?) connection <(\d+)>.*extension<(\d+), (\d+)>, (\d+).*SESSION<(\d{34}).*contact<($cid)" @("action", "final", "connection", "ds","ext","callid", "session","contact")|
    select time, action, ext, callid, session|ft
}
}

function Get-GapsInLogs{
param($threshold=100000)
begin{$prevdate=get-date}
process{
    foreach ($l in [System.IO.File]::ReadLines($_.fullname)){
        $tmpdate=($l.substring(20,23)|get-date);
        $diff=($tmpdate-$prevdate).ticks;
        if($diff -ge $threshold){
            [pscustomobject]@{timestamp=$prevdate;ticks=$diff}
        }
        $prevdate=$tmpdate
     }
}}                                                                                                                                         $tmpdate
