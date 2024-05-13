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
Function uzi{
    cd (Get-Clipboard)
    Unzip-AllRecursively
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
            $r.Add('time',(Get-Timestamp $_.line))
            for($i=1;$i -le $captions.count;$i++){
                $r.Add($captions[$i-1],$_.matches[0].groups[$i].value)
            }
            [pscustomobject]$r
        }
    }
}
Set-Alias clh Collect-Hits
Set-Alias np 'C:\Program Files\Notepad++\notepad++.exe'
# ---- functions using the serverversioninformation.xml - need a Conf folder somewhere 
Function Get-WFOVersion{
	[xml]$x=gc (ls -recurse serverversioninformation.xml| select -first 1)
	$x.ServerVersionInformation|select majorVersion, minorVersion, featurePack, hfr|foreach{"$($_.majorversion).$($_.minorversion) FP$($_.featurepack) HFR$($_.HFR)"}
}
Function Get-ComponentVersion{
    param ($pattern="Integration Service|IPCapture|Archiver")
	[xml]$x=gc (ls -recurse serverversioninformation.xml| select -first 1)
    $x.ServerVersionInformation.Components.Component | Where Name -Match $pattern
}

# ----- functions using ANY log file series
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

function Extract-LineRange{
<#
works only with full paths..
TODO: ls | select fullpath
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

<#
.EXAMPLES
# get the top 10 issues
ls int*.log|Get-GapsInLogs|sort ticks -Descending | select -first 10
#>
function Get-GapsInLogs{
param($threshold=1000000) #100 ms
begin{$prevdate=get-date}
process{
    foreach ($l in [System.IO.File]::ReadLines($_.fullname)){
        try{
            $tmpdate=($l.substring(20,23)|get-date -ErrorAction SilentlyContinue);
            $diff=($tmpdate-$prevdate).ticks;
            if($diff -ge $threshold){
                [pscustomobject]@{timestamp=$prevdate;ticks=$diff}
            }
        }catch{} finally{
            $prevdate=$tmpdate
        }
     }
}}         



# --- internal
function Get-Timestamp{
param($line)
    get-date $line.Substring(20,23)
}

# --- RIS log functions - expects (ls integrationservice*.xml) as input

function Get-RisAlarms{
    process{
        $_|Collect-Hits "raisealarm<([^>]+)> instance<([^>]+)>+ params<([^>]+)>" @("alarm","instance","params")
    }
}

function Get-Alarms{
    process{
        $_|clh "raisealarm alarmname:([^:]+):(.*)$" @("alarm", "details")|
            %{
                $d=$_.details.split(":");
                [pscustomobject]@{
                    "time"=$_.time;
                    "alarm"=$_.alarm;
                    $d[0]=$d[1];
                    $d[2]=$d[3];
                    $d[4]=$d[5,7,9,11]
                }
             }#|ft
    }
}

function Get-TaggedAttributes{
    process{
     $a=@();
     $_|sls "\[nga.*6 tag (.*)$"|
        %{
            $b=@{};
            $_.matches[0].groups[1].value -split ": "|%{
                $c=$_ -split ":";
                # 003a in the timestamps
                $b.add($c[0],$c[1].replace("%003a",":"));
            }
            $a+=[pscustomobject]$b
            }

    $a
    }
}

function Get-RISHostname{
# from HFR5 of 15.2 this is different..  Recorder Integration Service <VERINAPP19> Compon
    process{
        $_|sls "staticmode.*hostname - (.+)$"|select -first 1|%{$_.matches[0].groups[1].value}
    }
}

function Get-RisVersion{
# this is HFR9
    ls -Recurse int*.log|sls "\[VersionMai.*Core Service v([0-9\.]+)"|select -first 1 @{l="v";e={$_.matches[0].groups[1].value} }|select -ExpandProperty v
# TODO: pre-HFR8
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

function Get-CTIEvents{
    param($callids)
    $callids|%{
        (
            (ls int*.log|
                sls "iemess.*callid> = $_ " -context 30,10|
                % {$_.context.precontext,$_.line,$_.context.postcontext}
            )|% {
                $_|%{$_}}
        )|
        where {$_ -match "^\[iemess"}
    }|Out-File calls.txt
}

function Get-OSVEventsForCall{
	param($callid);
	ls int*|
		sls "Str<callID> = $callid" -context 5,0|%{
			$_.context.precontext|
				sls "eventname> = (.+)Event"|%{
					[pscustomobject]@{
						t=$_.line.substring(31,12);
						e=$_.matches[0].groups[1].value
					}
				}
		}
}


function Get-RISFailovers{
    ls int*.log|Get-RISRedundancyStatus|where {$_.sender -eq "PrimaryIS" -and $_.state -ne "Master"}
}

function Get-RestartTime{
    ls -Recurse *.log|sls "\*{3}"|where linenumber -gt 1
}

function List-SessionsForContact{
param ($cid)
process{ $_|
    Collect-Hits "\[session.*  (\w+)((?: final)?) connection <(\d+)>.*extension<(\d+), (\d+)>, (\d+).*SESSION<(\d{34}).*contact<($cid)" @("action", "final", "connection", "ds","ext","callid", "session","contact")|
    select time, action, ext, callid, session|ft
}
}

function Get-SipMessages{
process{
    $_|
        sls "\[(SipObject.*sent|ProxySipLi.*receive)" -context 0,50|
        %{
            "`n";$_.line;($_.context.postcontext|sls "^[^\[]")
        }
}
}

<#
locate a config file
ls -Recurse -File Servers.xml
 necessary for select-xml
 $ns=@{srv=$x.ChildNodes[1].Attributes[0].'#text'}

#>

<# ------- functions using Servers.xml ---------------- #>
function Get-Node{
param([string]$xpath='/') 
process{
    $_|
        Select-Xml -Namespace @{'svr'='http://www.verint.com/xml/em/servers'} $xpath|
        Select -ExpandProperty Node
}
}
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
	param($serial)# TODO: path[xml]$s=gc (ls -Recurse -File Servers.xml)#####gc $($env:impact360softwaredir)Conf\Cache\Servers.xml$ns=@{srv=$s.ChildNodes[1].Attributes[0].'#text'}$s|Select-Xml -Namespace $ns "//srv:Server[@SerialNumber=$serial]"|select -ExpandProperty Node|select -ExpandProperty HostName
}

# --- inum functions
function Get-HostnameByInum{
	param([string]$inum)
	Get-HostnameBySerial $inum.Substring(0,6)
}

function Get-PathFromInum{
#TODO: extract the buffer path from IPCaptureConfig.xml
	param([string]$inum,[string]$bufferdrive="g")
	"\\$(Get-HostnameByInum $inum)\$bufferdrive$\calls\$($inum.substring(0,6))\$($inum.substring(6,3))\$($inum.substring(9,2))\$($inum.substring(11,2))\$inum"
}


# --- operational functions

function Parse-VersionReport{
param ($comp = "Recorder")
    cat .\Version*.csv|
        select -Skip 2|
        ConvertFrom-Csv|
        where component -match $comp
}                                                                                                                                
#$tmpdate


function Start-NotepadPlusPlus{
	param ($excerpt=(get-clipboard)) 
	$p=$excerpt.Split(":"); 
	np "-n$($p[1])" $p[0]
}
Set-Alias n Start-NotepadPlusPlus

function Find-FirstOccurrences{
    param ($searchTerm=(get-clipboard))
    process{
    $_|
        sls $searchTerm|
        select Path, LineNumber|
        group Path|%{
            $p=$_.group[0];
            np "-n$($p.LineNumber)" $p.Path 
        }
    }
}
Set-Alias ff Find-FirstOccurrences


Function Get-VersionReport{ 
    cat .\version*.csv |select -skip 2 |ConvertFrom-Csv
}


# --- 3rd party regex highlighting 

function Select-ColorString {
     <#
    .SYNOPSIS

    Find the matches in a given content by the pattern and write the matches in color like grep.

    .NOTES

    inspired by: https://ridicurious.com/2018/03/14/highlight-words-in-powershell-console/

    .EXAMPLE

    > 'aa bb cc', 'A line' | Select-ColorString a

    Both line 'aa bb cc' and line 'A line' are displayed as both contain "a" case insensitive.

    .EXAMPLE

    > 'aa bb cc', 'A line' | Select-ColorString a -NotMatch

    Nothing will be displayed as both lines have "a".

    .EXAMPLE

    > 'aa bb cc', 'A line' | Select-ColorString a -CaseSensitive

    Only line 'aa bb cc' is displayed with color on all occurrences of "a" case sensitive.

    .EXAMPLE

    > 'aa bb cc', 'A line' | Select-ColorString '(a)|(\sb)' -CaseSensitive -BackgroundColor White

    Only line 'aa bb cc' is displayed with background color White on all occurrences of regex '(a)|(\sb)' case sensitive.

    .EXAMPLE

    > 'aa bb cc', 'A line' | Select-ColorString b -KeepNotMatch

    Both line 'aa bb cc' and 'A line' are displayed with color on all occurrences of "b" case insensitive,
    and for lines without the keyword "b", they will be only displayed but without color.

    .EXAMPLE

    > Get-Content app.log -Wait -Tail 100 | Select-ColorString "error|warning|critical" -MultiColorsForSimplePattern -KeepNotMatch

    Search the 3 key words "error", "warning", and "critical" in the last 100 lines of the active file app.log and display the 3 key words in 3 colors.
    For lines without the keys words, hey will be only displayed but without color.

    .EXAMPLE

    > Get-Content "C:\Windows\Logs\DISM\dism.log" -Tail 100 -Wait | Select-ColorString win

    Find and color the keyword "win" in the last ongoing 100 lines of dism.log.

    .EXAMPLE

    > Get-WinEvent -FilterHashtable @{logname='System'; StartTime = (Get-Date).AddDays(-1)} | Select-Object time*,level*,message | Select-ColorString win

    Find and color the keyword "win" in the System event log from the last 24 hours.
    #>

    [Cmdletbinding(DefaultParametersetName = 'Match')]
    param(
        [Parameter(
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$Pattern = $(throw "$($MyInvocation.MyCommand.Name) : " `
                + "Cannot bind null or empty value to the parameter `"Pattern`""),

        [Parameter(
            ValueFromPipeline = $true,
            HelpMessage = "String or list of string to be checked against the pattern")]
        [String[]]$Content,

        [Parameter()]
        [ValidateSet(
            'Black',
            'DarkBlue',
            'DarkGreen',
            'DarkCyan',
            'DarkRed',
            'DarkMagenta',
            'DarkYellow',
            'Gray',
            'DarkGray',
            'Blue',
            'Green',
            'Cyan',
            'Red',
            'Magenta',
            'Yellow',
            'White')]
        [String]$ForegroundColor = 'Black',

        [Parameter()]
        [ValidateSet(
            'Black',
            'DarkBlue',
            'DarkGreen',
            'DarkCyan',
            'DarkRed',
            'DarkMagenta',
            'DarkYellow',
            'Gray',
            'DarkGray',
            'Blue',
            'Green',
            'Cyan',
            'Red',
            'Magenta',
            'Yellow',
            'White')]
        [ValidateScript( {
                if ($Host.ui.RawUI.BackgroundColor -eq $_) {
                    throw "Current host background color is also set to `"$_`", " `
                        + "please choose another color for a better readability"
                }
                else {
                    return $true
                }
            })]
        [String]$BackgroundColor = 'Yellow',

        [Parameter()]
        [Switch]$CaseSensitive,

        [Parameter(
            HelpMessage = "Available only if the pattern is simple non-regex string " `
                + "separated by '|', use this switch with fast CPU.")]
        [Switch]$MultiColorsForSimplePattern,

        [Parameter(
            ParameterSetName = 'NotMatch',
            HelpMessage = "If true, write only not matching lines; " `
                + "if false, write only matching lines")]
        [Switch]$NotMatch,

        [Parameter(
            ParameterSetName = 'Match',
            HelpMessage = "If true, write all the lines; " `
                + "if false, write only matching lines")]
        [Switch]$KeepNotMatch
    )

    begin {
        $paramSelectString = @{
            Pattern       = $Pattern
            AllMatches    = $true
            CaseSensitive = $CaseSensitive
        }
        $writeNotMatch = $KeepNotMatch -or $NotMatch

        [System.Collections.ArrayList]$colorList =  [System.Enum]::GetValues([System.ConsoleColor])
        $currentBackgroundColor = $Host.ui.RawUI.BackgroundColor
        $colorList.Remove($currentBackgroundColor.ToString())
        $colorList.Remove($ForegroundColor)
        $colorList.Reverse()
        $colorCount = $colorList.Count

        if ($MultiColorsForSimplePattern) {
            # Get all the console foreground and background colors mapping display effet:
            # https://gist.github.com/timabell/cc9ca76964b59b2a54e91bda3665499e
            $patternToColorMapping = [Ordered]@{}
            # Available only if the pattern is a simple non-regex string separated by '|', use this with fast CPU.
            # We dont support regex as -Pattern for this switch as it will need much more CPU.
            # This switch is useful when you need to search some words,
            # for example searching "error|warn|crtical" these 3 words in a log file.
            $expectedMatches = $Pattern.split("|")
            $expectedMatchesCount = $expectedMatches.Count
            if ($expectedMatchesCount -ge $colorCount) {
                Write-Host "The switch -MultiColorsForSimplePattern is True, " `
                    + "but there're more patterns than the available colors number " `
                    + "which is $colorCount, so rotation color list will be used." `
                    -ForegroundColor Yellow
            }
            0..($expectedMatchesCount -1) | % {
                $patternToColorMapping.($expectedMatches[$_]) = $colorList[$_ % $colorCount]
            }

        }
    }

    process {
        foreach ($line in $Content) {
            $matchList = $line | Select-String @paramSelectString

            if (0 -lt $matchList.Count) {
                if (-not $NotMatch) {
                    $index = 0
                    foreach ($myMatch in $matchList.Matches) {
                        $length = $myMatch.Index - $index
                        Write-Host $line.Substring($index, $length) -NoNewline

                        $expectedBackgroupColor = $BackgroundColor
                        if ($MultiColorsForSimplePattern) {
                            $expectedBackgroupColor = $patternToColorMapping[$myMatch.Value]
                        }

                        $paramWriteHost = @{
                            Object          = $line.Substring($myMatch.Index, $myMatch.Length)
                            NoNewline       = $true
                            ForegroundColor = $ForegroundColor
                            BackgroundColor = $expectedBackgroupColor
                        }
                        Write-Host @paramWriteHost

                        $index = $myMatch.Index + $myMatch.Length
                    }
                    Write-Host $line.Substring($index)
                }
            }
            else {
                if ($writeNotMatch) {
                    Write-Host "$line"
                }
            }
        }
    }

    end {
    }
}
Set-Alias slcs Select-ColorString
function hl{
param($regex)
    process{
        $_ | sls $regex | select -ExpandProperty line | slcs $regex
    }
}

# ------------ WEBEX / outlook GUI

function Get-WebexDetails{
param(
    $casenumber="",
    $description="",
    $invitee="",
    $date=(get-date -f "MM/dd"),
    $time=((get-date -Minute 0 -Second 0).addhours(1) |get-date -f "HH:mm"),
    $wxurl="https://verintinc.webex.com/meet/tamas.pasztor",
    $wxbridge="739 168 899",
    $duration=30
)
    Add-Type -AssemblyName System.Windows.Forms

    $f=New-Object system.Windows.Forms.Form
    $f.ClientSize='500,500';$f.Text="WEBEX";$f.BackColor='#efeffe'

    $labels=New-Object System.Windows.Forms.Label[] 8
    $inputs=New-Object System.Windows.Forms.TextBox[] 8
    0..7|%{
	    $labels[$_]=New-Object system.Windows.Forms.Label
	    $labels[$_].AutoSize=$true
	    $labels[$_].Font='Consolas,8'
	    $labels[$_].Width=25
	    $labels[$_].Height=10
	    $labels[$_].Location=New-Object System.Drawing.Point(20,(20+($_*30)))
	
	    $inputs[$_]=New-Object System.Windows.Forms.TextBox
	    $inputs[$_].AutoSize=$true
	    $inputs[$_].Font='Consolas,10'
	    $inputs[$_].Width=350
	    $inputs[$_].Height=10
	    $inputs[$_].Location=New-Object System.Drawing.Point(120,(20+($_*30)))
	    $inputs[$_].Text=""
    }
    $labels[0].Text='Case number:'
    $labels[1].Text='about:'
    $labels[2].Text='email:'
    $labels[3].Text='date:'
    $labels[4].Text='duration (min):'
    $labels[5].Text='webex URL:'
    $labels[6].Text='webex number:'

    $labels[3].Width=60
    $labels[7].Location=New-Object System.Drawing.Point(220,110)
    $labels[7].Width=60
    $labels[7].Text='time:'
    $f.controls.AddRange($labels)

    $inputs[0].Text=$casenumber
    $inputs[1].Text=$description
    $inputs[2].Text=$invitee
    $inputs[3].Text=$date
    $inputs[3].Width=80
    $inputs[4].Text=$duration
    $inputs[5].Text=$wxurl
    $inputs[6].Text=$wxbridge
    $inputs[7].Text=$time
    $inputs[7].Location=New-Object System.Drawing.Point(290,110)
    $inputs[7].Width=80
    #todo: 3 - datetimepicker, today, now, 30 mins etc..
    $f.controls.AddRange($inputs)

    $btnOk=new-object System.Windows.Forms.Button
    $btnOk.Text="Ok"
    $btnOk.Width=50
    $btnOk.Height=20
    $btnOk.Location=New-Object System.Drawing.Point(350,460)
    $btnOk.DialogResult=[System.Windows.Forms.DialogResult]::Ok

    $btnUpdate=new-object System.Windows.Forms.Button
    $btnUpdate.Text="Update"
    $btnUpdate.Width=50
    $btnUpdate.Height=20
    $btnUpdate.Location=New-Object System.Drawing.Point(160,460)
    $btnUpdate.DialogResult=[System.Windows.Forms.DialogResult]::None

    $btnCancel=new-object System.Windows.Forms.Button
    $btnCancel.Text="Cancel"
    $btnCancel.Width=50
    $btnCancel.Height=20
    $btnCancel.Location=New-Object System.Drawing.Point(420,460)
    $btnCancel.DialogResult=[System.Windows.Forms.DialogResult]::Cancel



    $rtfSubject=New-Object System.Windows.Forms.RichTextBox
    $rtfSubject.Location=New-Object System.Drawing.Point(20,230)
    $rtfSubject.Width=450
    $rtfSubject.height=220

    $btnUpdate.Add_Click({
        $caseNumber=$inputs[0].Text
        $issueDescription=$inputs[1].Text
        $hyperlink=$inputs[5].Text
        $bridge=$inputs[6].Text

        $rtfSubject.Rtf="{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Calibri;}{\f2\fswiss\fprq2\fcharset0 Calibri Light;}}
{\colortbl ;\red0\green122\blue255;\red47\green84\blue150;}
\pard\widctlpar\sa160\sl252\slmult1\f0\fs22\lang1038 Hi,\par
I\rquote ve set up this call to discuss \cf1\b VERINT\cf0\b0  case \b $caseNumber\b0 , where \emdash  as per my understanding \emdash  you are experiencing an issue with \i $issueDescription.\i0\par
\pard\keep\keepn\widctlpar\sb240\sl252\slmult1\cf2\f2\fs32\lang1033 Please join my personal Webex room:\par
\pard\widctlpar {\cf0\f0\fs22{\field{\*\fldinst{HYPERLINK $hyperlink }}{\fldrslt{$hyperlink}}}}\par 
\pard\widctlpar\cf0\ulnone $bridge\par}"
    })


    $f.Controls.AddRange(@($btnOk, $btnUpdate, $btnCancel, $rtfSubject))
    
    if (($f.ShowDialog()) -eq [System.Windows.Forms.DialogResult]::OK){
        $startDate=$inputs[3].Text
        $startTime=$inputs[7].Text
        $meetingDuration=$inputs[4].Text
        $start=($startDate|get-date)+($startTime|get-date).TimeOfDay
	    [PSCustomObject]@{
            'casenumber'=$inputs[0].Text;
            'topic'=$inputs[1].Text;
            'invitee'=$inputs[2].Text;
            'start'=$start;
            'end'=$start.AddMinutes($meetingDuration);
            'subject'=$rtfSubject.Rtf}
    }

}
function Create-WebexAppointment{
	begin{
		$ol=New-Object -ComObject outlook.application
#		$store=($ol.GetNamespace("MAPI")).GetStoreFromID("0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000054616D61732E5061737A746F7240766572696E742E636F6D002F6F3D45786368616E67654C6162732F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D30303433333630333366373534303533623837393235336338393632363137352D5061737A746F722C20546100E94632F4440000000200000010000000540061006D00610073002E005000610073007A0074006F007200400076006500720069006E0074002E0063006F006D0000000000")
#		$calendar=$store.GetDefaultFolder(9)
        $calendar=$ol.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
		$myself="Pasztor, Tamas"
		$location="my webex"
		$categories="planned"
        $meetingstatus=[Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
#Organizer                     : Pasztor, Tamas
#MessageClass                  : IPM.Appointment
        $reminderminutes=5
        $reminderset=$true
        $messageclass="IPM.Appointment"
	}

    process{
        if($_ -eq $null){return;}
        $a=$calendar.Application.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)
        $a.MeetingStatus=$meetingstatus
        $a.Categories=$categories
        $a.Location=$location

        $a.start=$_.start
        $a.end=$_.end
        $a.ReminderSet=$reminderset
        $a.ReminderMinutesBeforeStart=$reminderminutes
        $a.BusyStatus=[Microsoft.Office.Interop.Outlook.OlBusyStatus]::olBusy
        $a.Subject="verint case {0} - about {1}" -f $_.casenumber, $_.topic
#        $a.Body=([System.Text.Encoding]::UTF8).GetBytes($_.subject)
        $a.Body=($_.subject)
        #olFormatHTML ??
        $a.BodyFormat=[Microsoft.Office.Interop.Outlook.OlBodyFormat]::olFormatRichText
        $r=$a.Recipients.Add($_.invitee)
        if($r.Resolve()){
            $a.Save()
            $a.Send()
        }
        #ConversationTopic             : teszt mihály
    }
    end{
#        foreach($ref in $r,$a,$calendar,$ol){
        foreach($ref in $calendar,$ol){
            if($ref -ne $null){[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    
    }
}
function Webex{
    param($casenumber,$about,$email=(Get-Clipboard))
    Get-WebexDetails $casenumber $about $email | Create-WebexAppointment
}

function wx-Invitation{
    $alt1="In case you are not available or not willing to join right now, please help me scheduling our troubleshooting session by replying with a time preference."
    $url="https://verintinc.webex.com/meet/tamas.pasztor"
    $bridge="739 168 899"
    $r="
<h3>Please join my webex room</h3>
<table>
    <tr><td>link:</td><td><a href=""$url"">$url</a></td></tr>
    <tr><td>bridge:</td><td><b>$bridge</b></td></tr>
</table>
<p>When there is an option, please ""notify host"" when you arrive.<br/>
The bridge will be kept open for 30 minutes if idle. </p>
<p>$alt1
</p>
"
    Set-Clipboard $r
    Write-Host -ForegroundColor Magenta "Invitation copied to the clipboard"
}


if($false){

# find all occurrences for CTI event regarding extension 75741
$hits=ls int*|sls "\[iemessage.*monitoreddevice> = 75741 ;"|select path, linenumber
# create object with all relevant data
$ex_75741_events=$hits|%{
    $event=gc $_.path|
        select -skip ($_.LineNumber - 6) -first 50 | 
        where{$_ -match "^\[IEMessage"};
    $e=for($i=0;$i -lt $event.count;$i++){
        if($event[$i] -match "Dispatching" -and ($i -lt 5)){$event[$i]}
        else {
            if($event[$i] -notmatch "Dispatching"){$event[$i]} else {break;}
        }
    };
    [pscustomobject]@{
        Timestamp=( $e[0]|%{$_.substring(20,29)|get-date}); 
        Event= $e|%{$_.substring(51)}; 
        Path= $_.path; 
        LineNumber=$_.LineNumber
    }
}


}
function cdc{
    $dir="~\CASES\$(Get-Clipboard)";
    cd $dir
    cd __FromCustomer__
}
Set-Alias c cdc

function Find-InLogs{
    param($pattern)
    ls *.log|sls $pattern
}
Set-Alias f Find-Inlogs

function Get-Manual{
param(
    [Parameter(Mandatory=$true)]$searchTerm,
    [string]$version="15.2"
)
    $basedir="~\Verint Systems Ltd"
    $locations=@{
        "11.1"="WFO Always On Library - Impact 360 V11.1 SP1 Documents\Impact 360 Documents\Documents";
        "15.1"="WFO Always On Library - WFO and CA V15.1 Documentation\WFO_V15.1_Always_On_Documentation\Documents"
        "15.2"="WFO Always On Library - WFO and CA V15.2 Documentation\Documents"
    }
    $results=ls "$basedir\$($locations[$version])" |where name -Match $searchTerm
    if($results.count -eq 1){
        # I'm feeling lucky
        Write-Host -NoNewline -ForegroundColor Green "Congratulations - ";Write-Host "there is only a single hit: $($results[0].Name)"
        start $results[0].FullName
    }else{ 
        if ($results.count -eq 2){
            # there is a pdf and a non-pdf
            Write-Host -NoNewline -ForegroundColor Green "Launching non-pdf version - ";Write-Host "$($results[0].Name)"
            start ($results|where name -NotMatch "\.pdf$"|select -ExpandProperty fullname)
        }else{ 
            if ($results.count -eq 0){ 
                Write-host -ForegroundColor Red "No hits"
            } else{
                Write-Host -NoNewline -ForegroundColor Yellow "There are multiple hits - "; Write-Host "refine your search..."
                Write-Host ("═"*47)
                $results | select -ExpandProperty name
            }
        }
    }
}


#### INIT
Add-Type -Path C:\windows\assembly\gac_msil\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Outlook.dll
#cd ~\CASES
