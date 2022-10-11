<# v0.9 10/11/2022
Scripts to break (and fix) recorder functionality - for training purposes


tamas.pasztor@verint.com


#TODO
- execute each function on their role only.
- E:\Impact360\Software\ vs.  \\$server\E$\Impact360\Software\
# GOOD TO HAVE
- state machine for each excercise (broken? fixed?)
- progress bar for service restarts etc.
- fix everything before exit??
- 
#>
########## GENERIC FUNCTIONS

Function Restart-RecorderServices{
    param (
		[Parameter(Mandatory=$true)] $role,
		[Parameter(Mandatory=$true)] [string]$service
	)
	Get-ServerAddressByRole $role|%{
		Invoke-Command -ComputerName $_ -Args $service	{
			Restart-Service $args[0]
		}
	}
}

Function Stop-RecorderService{
    param (
		[Parameter(Mandatory=$true)] $server,
		[Parameter(Mandatory=$true)] [string]$service
	)
	Invoke-Command -ComputerName $server -Args $service {
		get-service |where name -match "watchdog|$($Args[0])"|where starttype -eq "Automatic"|Stop-Service
	}
}

Function Start-RecorderService{
    param (
		[Parameter(Mandatory=$true)] $server,
		[Parameter(Mandatory=$true)] [string]$service
	)
	Invoke-Command -ComputerName $server -Args $service {
		get-service |where name -match "watchdog|$($Args[0])"|where starttype -eq "Automatic"|Start-Service
	}
}

Function Check-Elevation{
    (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}


Function Get-ServerAddressByRole{
param([Parameter(Mandatory=$true)] $role)
    [xml]$s=cat "$env:IMPACT360SOFTWAREDIR\Conf\Cache\Servers.xml"
    $ns=@{svr=$s.Servers.svr}
    $s|select-xml -Namespace $ns "//*/svr:ServerRole[@Name='$role']/../@Hostname"|
        select -ExpandProperty Node|select -ExpandProperty '#text'
}

Function Update-Signature{
    param([Parameter(Mandatory=$true)][string] $configFile)
	#this should be local...
    if(test-path -PathType Leaf "$env:IMPACT360SOFTWAREDIR\ContactStore\ChecksumUtil.exe") {$cmd="$env:IMPACT360SOFTWAREDIR\ContactStore\ChecksumUtil.exe"} else
		{$cmd="$env:IMPACT360SOFTWAREDIR\ContactStore\Tools\ChecksumUtil.exe"}
	
    iex "$cmd -g $configFile"
    
}

### only on recorder!!!
Function Find-PhoneDatasources{
	$role="IP_RECORDER"
    ls "$env:IMPACT360SOFTWAREDIR\Conf\Cache\data*.xml"| 
        sls -list "subtype=""([^""]+)"" type=""phone"""|%{
            [pscustomobject]@{
                t=$_.matches[0].groups[1].value;
                path=$_.path
            }
        }
}

#TODO multiple servers
Function Get-AvayaConfigFile{
    (Find-PhoneDatasources|where t -eq Avaya).path
}


Function Backup-ConfFile{
	param($confFile)
	cp $confFile (Conf-2-Temp $confFile)
}

Function Restore-ConfFile{
	param($confFile)
	mv -Force (Conf-2-Temp $confFile) $confFile
}

Function Conf-2-Temp{param($confFile) $env:TEMP +"\"+ $confFile.Replace("\","_").replace("$","-colon-").replace(":","-colon-")}

############# EXCERCISE 1 #################### 
Function Break-AvayaMRPassCode{
    $role="IP_RECORDER"
	$confFile="\e$\Impact360\Software\conf\IntegrationService.xml"
	$services="Recorder Integration Service"
	
	Get-ServerAddressByRole $role|%{
		Stop-RecorderService -server $_ -service $services
		$cffull="\\$_$confFile"
		Backup-confFile $cffull
		[xml]$is=cat $cffull
		(($is.IFService.Integrations.Integration|where type -eq "AvayaCMAPIAdapter").set|
			where key -eq DevicePassCode).value="0123457777"	
		$is.Save($cffull)
		Update-Signature $cffull
		Start-RecorderService -server $_ -service $services
	}
}

Function Fix-AvayaMRPassCode{
    $role="IP_RECORDER"
	$confFile="\e$\Impact360\Software\conf\IntegrationService.xml"
	$services="Recorder Integration Service"
	Get-ServerAddressByRole $role|%{
		Stop-RecorderService -server $_ -service $services
		Restore-ConfFile "\\$_$confFile"
		Start-RecorderService -server $_ -service $services
	}
}

############# EXCERCISE 2 #################### 
#TODO 	2. NIC configuration -- multiple ones GUID? Wrong guid  <x:Adapters><x:Adapter><x:AdapterId>\Device\NPF_{6D290CBC-C8B7-4FF9-99E7-7C48FFA6C961}</x:AdapterId><x:AdapterName>VMware vmxnet3 virtual network device</x:AdapterName>

Function Break-NICConfiguration{
	$role="IP_RECORDER"
	$confFile="\e$\Impact360\Software\contactstore\IPCaptureConfig.xml"
	$services="Recorder IP CaptureEngine"
	Get-ServerAddressByRole $role|%{
		Stop-RecorderService -server $_ -service $services
		$cffull="\\$_$confFile"
		Backup-confFile $cffull
		[xml]$ipc=cat $cffull
		($ipc.captureconfiguration.Adapters.Adapter|where Mode -eq "Delivery").AdapterId = "\Device\NPF_{6D290CBC-C8B7-4FF9-99E7-0000000000}"
		$ipc.Save($cffull)	
		
		Update-Signature $cffull
		Start-RecorderService -server $_ -service $services
	}
}
Function Fix-NICConfiguration{
    $tempFile="$env:TEMP\ipc.xml"
    mv -Force $tempFile "$(env:impact360softwaredir)\contactstore\IPCaptureConfig.xml"
	Restart-RecorderService $role "Recorder IP CaptureEngine"
}
############# EXCERCISE 3 #################### 
# 3c media definition - break the NAS string
Function Break-ArchiveMedia{
param ($role="ENTERPRISE_ARCHIVER")
    $confFile="$env:IMPACT360SOFTWAREDIR\ContactStore\ArchiverConfig.xml"
    Stop-RecorderService $role "archiver"
    $tempFile="$env:TEMP\arc.xml"
    mv $confFile $tempFile
    [xml]$a=cat $tempFile
    $a.archiver.Devices.Device.DeviceTargets.DeviceTarget.PhysicalDevice="G:\archifoobar"
    $a.Save($confFile)
    Update-Signature $confFile
    Start-RecorderService $role "archiver"
}
Function Fix-ArchiveMedia{
param ($role="ENTERPRISE_ARCHIVER")
    $tempFile="$env:TEMP\arc.xml"
	Stop-RecorderService $role "archiver"
    mv -Force $tempFile "$env:IMPACT360SOFTWAREDIR\ContactStore\ArchiverConfig.xml"
	Start-RecorderService $role "archiver"
}


############# EXCERCISE 4 #################### 
#TODO	4. delete the buffer config for disk manager so that calls are not stored anymore

Function Break-CallsBuffer{
param($role="IP_RECORDER")
    $tempFile="$env:TEMP\rg.xml"
    $confFile="$env:IMPACT360SOFTWAREDIR\ContactStore\RecorderGeneral.xml"
    ###stop ip capture, screen capture
	Stop-RecorderService $role "capture"
    mv $confFile $tempFile
    [xml]$r=cat $tempFile
	
	Start-RecorderService $role "capture"
}
Function Fix-CallsBuffer{
param($role="IP_RECORDER")
    $tempFile="$env:TEMP\rg.xml"
    Stop-RecorderService $role "capture"
	mv -Force $tempFile "$env:IMPACT360SOFTWAREDIR\ContactStore\RecorderGeneral.xml"
	Start-RecorderService $role "capture"
}
############# EXCERCISE 5 #################### 
#5. TSLIB.ini file modification to break connection with aes

Function Break-TSAPI{
param($role="INTEGRATION_FRAMEWORK")
	$tempFile="$env:TEMP\tslib.ini"
	$confFile= 'C:\Program Files (x86)\Avaya\AE Services\TSAPI Client\TSLIB.INI'
	Stop-RecorderService $role "recorder integration"
	$content=[System.IO.File]::ReadAllText($confFile)
	$c=$content.split("`n") #|where {$_ -notmatch "^;"}|where {$_.length -gt 0}
	for ($i=0;$i -lt $c.Count;$i++){
		if($c[$i] -match "^[^;].*=450") {$c[$i]="0.1.2.3=450"}
	}
	[System.IO.File]::WriteAllText(
		$confFile, (
			($c |where {[string]::IsNullOrEmpty($_) -eq $false}) -join "`r`n"
		)
	)
	Start-RecorderService $role "recorder integration"
	
}
Function Fix-TSAPI{
param($role="INTEGRATION_FRAMEWORK")
	$tempFile="$env:TEMP\tslib.ini"
	## TODO!! network location - \\hostname\c$\ ...
	$confFile= 'C:\Program Files (x86)\Avaya\AE Services\TSAPI Client\TSLIB.INI'
	Stop-RecorderService $role "recorder integration"
	mv -Force $tempFile $confFile
	Start-RecorderService $role "recorder integration"
}
############# EXCERCISE 6 #################### 
#TODO 6. remove extensions from data source. 

Function Break-Extensions{
}
Function Fix-Extensions{
}
############# EXCERCISE 7 #################### 
#TODO SCM registry update -- already done, check INNO LAB
### this one needs to run on a remote PC via Remote Registry service. if that is not enabled, we'll need to Remote PS
function Break-SCMRis{
param($hostName="ADVDSK9")
    Invoke-Command -ComputerName $hostName {
        get-service *screencapture*|stop-service
        $oldRIS=Get-ItemProperty -path 'HKLM:\SOFTWARE\WOW6432Node\Witness Systems\eQuality Agent\Capture\CurrentVersion\' -name IntegrationServicesServersList

        
        $newRIS="dummy.wfo.verint.training:29522"
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Witness Systems\eQuality Agent\Capture\CurrentVersion\' -name IntegrationServicesServersList -Value $newRIS
        Start-Service *scree
    }
}
function Fix-SCMRis{

}



############# EXCERCISE 8 #################### 
#TODO 8. remove hunt group from DS to produce tagging issue
function Break-HuntGroup{}
function Fix-HuntGroup{}
############# EXCERCISE 9 #################### 
#TODO 	9. Hosts file, app server false entry! -- this is rather infra.. Ris to rec rec to archiver… app server???
function Break-AppServerAddress{}
function Fix-AppServerAddress{}
############# EXCERCISE 10 #################### 
#TODO 	10. Remove adapter

############# EXCERCISE 11 #################### 
#TODO	11. Disable adapter


#TODO 	12. archive user account - ?? change share rights to prevent IMSA user
Function Break-ArchiveUser{
}
Function Fix-ArchiveUser{
}

### MAIN ENTRY ###
if ((Check-Elevation) -eq $false) {
    Write-Host -ForegroundColor Red "Please restart as Administrator."
#    Exit
}

# TODO
# populate integers, all tasks bound to buttons

# $Global:STATES=@(@($true,$false),@($true,$false),@($true,$false),@($true,$false))

############### GUI
$excercise_count=11

Add-Type -AssemblyName System.Windows.Forms

$f=New-Object system.Windows.Forms.Form
$f.ClientSize='250,{0}' -f (40+$excercise_count*30);
$f.Text="Administrator - Breaking the lab";
$f.BackColor='#efeffe'


$labels=New-Object System.Windows.Forms.Label[] $excercise_count
$breakers=New-Object System.Windows.Forms.Button[] $excercise_count
$fixers=New-Object System.Windows.Forms.Button[] $excercise_count

0..($excercise_count-1)|%{
    $labels[$_]=New-Object System.Windows.Forms.Label
    $labels[$_].AutoSize=$true
    $labels[$_].Text=($_+1)
    $labels[$_].Location=New-Object System.Drawing.Point(15,(20+($_*30)))

    $breakers[$_]=New-Object System.Windows.Forms.Button
    $breakers[$_].Text="Break"
    $breakers[$_].Location=New-Object System.Drawing.Point(40,(15+($_*30)))

    $fixers[$_]=New-Object System.Windows.Forms.Button
    $fixers[$_].Text="Fix"
    $fixers[$_].Location=New-Object System.Drawing.Point(120,(15+($_*30)))
    
}

$breakers[0].Add_Click({Break-AvayaMRPassCode});
$fixers[0].Add_Click({Fix-AvayaMRPassCode});
$breakers[1].Add_Click({Break-NICConfiguration});
$fixers[1].Add_Click({Fix-NICConfiguration});
$breakers[2].Add_Click({Break-ArchiveMedia});
$fixers[2].Add_Click({Fix-ArchiveMedia});
$breakers[3].Add_Click({Break-CallsBuffer});
$fixers[3].Add_Click({Fix-CallsBuffer});

$f.Controls.AddRange($labels)
$f.Controls.AddRange($breakers)
$f.Controls.AddRange($fixers)



$f.ShowDialog()