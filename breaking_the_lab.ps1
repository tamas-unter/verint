<# v1.37 1/10/2023
Scripts to break (and fix) recorder functionality - for training purposes


tamas.pasztor@verint.com


# GOOD TO HAVE
- state machine for each excercise (broken? fixed?)
- progress bar for service restarts etc.
- fix everything before exit??
- feedback after the operation was successful
- error handling
#>

### DESKTOP hostnames
$Global:desktops=@("B155-D01","B155-D02")

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
    param(
		[Parameter(Mandatory=$true)][string] $server,
		[Parameter(Mandatory=$true)][string] $configFile
	)
    if(test-path -PathType Leaf "\\$server\e$\IMPACT360\SOFTWARE\ContactStore\ChecksumUtil.exe") {$cmd="\\$server\e$\IMPACT360\SOFTWARE\ContactStore\ChecksumUtil.exe"} else
		{$cmd="\\$server\e$\IMPACT360\SOFTWARE\ContactStore\Tools\ChecksumUtil.exe"}
	
    iex "$cmd -g ""$configFile"""
    
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

Function Do-BreakExcercise{
	param($sender,$e)
	$index=$sender.TabIndex
	Write-Host ("Breaking excercise {0}" -f $index)
	Break-Config -role $Global:excercises[$index].role -confFile $Global:excercises[$index].confFile -services $Global:excercises[$index].services -breakFunction $Global:excercises[$index].breakFunction
}
Function Do-FixExcercise{
	param($sender,$e)
	$index=$sender.TabIndex
	Write-Host ("Fixing excercise {0}" -f $index)
	Fix-Config -role $Global:excercises[$index].role -confFile $Global:excercises[$index].confFile -services $Global:excercises[$index].services  -fixFunction $Global:excercises[$index].fixFunction
}


Function Fix-Config{
	param(
		$role,
		$confFile,
		$services,
		$fixFunction
	)
	if ($fixFunction -eq $null) {
		Get-ServerAddressByRole $role|%{
			Stop-RecorderService -server $_ -service $services
			Restore-ConfFile "\\$_$confFile"
			Start-RecorderService -server $_ -service $services
		}
	}	else {
		Invoke-Command -ArgumentList $confFile -ScriptBlock $fixFunction
	}
}

Function Break-Config{
	param(
		$role,
		$confFile,
		$services,
		$breakFunction
	)	
	if($role -ne $null) {
		Get-ServerAddressByRole $role|%{
			Stop-RecorderService -server $_ -service $services
			$cffull="\\$_$confFile"
			Backup-confFile $cffull
			
			Invoke-Command -ArgumentList $cffull -ScriptBlock $breakFunction
			if($confFile -match "\.xml$") {Update-Signature $_ $cffull}	
			Start-RecorderService -server $_ -service $services
		}
	} else {
		Invoke-Command -ArgumentList $confFile -ScriptBlock $breakFunction
		
	}
}

# in this array, each element is containing the definition of the excercise. In case the $role is missing ($null), $confFile will be the argument for the breakFunction
## BEGINNING of excercise definition
$Global:excercises=@(
	#1: AvayaMRPassCode
	@{
		role="IP_RECORDER"
		confFile="\e$\Impact360\Software\conf\IntegrationService.xml"
		services="Recorder Integration Service"
		breakFunction={
			param($cffull)
			[xml]$is=cat $cffull
			(($is.IFService.Integrations.Integration|where type -eq "AvayaCMAPIAdapter").set|
				where key -eq DevicePassCode).value="0123457777"	
			$is.Save($cffull)
		}
	},
	#2: NICConfiguration
	@{
		role="IP_RECORDER"
		confFile="\e$\Impact360\Software\contactstore\IPCaptureConfig.xml"
		services="Recorder IP CaptureEngine"
		breakFunction={
			param($cffull)
			[xml]$ipc=cat $cffull
			($ipc.captureconfiguration.Adapters.Adapter|where Mode -eq "Delivery").AdapterId = "\Device\NPF_{6D290CBC-C8B7-4FF9-99E7-0000000000}"
			$ipc.Save($cffull)
		}
	},
	#3: ArchiveMedia
	@{
		role="ENTERPRISE_ARCHIVER"
		confFile="\e$\Impact360\Software\ContactStore\ArchiverConfig.xml"
		services="archiver"
		breakFunction={
			param($cffull)
			[xml]$a=cat $cffull
			$a.archiver.Devices.Device.DeviceTargets.DeviceTarget.PhysicalDevice="G:\archifoobar"
			$a.Save($cffull)
		}
	},
	#4: CallsBuffer
	@{
		role="IP_RECORDER"
		confFile="\e$\Impact360\Software\ContactStore\RecorderGeneral.xml"
		services= "capture"
		breakFunction={
			param($cffull)
			[xml]$r=cat $cffull
			$r.recordergeneral.callspath=""
			$r.Save($cffull)
		}
	},
	#5: TSAPI
	@{
		role="INTEGRATION_FRAMEWORK"
		confFile='\C$\Program Files (x86)\Avaya\AE Services\TSAPI Client\TSLIB.INI'
		services= "recorder integration"
		breakFunction={
			param($cffull)
			$content=[System.IO.File]::ReadAllText($cffull)
			$c=$content.split("`n") #|where {$_ -notmatch "^;"}|where {$_.length -gt 0}
			for ($i=0;$i -lt $c.Count;$i++){
				if($c[$i] -match "^[^;].*=450") {$c[$i]="0.1.2.3=450"}
			}
			[System.IO.File]::WriteAllText(
				$cffull, (
					($c |where {[string]::IsNullOrEmpty($_) -eq $false}) -join "`r`n"
				)
			)			
		}
	},
	#11: disable tsapi adapter
	@{
		role="INTEGRATION_FRAMEWORK"
		confFile='\e$\Impact360\Software\Conf\IntegrationService.xml'
		services="recorder integration"
		breakFunction={
			param($cffull)
			[xml]$r=cat $cffull
			($r.IFService.Integrations.Integration|where Type -eq "TSAdapter").StartupType="Disabled"
			$r.Save($cffull)
		}
	},
	#13: ARCHIVER media location
	@{
		confFile='HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares'
		breakFunction={
			param($regPath)
            Get-ServerAddressByRole "ENTERPRISE_ARCHIVER"|%{
			    Invoke-Command -ComputerName $_ -ArgumentList $regPath -ScriptBlock {
				    param($regPath)
				    stop-service LanmanServer
				    $oldRIS=Get-ItemProperty -path $regPath -name archive_media
				    $oldRIS|select -ExpandProperty archive_media|out-file $env:TEMP\oldshare.txt
				
				    $newRIS=$oldRIS.Replace("K:","X:")
				    Set-ItemProperty -Path $regPath -name archive_media -Value $newRIS
				    Start-Service LanmanServer
			    }
            }			
		}
		fixFunction={
			param($regPath)
            Get-ServerAddressByRole "ENTERPRISE_ARCHIVER"|%{
			    Invoke-Command -ComputerName $_ -ArgumentList $regPath -ScriptBlock {
				    param($regPath)
				    stop-service LanmanServer
				    $newRIS=cat $env:TEMP\oldshare.txt
				    Set-ItemProperty -Path $regPath -name archive_media -Value $newRIS
				    Start-Service LanmanServer
			    }
            }
		}
	},
	#14: tlink definition 
		role='INTEGRATION_FRAMEWORK'
		services="recorder integration"
		conffile='\e$\Impact360\Software\Conf\IntegrationService.xml'
		breakFunction={
			param($cffull)
			[xml]$r=cat $cffull
			($r|Select-Xml -namespace @{x=$r.IFService.xmlns} "//*/x:Integration[@Type='TSAdapter']/x:Set[@Key='ServiceId']").Node.Value="AVAX#BRRR#CSTA#AAA"
			
			
			$r.Save($cffull)
		}
	}
	#17: java path broken
	@{
		conffile='E:\Impact360\Software\OpenJDK\bin\java.exe'
		breakFunction={
			param($conffile)
			Get-ServerAddressByRole "INTEGRATION_FRAMEWORK"|%{
				Invoke-Command -ComputerName $_ -ArgumentList $conffile -ScriptBlock {
					param($java)
					stop-service "Recorder Integration Service"
					mv $java ($java.Replace("java.exe","j.exe"))
					start-service "Recorder Integration Service"
				}
				
			}
		}
		fixFunction={
			param($conffile)
			Get-ServerAddressByRole "INTEGRATION_FRAMEWORK"|%{
				Invoke-Command -ComputerName $_ -ArgumentList $conffile -ScriptBlock {
					param($java)
					stop-service "Recorder Integration Service"
					mv $java ($java.Replace("j.exe","java.exe"))
					start-service "Recorder Integration Service"
				}
				
			}
			
		}
	}


# network shares under this key
## HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares
<#
CATimeout=0
CSCFlags=0
MaxUses=4294967295
Path=K:\archive_media
Permissions=0
Remark=FOO bar
ShareName=archive_media
Type=0


#>

<#
,
	#7: SCMRis
	@{
		confFile='HKLM:\SOFTWARE\WOW6432Node\Witness Systems\eQuality Agent\Capture\CurrentVersion\'
		breakFunction={
			param($regPath)
            $Global:desktops|%{
			    Invoke-Command -ComputerName $_ -ArgumentList $regPath -ScriptBlock {
				    param($regPath)
				    get-service *screencapture*|stop-service
				    $oldRIS=Get-ItemProperty -path $regPath -name IntegrationServicesServersList
				    $oldRIS|select -ExpandProperty IntegrationServicesServersList|out-file $env:TEMP\oldris.txt
				
				    $newRIS="dummy.wfo.verint.training:29522"
				    Set-ItemProperty -Path $regPath -name IntegrationServicesServersList -Value $newRIS
				    Start-Service *scree*
			    }
            }			
		}
		fixFunction={
			param($regPath)
            $Global:desktops|%{
			    Invoke-Command -ComputerName $_ -ArgumentList $regPath -ScriptBlock {
				    param($regPath)
				    get-service *screencapture*|stop-service
				    $newRIS=cat $env:TEMP\oldris.txt
				    Set-ItemProperty -Path $regPath -name IntegrationServicesServersList -Value $newRIS
				    Start-Service *scree
			    }
            }
		}
	}
#>
#### TODO database change EM
)

## END of excercise definition

############# EXCERCISE 8 #################### 
#TODO 8. remove hunt group from DS to produce tagging issue
function Break-HuntGroup{}
function Fix-HuntGroup{}
############# EXCERCISE 9 #################### 
#TODO 	9. Hosts file, app server false entry! -- this is rather infra.. Ris to rec rec to archiverâ€¦ app server???
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
$excercise_count=$Global:excercises.Count

Add-Type -AssemblyName System.Windows.Forms

$f=New-Object system.Windows.Forms.Form
$f.ClientSize='250,{0}' -f (40+$excercise_count*30);
$f.Text="Administrator - Breaking the lab";
$f.BackColor='#efeffe'


$labels=New-Object System.Windows.Forms.Label[] $excercise_count
$breakers=New-Object System.Windows.Forms.Button[] $excercise_count
$fixers=New-Object System.Windows.Forms.Button[] $excercise_count

# need to stick to the documentation. excercise numbers are not contiguous
$ex_numbers=1,2,3,4,5,11,13,14,16,24,25,26

0..($excercise_count-1)|%{
    $labels[$_]=New-Object System.Windows.Forms.Label
    $labels[$_].AutoSize=$true
    $labels[$_].Text=($ex_numbers[$_])
    $labels[$_].Location=New-Object System.Drawing.Point(15,(20+($_*30)))

    $breakers[$_]=New-Object System.Windows.Forms.Button
    $breakers[$_].Text="Break"
    $breakers[$_].TabIndex=$_
    $breakers[$_].Location=New-Object System.Drawing.Point(40,(15+($_*30)))
	$breakers[$_].Add_Click({Do-BreakExcercise $this $_})

    $fixers[$_]=New-Object System.Windows.Forms.Button
    $fixers[$_].Text="Fix"
	$fixers[$_].TabIndex=$_
    $fixers[$_].Location=New-Object System.Drawing.Point(120,(15+($_*30)))
	$fixers[$_].Add_Click({Do-FixExcercise $this $_})
    
}

$f.Controls.AddRange($labels)
$f.Controls.AddRange($breakers)
$f.Controls.AddRange($fixers)



$f.ShowDialog()