[pscredential]$cred=New-Object pscredential ("Administrator", (ConvertTo-SecureString "TPVerint0123456789" -AsPlainText -Force))
$prepend="E:\Impact360\Software\ContactStore;"


<#
0) press F5 or debug (to add the script in ISE Memory by hitting green arrown (run)) -- run


1) create a list of all wfo servers and paths into a csv file: will be created in %userprofile% (This envoke server.xml file to pull the list)
        > create-list LIST.CSV

        NOTE: First rederror text is expected from the server the script is going to run from as you can't envoke remote command on local host

2) edit file notepad or excel, remove empty valued lines -having "","", review the list , DO NOT LEAVE empty lines there..
This file will be used to update for next step
        > notepad LIST.CSV


3) push the changes
        > update-all LIST.CSV

#>



# this is retrieveing the list of all servers on WFO
function get-allwfohosts{
    [xml]$x=gc ($env:impact360softwaredir+"Conf\Cache\servers.xml")
    $ns=@{svr=$x.ChildNodes[1].Attributes[0].'#text'}
    $x|Select-Xml -Namespace $ns "//svr:Server/@Hostname"|select -ExpandProperty node|select -ExpandProperty '#text'
}

# this is getting the current path value from a machine $h
function get-remotepath{
param($h)
    Invoke-Command -ComputerName $h {
        [System.Environment]::GetEnvironmentVariable("path",[System.EnvironmentVariableTarget]::Machine)
       }
}


# this creates the preferred version of the path (not actually updating anything..)
function update-path{
param($remotepath)
    if ($remotepath.IndexOf("AVS") -eq -1) {$remotepath} else{

        $prefix="E:\Impact360\Software\ContactStore;"
        $p=$remotepath.replace($prefix,"")

	    $prefix,$p -join ""

    }
}



# opposite of get-remotepath: overwrites the path env variable on a machine $h
function set-remotepath{
param($h,$p)
    Invoke-Command -ComputerName $h  -ArgumentList $p { 
        [System.Environment]::SetEnvironmentVariable( 'Path', $args, [System.EnvironmentVariableTarget]::Machine )
    }
}


# this creates a CSV file having 3 columns:  host, old path, new path
function create-list{
param($filename)
    $wfohosts=get-allwfohosts
    $count=$wfohosts.count
    $i=0
    $wfohosts|%{
        $i++
        Write-Progress -Activity "reading values from server" -Status $_ -PercentComplete (100*$i/$count)
        $orig=(get-remotepath $_)
        [pscustomobject]@{
            host=$_
            original=$orig
            modified=(update-path $orig)
        }
    }|Export-Csv $filename

}


# this reads the CSV file and updates the values for each server
function update-all{
param($filename)
    $list=Import-Csv $filename
    $count=$list.count
    $i=0
    $list|%{
        $i++
        Write-Progress -Activity "pushing changes to server" -Status $_.host -PercentComplete (100*$i/$count)
        set-remotepath $_.host $_.modified
    }
        
}




function get-remoteEnv{
[cmdletbinding()] 
param( 
    [string[]]$ComputerName =$env:ComputerName, 
    [string]$Name 
) 
 
foreach($Computer in $ComputerName) { 
    Write-Verbose "Working on $Computer" 
    if(!(Test-Connection -ComputerName $Computer -Count 1 -quiet)) { 
        Write-Verbose "$Computer is not online" 
        Continue 
    } 
     
    try { 
        $EnvObj = @(Get-WMIObject -Class Win32_Environment -ComputerName $Computer -EA Stop) 
        if(!$EnvObj) { 
            Write-Verbose "$Computer returned empty list of environment variables" 
            Continue 
        } 
        Write-Verbose "Successfully queried $Computer" 
         
        if($Name) { 
            Write-Verbose "Looking for environment variable with the name $name" 
            $Env = $EnvObj | Where-Object {$_.Name -eq $Name} 
            if(!$Env) { 
                Write-Verbose "$Computer has no environment variable with name $Name" 
                Continue 
            } 
            $Env             
        } else { 
            Write-Verbose "No environment variable specified. Listing all" 
            $EnvObj 
        } 
         
    } catch { 
        Write-Verbose "Error occurred while querying $Computer. $_" 
        Continue 
    } 
 
}
}

<#
1) server names - all servers
[xml]$x=gc ($env:impact360softwaredir+"Conf\Cache\servers.xml")
$ns=@{svr=$x.ChildNodes[1].Attributes[0].'#text'}
$wfohosts=$x|Select-Xml -Namespace $ns "//svr:Server/@Hostname"|select -ExpandProperty node|select -ExpandProperty '#text




2) retrieve path env var remotely

 $e=Get-WmiObject -ComputerName $h -Credential $cred -Class win32_environment

$p=$e|where {$_.name -eq "path"}|select -ExpandProperty variablevalue


3) update

 if ($p.IndexOf("AVS") -gt -1) { }

$prefix="E:\Impact360\SOFTWARE\ContactStore;"
$newpath=$prefix,$p -join ""


4) export csv

5) import csv

6) push changes

$result = Invoke-Command -ComputerName $remoteMachines -Credential $credentials -ArgumentList $Environment {
  [System.Environment]::SetEnvironmentVariable( 'ASPNETCORE_ENVIRONMENT', $args, [System.EnvironmentVariableTarget]::Machine )
}






PS C:\Users\tamas> $e|where {$_.name -eq "path"}|select -ExpandProperty variablevalue
E:\Impact360\Software\ContactStore;C:\ProgramData\Oracle\Java\javapath;E:\Impact360\Software\\Utils;%SystemRoot%\system
32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SYSTEMROOT%\System32\WindowsPowerShell\v1.0\;C:\Program Files\Microsoft SQL
 Server\120\DTS\Binn\;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\;C:\Program Files (x86)\Micr
osoft SQL Server\120\Tools\Binn\;C:\Program Files\Microsoft SQL Server\120\Tools\Binn\;C:\Program Files (x86)\Microsoft
 SQL Server\120\Tools\Binn\ManagementStudio\;C:\Program Files (x86)\Microsoft SQL Server\120\DTS\Binn\;E:\Impact360\Sof
tware\jre\bin;E:\Impact360\Software\JRE64\bin;E:\Impact360\Software\VerintSDK\Bin\;E:\Impact360\Software\TeleflowServer
\;E:\Impact360\Software\TeleflowServer\TFx86\;E:\Impact360\Software\intelliminer
PS C:\Users\tamas> $p=$e|where {$_.name -eq "path"}|select -ExpandProperty variablevalue
PS C:\Users\tamas> $p.indexof("AVS")
-1
PS C:\Users\tamas> $p.indexof("ContactStore\;")
-1
PS C:\Users\tamas> $p.indexof("Binn\;")
254




#>