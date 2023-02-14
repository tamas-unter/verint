 <#
 error
 
 "Error 1720.There is a problem with this Windows Installer package. A script required for this install to complete could not be run. Contact your support personnel or package vendor.  Custom action ScreenCaptureProxy_VBRemove script error -2146828235, Microsoft VBScript runtime error: File not found Line 97, Column 2,  "
 
 
 #>


Function Get-ErrorLocation{
    param($errorLine)
    
    $errorLine|
        sls "Custom action (\w+) script error (-?\d+), Microsoft VBScript runtime error: File not found Line (\d+), Column (\d+),"|%{
            "customaction={0},`n line={1},`n column={2}" -f $_.Matches[0].Groups.Value[1,3,4]
        }
}

Function Extract-InstallerScript{
    param(
        $installer_file="C:\Users\tpasztor\OneDrive - Verint Systems Ltd\KB212451-15.2.9.563.msi",
        $table="CustomAction",
        $action="ScreenCaptureProxy_VBRemove",
        $lineNumber=63
    )
    #"select Name from _Tables"

    $winst=New-Object -ComObject WindowsInstaller.Installer


    $msi=$winst.OpenDatabase($installer_file, 0)
    $data=$msi.OpenView("SELECT * FROM $table WHERE Action='$action'")
    $data.Execute()
<#
    $a=while ($d=$data.Fetch()){
        [pscustomobject]@{
            action=$d.StringData(1);
            type=$d.StringData(2);
            source=$d.StringData(3);
            target=$d.StringData(4);
        }
    } 


#    $a| out-gridview

    $vb=$a.target[$a.action.IndexOf($action)].Split("`n")
#>
    $vb=($data.Fetch()).StringData(4).Split("`n")

    for($i=0; $i -lt $vb.Count; $i++){
        if($i -eq ($lineNumber-1)){
            Write-Host -ForegroundColor DarkBlue -BackgroundColor Yellow ("{0}`t{1}" -f @(($i+1),$vb[$i]))
        }else{
            Write-Host ("{0}`t{1}" -f @(($i+1),$vb[$i]))
        }
    }
}



<#

function Get-Property{
    param($o,[object[]]$arguments)
    return $o.GetType().InvokeMember("StringData", 'Public, Instance, GetProperty', $null, $o, $arguments)
}


// action=Get-Property $d 1;
#>