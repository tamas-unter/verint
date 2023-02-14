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