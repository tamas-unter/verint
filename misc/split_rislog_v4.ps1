function Extract-RISProcesses{
param($processes)
begin{
    New-Item -ErrorAction SilentlyContinue -ItemType Directory "_processes"
#    $files=gci ".\int*.log"
}
process{
    foreach ($process in $processes){
        Write-Progress "Parsing RIS log file" -Status "exporting $process ($($processes.indexof($process)+1)/$($processes.count))" -PercentComplete (($processes.indexof($process)+1)/$processes.count * 100)
		gc $_|sls "^\[$process" | out-file -append -width 2000 -filepath "_processes\$process.log"
    }
}
}