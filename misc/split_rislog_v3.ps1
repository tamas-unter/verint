function Get-RISProcess{
param($process)
begin{
#    New-Item -ErrorAction SilentlyContinue -ItemType Directory "_processes"
#    $files=gci ".\int*.log"
}
process{
		gc $_|sls "^\[$process" #| out-file -append -width 2000 -filepath "_processes\$process.log"
}
}