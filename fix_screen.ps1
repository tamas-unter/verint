$regPath='HKLM:\SOFTWARE\WOW6432Node\Witness Systems\eQuality Agent\Capture\CurrentVersion\'


get-service *scree*|stop-service

# restore from backup
(Import-Csv $env:TEMP\oldris.txt).psObject.Properties|%{
	Set-ItemProperty -Path $regPath -name $_.Name -Value $_.Value
}

Start-Service *scree*