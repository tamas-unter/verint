$regPath='HKLM:\SOFTWARE\WOW6432Node\Witness Systems\eQuality Agent\Capture\CurrentVersion\'
$fakes=@{
	CloudIntegrationServicesServerList="cloudris:29435";
	IntegrationServicesServersList="legacyris:29522";
	IntegrationServicesSecureServerList="secureris:29436";
	IntegrationServicesNonSecureServerList="nonsecureris:29435"
}

get-service *scree*|stop-service

#back up old values
Get-ItemProperty $regPath|
	select CloudIntegrationServicesServerList,IntegrationServicesServersList,IntegrationServicesSecureServerList,IntegrationServicesNonSecureServerList|
	Export-Csv $env:TEMP\oldris.txt
	
# update the values from fakes
$fakes.getEnumerator()|%{
	Set-ItemProperty -Path $regPath -name $_.Key -Value $_.Value
}

Start-Service *scree*