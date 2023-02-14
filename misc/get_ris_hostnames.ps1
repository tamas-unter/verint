function Get-RISHosts{
	[xml]$x=gc .\Conf\Cache\Servers.xml
	$ns=@{srv=$x.ChildNodes[1].Attributes[0].'#text'}
	($x|select-xml -Namespace $ns "//srv:ServerRole[@Name='INTEGRATION_FRAMEWORK']/..").Node|select HostName, Identifier, Name, Version
}