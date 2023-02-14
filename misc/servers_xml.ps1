# get all recorders
#$servers_xml|Select-Xml -Namespace $ns "//svr:ServerRole[contains(@Name,'_RECORDER')]/.."|select -ExpandProperty Node|select name,hostname
#query
function Load-ServersXml{
    [xml]$global:servers_xml=ls -Recurse Servers.xml|select -first 1|%{cat $_}
    $global:ns=@{svr=$global:servers_xml.Servers.svr}

}

function Query-Servers{
param([string]$q)
    $global:servers_xml|Select-Xml -Namespace $global:ns $q|
        select -ExpandProperty Node
}