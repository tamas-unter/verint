$c="9137995682660001781"
$log=".\_processes\Session.log"
$regex= "\d{4}-\d{2}-\d{2} (\d{2}:\d{2}:\d{2}\.\d{3})\S+\s+(\w+).+Connection <(\d+)> <Extension(<[^>]+>).* Session<(\d+)> State<(\w+)> .*Contact<(\d+)>"
gc $log|sls $regex|
select @{N="time";E={$_.matches[0].groups[1].Value}}, 
	@{N="action";E={$_.matches[0].groups[2].Value}}, 
	@{N="connection";E={$_.matches[0].groups[3].Value}},
	@{N="extension";E={$_.matches[0].groups[4].Value}},
	@{N="session";E={$_.matches[0].groups[5].Value}},
	@{N="state";E={$_.matches[0].groups[6].Value}}, 
	@{N="contact";E={$_.matches[0].groups[7].Value}} |
where contact -eq $c |
ft
