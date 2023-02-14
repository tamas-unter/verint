$ext="62019"
$log="Session.log"
gc $log|sls "\d{4}-\d{2}-\d{2} (\d{2}:\d{2}:\d{2}\.\d{3}).*Extension.*$ext.*to new Session<(\d+)>.*in new Contact<(\d+)>"|
select @{N="time";E={$_.matches[0].groups[1].Value}}, @{N="session";E={$_.matches[0].groups[2].Value}}, @{N="contact";E={$_.matches[0].groups[3].Value}}
#foreach { $_.matches[0].groups|select -last 3 Value}
#foreach{select @{N="time";E={$_.matches[0].groups[1].Value}}, @{N="session"; E={$_.matches[0].groups[2].Value}}, @{N="contact"; E={$_.matches[0].groups[3].Value}}}

