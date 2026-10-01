### personal 
<#
#plain name First Last
$myself
#outlook name Last, First
$email
#api user
$cred=@{user;password}
#>
ipmo "~\OneDrive - Verint Systems Ltd\ps\secrets.ps1"
function activity_emoji{
param($a)

    switch ($a){
        "VAS_TKS_Recorder" {"🏠"}
        "VAS_TKS_Recorder_Hybrid" {"☁"}
        default {"⭐­"}
    }
}
function Create-AppointmentsForToday{
	begin{
		$ol=New-Object -ComObject outlook.application
        $calendar=$ol.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
	}
	process{
		$start=$_.from
		$end=$_.to
		
		if($_.with.length -eq 0){
			$subject="TW ALONE";
		} else{
			$subject="TW #$($_.sequence) with $($_.with)"
		}

		$body=@"
$($_.schedule)

Previous Hour: `t$($_.before)
Next Hour:   `t$($_.after)
"@
		$a=$calendar.Application.CreateItem(1)
		$a.MeetingStatus=0
		$a.Start=$start
		$a.End=$end
		$a.RequiredAttendees=$email
		$a.ReminderSet=$true
        $a.ReminderMinutesBeforeStart=0
		$a.Subject=$subject
		$a.Location=$_.activity
		$a.Body=$body
		$a.Categories="TicketWatch"
		$a.Save()
		Write-Host ("{0:HH:mm}-{1:HH:mm}: {2}" -f $start,$end,$subject)
	}
    end{
        Write-Host "."
    }
}

$wfo="wfo.f2.verintcloudservices.com"
$uri="https://$wfo/wfo/rest/core-api/auth/token"

# get token for further request authentications
$r=Invoke-RestMethod -Method Post -Uri $uri -Body ($cred|ConvertTo-Json) -Headers @{host=$wfo} -ContentType 'application/json'
$token=$r.AuthToken.token

# initialize
$employees=@()
$schedules=@()
# only today's schedule is needed.
$startTime=get-date -Format "yyyy-MM-ddT00:00:00Z"
$endTime=(get-date).AddDays(1)| get-date -Format "yyyy-MM-ddT00:00:00Z"

# get all employees from the organizations

# these are the org ids, including the parent "VAS_RECORDER"
3566..3569|%{
    #get staff
    $uri="https://$wfo/wfo/user-mgmt-api/v1/organizations/$_/employees"
    $r=irm -Method Get -Uri $uri -Headers @{host=$wfo;Impact360AuthToken=$token}
    $employees+=$r.data|%{[pscustomobject] @{id=$_.id;attributes=$_.attributes}}
    #get schedules
    $uri="https://$wfo/wfo/rest/fs-api/schedule/get-with-shifts?startTime=$startTime&endTime=$endTime&organizationId=$_"
    $r=irm -Method Get -Uri $uri -Headers @{host=$wfo;Impact360AuthToken=$token} 
    $schedules+=$r.data.schedules|%{[pscustomobject]@{
        employeeId=$_.employeeId;
        activityId=$_.activityId;
        eventType=$_.eventType;
        activityName=$_.activityName;
        startTime=$_.startTime | get-date;
        endTime=$_.endTime | get-date
    }}
}

# every other week the order is A->Z otherwise Z->A.
$az= -not (([cultureinfo]::CurrentCulture.Calendar.GetWeekOfYear((get-date),'FirstDay', 'Monday') %2) -eq 0)

$d=get-date -Hour 0 -Minute 0
$p=@{Descending=-not $az;Property={$_.e.attributes.lastName}}
#$employees=$employees|sort @p

$s=0..23|%{
    [pscustomobject]@{
        h=$_;
        e=$schedules|
            where starttime -le ($d.AddHours($_))| 
            where endtime -ge ($d.AddHours($_))|
            where eventtype -notin (64,128,512)|%{
                [pscustomobject]@{
                    e=$employees|where id -EQ $_.employeeId
                    a=$_.activityName
                }
            }| sort @p
    }
}

$r=$s|%{
    [pscustomobject]@{
        h=$_.h;
        s=$_.e|%{
            [pscustomobject]@{
                a=$_.a;
                n=("{0} {1}" -f $_.e.attributes.firstName.TRIM(),$_.e.attributes.lastName.trim())
            }
        }
    }
}

$mywatches=@()
0..23|%{
    if ($r[$_].s.n -match $myself){
        if($_ -le 0){$_before=""} else {$_before=($r[$_-1].s|%{"{0}{1}" -f (activity_emoji $_.a),$_.n}) -join " + "}
        if($_ -ge 23){$_after=""} else {$_after=($r[$_+1].s|%{"{0}{1}" -f (activity_emoji $_.a),$_.n}) -join " + "}
        $i=1
        $mywatches+=([pscustomobject]@{
            from=get-date -hour $_ -minute 0 -Second 0
            to=get-date -hour ($_+1) -minute 0 -Second 0
            ### use emojis
            with=($r[$_].s.n|where {$_ -notmatch $myself}) -join "+"
            before=$_before
            schedule=($r[$_].s|%{"{2}.{0}{1}" -f (activity_emoji $_.a),$_.n,$i++}) -join "`n"
            after=$_after
            sequence=$r[$_].s.n.IndexOf($myself)+1
            activity=$r[$_].s|where {$_.n -match $myself}|select -ExpandProperty a
        })
    }
}

# dump today's schedule
$r|%{
    [pscustomobject]@{
        hour=$_.h;
        schd=($_.s|%{"{0}{1}" -f ((activity_emoji $_.a),$_.n)}) -join ", "
    }
}|Out-GridView -Title "TODAY'S ROSTER" 





#### create the outlook items
#$mywatches |Create-AppointmentsForToday


# SIG # Begin signature block
# MIIFbQYJKoZIhvcNAQcCoIIFXjCCBVoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUlHVUWo86lcGec62525d1j/Tz
# ckagggMIMIIDBDCCAeygAwIBAgIQGBcyty7cFrZDfL0x3Dyg2zANBgkqhkiG9w0B
# AQsFADAaMRgwFgYDVQQDDA9Qw6FzenRvciBUYW3DoXMwHhcNMjYwNDA4MTY0MTM4
# WhcNMjcwNDA4MTcwMTM4WjAaMRgwFgYDVQQDDA9Qw6FzenRvciBUYW3DoXMwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDTgh8TsG2Szup66jFsBgwMLAaK
# 1wJSJiiWM9sFlIdOZ15+SoKoP+w3gR6S6Sq4d0/+jXK9Z1Df3sBBzt2Ti0Btuq7s
# chikqMRcyb3fS6qCgM4VwaMEgOhT4Eq/sLjQFGRbk0JkOBuyAJaUJvLX/nqhirJt
# i5Zxsa+Vzbsxgwsz/PGIyi5zhVl9o7z1ijxgPhjB8DNQxgO+KRO7jMgXs0cUguBU
# /J1pASzp9DQh3L/riSOdo0XykL9/Chj8sD2l56XxiO9ycn0UFn/GDASm2j3bpZB+
# dmijks/TRSk8MPRpM7Am+wTltrQkhyseh5di2B5lSU38nM7OvznYp7Pg1jv5AgMB
# AAGjRjBEMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQUIi9tK/EtLCoxxH1IaNJjv0g5cQMwDQYJKoZIhvcNAQELBQADggEBAJnY
# ZpFvuYYeAC0P+iFvTuFuqAUcMOfhh9416B3uEApUmDlVGRCBiYahuDGs0/tVgDw4
# FsLi451sMwPJxofTXZhRbIAyKz7jgUTvsbpB6sysg1cN7Igm6FMfsE7/Zt652DSz
# BNFEoXslQMzVknCfZYpsAha9PTYkBzjpEaIiUC3ovtW/toG4VIuyoQyEtmUaeg/W
# NoPgrARY0X0U31b998yKfSZkweta0VAe5DjsJkxFkS+FnPnUnfvVHIfXqCndVLZ1
# iXHyZkwEedq/3rK4TDdIN9AKtAKGa/iUk1V8NQDSemz9r902xacobjBh1GzUrEAz
# e3i1+hf2WaoOrC/3nOMxggHPMIIBywIBATAuMBoxGDAWBgNVBAMMD1DDoXN6dG9y
# IFRhbcOhcwIQGBcyty7cFrZDfL0x3Dyg2zAJBgUrDgMCGgUAoHgwGAYKKwYBBAGC
# NwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUzdj97WC2
# Tqk+8lilZX5W4+ZFDnkwDQYJKoZIhvcNAQEBBQAEggEAUqPlkiKfbgt6BL6wqcBe
# ZSYt2j4DIo2+lQB9V1hEi06lYhc0YWWFETj4MIPuRi4+dCC9WiKBGTd72e0L+/JQ
# sK07N1VlJxajnbgnlzXUIm8G8f7RDpClr8LPbWbj7piWlaDAQtJvNIsNE0bifrxO
# cyZYvbf4nVbm9HLx7mYlmtk9rvPJa6px/P6/40/PpDrOsArBTM5oECH385X8mfNX
# oPhQWQT/+JIG6KzTDQB3uZeatV+gxsSmW1D53MS9Wky1gvKaplNf42GCpfGx4iBn
# ahF4rFHRzCpTGO7+XYo4ySODzHMZ8WM5mlFHHTFLbNAjf05CXNDUvksNZjiv+T+y
# 6w==
# SIG # End signature block
