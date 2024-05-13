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
        default {"⭐"}
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

$r=Invoke-RestMethod -Method Post -Uri $uri -Body ($cred|ConvertTo-Json) -Headers @{host=$wfo} -ContentType 'application/json'
#$r=@{AuthToken=@{id=10721; token="MwMXnqq299g4csam"; extendSession=""}}

$token=$r.AuthToken.token

$employees=@()
$schedules=@()
$startTime=get-date -Format "yyyy-MM-ddT00:00:00Z"
$endTime=(get-date).AddDays(1)| get-date -Format "yyyy-MM-ddT00:00:00Z"

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


$az= -not (([cultureinfo]::CurrentCulture.Calendar.GetWeekOfYear((get-date),'FirstDay', 'Monday') %2) -eq 0)
$d=get-date -Hour 0 -Minute 1
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
}|Out-GridView





#### create the outlook items
$mywatches|Create-AppointmentsForToday

