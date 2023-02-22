$myself="Pasztor, Tamas"
#$myself="Reddy, Bavanasi"


function Create-AppointmentsForToday{
	begin{
		$ol=New-Object -ComObject outlook.application
        $calendar=$ol.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
		$location="Red Zone"
	}
	process{
		$start=Get-Date -hour $_.from -minute 0 -second 0
		$end=Get-Date -hour $_.to -minute 0 -second 0
		#$sequence=$_.watchers.IndexOf($myself) +1
		$watchers=$_.with -join " + "
		
		if($_.with.Count -eq 0){
			$subject="TW ALONE";
		} else{
			$subject="TW #$($_.sequence) with $watchers"
		}

		$body=@"
TicketWatch $start -- $end

Also watching: $($_.with -join " + ")
Previous Hour: $($_.before -join " + ")
Next Hour: $($_.after -join " + ")
"@
		$a=$calendar.Application.CreateItem(1)
		$a.MeetingStatus=0
		$a.Start=$start
		$a.End=$end
		$a.RequiredAttendees=$myself
		$a.ReminderSet=$true
        $a.ReminderMinutesBeforeStart=0
		$a.Subject=$subject
		$a.Location=$location
		$a.Body=$body
		$a.Categories="TicketWatch"
		$a.Save()
		Write-Host ("{0:HH:mm}-{1:HH:mm}: #{2} with {3}" -f $start,$end,$_.sequence,$watchers)
	}
    end{
        Write-Host "."
    }
}


$az=([cultureinfo]::CurrentCulture.Calendar.GetWeekOfYear((get-date),'FirstDay', 'Monday') %2) -eq 0

#$raw=cat "C:\Users\tpasztor\OneDrive - Verint Systems Ltd\Desktop\tw_20230214.txt"
$raw=clipboard
$namepattern="^[A-Z][a-z]* ?, [A-Z][a-z\.]+"
$watchpattern=" VAS_TKS_[T|R].*\s+(\d+):(\d+) ([AP])M - (\d+):(\d+) ([AP])M"


$roster=$raw|sls "($namepattern)\tVAS_RECORDER_AMER" -context 0,10|%{
    $e=$_.context.postcontext
    $next=$e|sls "$namepattern"|select -first 1 -ExpandProperty linenumber
    if($next -eq $null){ $next=11;}
    if($next -lt 2){ $next =2;}
    $next--
    $watches=$e[0..$next]|sls "$watchpattern"|%{
        $m=$_.matches[0].groups.value
        
        if($m[3] -eq "P"){$start_p=12} else {$start_p=0}
        if($m[6] -eq "P"){$end_p=12} else {$end_p=0}
        
        $t1=([int]$m[1] % 12) + $start_p
        $t2=([int]$m[4] % 12) + $end_p - 1
        if($t1 -gt $t2){
            0..$t2|%{
                [pscustomobject]@{
                    start=$_
                    end=$_ + 1
                }
            }
            $t1..23|%{
                [pscustomobject]@{
                    start=$_
                    end=$_ + 1
                }
            }
        } else {
            $t1..$t2|%{
                [pscustomobject]@{
                    start=$_
                    end=$_ + 1
                }
            }
        }
    }

    [pscustomobject]@{
        name=$_.matches[0].groups[1].value
        watches=$watches
    }

}
$p=@{Descending=-not $az;Property="name"}

$roster=$roster|sort @p

$myindex=$roster.name.IndexOf($myself)

$mywatches=@()
$roster|where name -eq $myself|select -ExpandProperty watches|%{
    $s=$_.start
    $e=$_.end
    
    $before=$with=$after=$withindices=@()
    $roster|where name -ne $myself|%{
        $entry=$_
        $_.watches|%{
            if($_.start -le $s -and $_.end -ge $e){
                $with+="{2} {0}" -f $entry.name.Split(", ")
                $withindices+=$roster.name.IndexOf($entry.name)
            }
            if($_.start -lt $s -and $_.end -ge $s){
                $before+="{2} {0}" -f $entry.name.Split(", ")
            }
            if($_.start -le $e -and $_.end -gt $e){
                $after+="{2} {0}" -f $entry.name.Split(", ")
            }
        }
    }
    $mywatches+=([pscustomobject]@{
                        from=$s
                        to=$e
                        with=$with
                        before=$before
                        after=$after
                        sequence=($withindices|where {$_ -lt $myindex}).Count + 1
    })
}

$mywatches|Create-AppointmentsForToday
