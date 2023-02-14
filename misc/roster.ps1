function Create-AppointmentsForToday{
	begin{
		$ol=New-Object -ComObject outlook.application
		#$store=($ol.GetNamespace("MAPI")).GetStoreFromID("0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000054616D61732E5061737A746F7240766572696E742E636F6D002F6F3D45786368616E67654C6162732F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D30303433333630333366373534303533623837393235336338393632363137352D5061737A746F722C20546100E94632F4440000000200000010000000540061006D00610073002E005000610073007A0074006F007200400076006500720069006E0074002E0063006F006D0000000000")
		#$calendar=$store.GetDefaultFolder(9)
        $calendar=$ol.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
		
        $myself="Pasztor, Tamas"
		$location="http://verintims.crm.dynamics.com"
		
	}
	process{
		$start=Get-Date -hour $_.start -minute 0 -second 0
		$end=Get-Date -hour $_.end -minute 0 -second 0
		$sequence=$_.watchers.IndexOf($myself) +1
		$watchers=($_.watchers | where -FilterScript{$_ -ne $myself} | %{($_ -split ',')[1]}) -join " + "
		$all_watchers=$_.watchers -join "`n"
		if($watchers -eq ""){
			$subject="TW ALONE";
		} else{
			$subject="TW #$sequence with $watchers"
		}
		$body="TicketWatch $start -- $end`n$all_watchers"
		
		$a=$calendar.Application.CreateItem(1)
		$a.MeetingStatus=0
		$a.Start=$start
		$a.End=$end
		$a.RequiredAttendees=$myself
		$a.ReminderSet=$false
		$a.Subject=$subject
		$a.Location=$location
		$a.Body=$body
		$a.Categories="TicketWatch"
		$a.Save()
		Write-Host ("{0:HH:mm}-{1:HH:mm}: #{2} with {3}" -f $start,$end,$sequence,$watchers)
	}
    end{
        Write-Host "Done."
    }
}

function Get-MyWatch{
	param($myself="Pasztor, Tamas")
	
	#today's file, copied from the middle block (select-all) of the ROSTER/SCHEDULES - textual page
	#$file=gc ("$location\roster{0}.txt" -f (Get-Date -f "yyyyMMdd"))
    # no longer need to have a file!!!
    $file=Get-Clipboard
	$namepattern="^[A-Z][a-z]* ?, [A-Z][a-z\.]+"
	$watchpattern="VAS_TKS_[T|R].* +(\d+):(\d+) ([AP])M - (\d+):(\d+) ([AP])M"
	
	#collecting everyone's schedule starting line numbers
	$staff=($file | sls $namepattern|%{[pscustomobject]@{'name'= $_.line; 'number'= $_.linenumber}})
	
	#parsing the watch events only - here AM PM is used
	$roster=($file|sls $watchpattern|
		%{[pscustomobject]@{
			'start'=[int]($_.matches[0].groups[1].value);
			'startpm'=($_.matches[0].groups[3].value -eq 'P');
			'end'=[int]($_.matches[0].groups[4].value);
			'endpm'=($_.matches[0].groups[6].value -eq 'P');
			'who'=($staff|where number -lt $_.linenumber|select -last 1 name).name
		}})
	
	#we convert the AM/PM hours to 24hr format
	$roster|%{if($_.startpm -and $_.start -ne 12) {$_.start+=12};if($_.endpm -and $_.end -ne 12) {$_.end+=12}}
	
	#logic to determine order of support (every even week alphabetical, odd weeks are Z-->A
	$az=([cultureinfo]::CurrentCulture.Calendar.GetWeekOfYear((get-date),'FirstDay', 'Monday') %2) -eq 0
	if (-not $az){$r=$roster|sort -Descending -Property 'who'}else {$r=$roster|sort -Property 'who'}
	
	#assuming each watch has exactly one hour duration
	#TODO: return end as well for longer durations -- the above assumption turns out to be wrong
    #output $my_watch=
	$r | where who -eq $myself | 
		select start, end|
			%{ [pscustomobject]@{
				'start'=$_.start;
                'end'=$_.end;
				'watchers'=@($r |where start -ge $_.start|where start -lt $_.end| select -ExpandProperty who)
			}}
}
Get-Mywatch | Create-AppointmentsForToday