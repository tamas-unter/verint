function Get-WebexDetails{
param(
    $casenumber="",
    $description="",
    $invitee="",
    $date=(get-date -f "MM/dd"),
    $time=((get-date -Minute 0 -Second 0).addhours(1) |get-date -f "HH:mm"),
    $wxurl="https://verintinc.webex.com/meet/tamas.pasztor",
    $wxbridge="739 168 899",
    $duration=30
)
    Add-Type -AssemblyName System.Windows.Forms

    $f=New-Object system.Windows.Forms.Form
    $f.ClientSize='500,500';$f.Text="WEBEX";$f.BackColor='#efeffe'

    $labels=New-Object System.Windows.Forms.Label[] 8
    $inputs=New-Object System.Windows.Forms.TextBox[] 8
    0..7|%{
	    $labels[$_]=New-Object system.Windows.Forms.Label
	    $labels[$_].AutoSize=$true
	    $labels[$_].Font='Consolas,8'
	    $labels[$_].Width=25
	    $labels[$_].Height=10
	    $labels[$_].Location=New-Object System.Drawing.Point(20,(20+($_*30)))
	
	    $inputs[$_]=New-Object System.Windows.Forms.TextBox
	    $inputs[$_].AutoSize=$true
	    $inputs[$_].Font='Consolas,10'
	    $inputs[$_].Width=350
	    $inputs[$_].Height=10
	    $inputs[$_].Location=New-Object System.Drawing.Point(120,(20+($_*30)))
	    $inputs[$_].Text=""
    }
    $labels[0].Text='Case number:'
    $labels[1].Text='about:'
    $labels[2].Text='email:'
    $labels[3].Text='date:'
    $labels[4].Text='duration (min):'
    $labels[5].Text='webex URL:'
    $labels[6].Text='webex number:'

    $labels[3].Width=60
    $labels[7].Location=New-Object System.Drawing.Point(220,110)
    $labels[7].Width=60
    $labels[7].Text='time:'
    $f.controls.AddRange($labels)

    $inputs[0].Text=$casenumber
    $inputs[1].Text=$description
    $inputs[2].Text=$invitee
    $inputs[3].Text=$date
    $inputs[3].Width=80
    $inputs[4].Text=$duration
    $inputs[5].Text=$wxurl
    $inputs[6].Text=$wxbridge
    $inputs[7].Text=$time
    $inputs[7].Location=New-Object System.Drawing.Point(290,110)
    $inputs[7].Width=80
    #todo: 3 - datetimepicker, today, now, 30 mins etc..
    $f.controls.AddRange($inputs)

    $btnOk=new-object System.Windows.Forms.Button
    $btnOk.Text="Ok"
    $btnOk.Width=50
    $btnOk.Height=20
    $btnOk.Location=New-Object System.Drawing.Point(350,460)
    $btnOk.DialogResult=[System.Windows.Forms.DialogResult]::Ok

    $btnUpdate=new-object System.Windows.Forms.Button
    $btnUpdate.Text="Update"
    $btnUpdate.Width=50
    $btnUpdate.Height=20
    $btnUpdate.Location=New-Object System.Drawing.Point(160,460)
    $btnUpdate.DialogResult=[System.Windows.Forms.DialogResult]::None

    $btnCancel=new-object System.Windows.Forms.Button
    $btnCancel.Text="Cancel"
    $btnCancel.Width=50
    $btnCancel.Height=20
    $btnCancel.Location=New-Object System.Drawing.Point(420,460)
    $btnCancel.DialogResult=[System.Windows.Forms.DialogResult]::Cancel



    $rtfSubject=New-Object System.Windows.Forms.RichTextBox
    $rtfSubject.Location=New-Object System.Drawing.Point(20,230)
    $rtfSubject.Width=450
    $rtfSubject.height=220

    $btnUpdate.Add_Click({
        $caseNumber=$inputs[0].Text
        $issueDescription=$inputs[1].Text
        $hyperlink=$inputs[5].Text
        $bridge=$inputs[6].Text

        $rtfSubject.Rtf="{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Calibri;}{\f2\fswiss\fprq2\fcharset0 Calibri Light;}}
{\colortbl ;\red0\green122\blue255;\red47\green84\blue150;}
\pard\widctlpar\sa160\sl252\slmult1\f0\fs22\lang1038 Hi,\par
I\rquote ve set up this call to discuss \cf1\b VERINT\cf0\b0  case \b $caseNumber\b0 , where \emdash  as per my understanding \emdash  you are experiencing an issue with \i $issueDescription.\i0\par
\pard\keep\keepn\widctlpar\sb240\sl252\slmult1\cf2\f2\fs32\lang1033 Please join my personal Webex room:\par
\pard\widctlpar {\cf0\f0\fs22{\field{\*\fldinst{HYPERLINK $hyperlink }}{\fldrslt{$hyperlink}}}}\par 
\pard\widctlpar\cf0\ulnone $bridge\par}"
    })


    $f.Controls.AddRange(@($btnOk, $btnUpdate, $btnCancel, $rtfSubject))
    
    if (($f.ShowDialog()) -eq [System.Windows.Forms.DialogResult]::OK){
        $startDate=$inputs[3].Text
        $startTime=$inputs[7].Text
        $meetingDuration=$inputs[4].Text
        $start=($startDate|get-date)+($startTime|get-date).TimeOfDay
	    [PSCustomObject]@{
            'casenumber'=$inputs[0].Text;
            'topic'=$inputs[1].Text;
            'invitee'=$inputs[2].Text;
            'start'=$start;
            'end'=$start.AddMinutes($meetingDuration);
            'subject'=$rtfSubject.Rtf}
    }

}
function Create-WebexAppointment{
	begin{
		$ol=New-Object -ComObject outlook.application
#		$store=($ol.GetNamespace("MAPI")).GetStoreFromID("0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000054616D61732E5061737A746F7240766572696E742E636F6D002F6F3D45786368616E67654C6162732F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D30303433333630333366373534303533623837393235336338393632363137352D5061737A746F722C20546100E94632F4440000000200000010000000540061006D00610073002E005000610073007A0074006F007200400076006500720069006E0074002E0063006F006D0000000000")
#		$calendar=$store.GetDefaultFolder(9)
        $calendar=$ol.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
		$myself="Pasztor, Tamas"
		$location="my webex"
		$categories="planned"
        $meetingstatus=[Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
#Organizer                     : Pasztor, Tamas
#MessageClass                  : IPM.Appointment
        $reminderminutes=5
        $reminderset=$true
        $messageclass="IPM.Appointment"
	}

    process{
        if($_ -eq $null){return;}
        $a=$calendar.Application.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)
        $a.MeetingStatus=$meetingstatus
        $a.Categories=$categories
        $a.Location=$location

        $a.start=$_.start
        $a.end=$_.end
        $a.ReminderSet=$reminderset
        $a.ReminderMinutesBeforeStart=$reminderminutes
        $a.BusyStatus=[Microsoft.Office.Interop.Outlook.OlBusyStatus]::olBusy
        $a.Subject="verint case {0} - about {1}" -f $_.casenumber, $_.topic
        $a.Body=([System.Text.Encoding]::UTF8).GetBytes($_.subject)
        #olFormatHTML ??
        $a.BodyFormat=[Microsoft.Office.Interop.Outlook.OlBodyFormat]::olFormatRichText
        $r=$a.Recipients.Add($_.invitee)
        if($r.Resolve()){
            $a.Save()
            $a.Send()
        }
        #ConversationTopic             : teszt mihály
    }
    end{
#        foreach($ref in $r,$a,$calendar,$ol){
        foreach($ref in $calendar,$ol){
            if($ref -ne $null){[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    
    }
}
function Webex{
    param($casenumber,$about,$email)
    Get-WebexDetails $casenumber $about $email | Create-WebexAppointment
}