#0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000054616D61732E5061737A746F7240766572696E742E636F6D002F6F3D45786368616E67654C6162732F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D30303433333630333366373534303533623837393235336338393632363137352D5061737A746F722C20546100E94632F4440000000200000010000000540061006D00610073002E005000610073007A0074006F007200400076006500720069006E0074002E0063006F006D0000000000
#$storeID="0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000054616D61732E5061737A746F7240766572696E742E636F6D002F6F3D45786368616E67654C6162732F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D30303433333630333366373534303533623837393235336338393632363137352D5061737A746F722C20546100E94632F4440000000200000010000000540061006D00610073002E005000610073007A0074006F007200400076006500720069006E0074002E0063006F006D0000000000"
#
#
$ol=New-Object -ComObject outlook.application
$store=($ol.GetNamespace("MAPI")).GetStoreFromID("0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000054616D61732E5061737A746F7240766572696E742E636F6D002F6F3D45786368616E67654C6162732F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D30303433333630333366373534303533623837393235336338393632363137352D5061737A746F722C20546100E94632F4440000000200000010000000540061006D00610073002E005000610073007A0074006F007200400076006500720069006E0074002E0063006F006D0000000000")
$calendar=$store.GetDefaultFolder(9)
$hour=6
$start=Get-Date -hour $hour -minute 0 -second 0
$end=$start.AddHours(1)
$req="Pasztor, Tamas"
$sequence=0
$subject="TW $sequence with Gabor, Sachin"
$location="http://verintims.crm.dynamics.com"
$body="body placeholder"
$a=$calendar.Application.CreateItem(1)
$a.MeetingStatus=0
$a.Start=$start
$a.End=$end
$a.RequiredAttendees=$req
$a.ReminderSet=$false
$a.Subject=$subject
$a.Location=$location
$a.Body=$body
$a.Categories="TicketWatch"
$a.Save()