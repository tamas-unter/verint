$src="."
$dst=".\target"
$log=$dst\progress.log
$inums=$dst\inums.txt

$tmp=".\temp"

mkdir -p $dst
mkdir -p $tmp
$scope=ls -Recurse "$src\*.tar"

$i=0;
cd $tmp
$scope|%{
	$i++
	Write-Progress -Activity "extracting $($_.fullname)" -Status "$i/$($scope.count)" -PercentComplete ([int]($i/$($scope.count)*100)) 
	log "Processing $($_.fullname) ... " $true
	tar -x -f $_.fullname *.xml
	log "done" 
	move-filesaway
	#rm $_.fullname
}
cd ..
rm $tmp


function move-filesaway{
	$i=0
	ls|%{
		$p=$_.name|
			sls "(\d{6})(\d{3})(\d{2})(\d{2})(\d{2})"|
			%{$_.matches[0].groups.value}
		$dir="{1}\{2}\{3}\{4}" -f $p
		$target=mkdir -p $dst\$dir
		mv $_.name $target\
		$i++
	}
	log " -- $i files put into their directories"
}
function log{
	param($msg,$newline=$false)
	if($newline) {$msg|Out-File -append $log}
	else {$msg|Out-File -append -nonewline $log}
}