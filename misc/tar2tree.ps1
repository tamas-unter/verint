$src="F:\2020"
$dst="F:\target_2020"
$log="$dst\progress.log"

$tmp="F:\_"

#we're in the temp folder. let's distribute the inums into their respective target directories
function move-filesaway{
	$i=0
    $time_move=Measure-command{
	    ls|%{
		    $p=$_.name|
			    sls "(\d{6})(\d{3})(\d{2})(\d{2})(\d{2})"|
			    %{$_.matches[0].groups.value}
		    $dir="{1}\{2}\{3}\{4}" -f $p
		    $target=mkdir -p $dst\$dir -ErrorAction SilentlyContinue
            if($target.length -eq 0){$target="$dst\$dir"}
		    mv $_ $target\
		    $_.name.split(".")[0]|
                sls "(\d{6})(\d{9})"|%{
                    $_.line|
                        out-file -Append "$dst\inums_$($_.matches[0].groups[1].value).txt"
                }
		    $i++
	    }
    }
	log " -- $i files in $($time_move.TotalSeconds) seconds" $true
}
function log{
	param($msg,$newline=$false)
	if($newline) {$msg|Out-File -append $log}
	else {$msg|Out-File -append -nonewline $log}
}


mkdir -p $dst
mkdir -p $tmp
$scope=ls -Recurse "$src\*.tar"

$i=0;
cd $tmp
$scope|%{
	$i++
	Write-Progress -Activity "extracting $($_.fullname)" -Status "$i/$($scope.count)" -PercentComplete ([int]($i/$($scope.count)*100)) 
	log "$("{0:yyyy-MM-dd_HH:mm:ss}" -f (get-date)) $($_.fullname) ... " $false
    $time_untar=Measure-Command{	
        tar -x -f $_.fullname *.xml
    }
	log "done." $false
	move-filesaway
    rm $tmp\*
}
cd ..

