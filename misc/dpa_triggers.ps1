
Add-Type -AssemblyName System.IO.Compression

Add-Type -AssemblyName System.IO.Compression.FileSystem

# task 1 - find all DPA calls in the RIS log
# determine scope
# select origin - TODO populate automatically
$archive=@{
    path="C:\Users\tpasztor\Downloads\1179192\0921\RIS1 PRI\integrationservice_2021_09_14_4701 (2).zip";
}
$outputFile="C:\Users\tpasztor\Downloads\1179192\0921\dpa_caseid_requests.log"
# create temporary folder
$temp="C:\Users\tpasztor\Downloads\1179192\0921\temp"
mkdir -ErrorAction Continue $temp
cd $temp
# open the archive
[System.IO.Compression.ZipArchive]$a=
    [System.IO.Compression.ZipFile]::OpenRead($archive.path)
        
# loop through the archive
for($i=0;$i -le $a.Entries.Count;$i+=5){
	Write-Progress -id 1 -Activity "BIG archive unzip" -Status "$($_.Name) ($i/$($a.Entries.Count))" -PercentComplete ($i/$a.Entries.Count*100);

    # unzip files to temporary folder
    try{
        $a.Entries[($i)..($i+4)]|%{
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_,"$(pwd)\$($_.fullname)",$true)
        }
        try{
       	    $zips=ls *.zip;
    	    $j=0
	        $zips|%{
		        $j++
		        Write-Progress -ParentId 1 -Activity "Unzipping" -Status "$($_.Name) ($j/$($zips.Count))" -PercentComplete ($j/$zips.Count*100);
			    [System.IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, (pwd).Path);
			    Remove-Item $_.FullName
            }
            Write-Progress -ParentId 1 -Activity "Parsing" -Status "startted" -PercentComplete 0;

            ls *.log|sls "\[connectser.*deliver.*caseid"|Out-File -Append $outputFile -Width 500
            rm *.log
            Write-Progress -ParentId 1 -Activity "Parsing" -Status "done" -PercentComplete 100;
        }
        catch{Write-Host "error parsing"}
	}
	catch{Write-Host "error unzipping" }
    
}
$a.Dispose()
# remove temporary folder
cd ..
rm -ErrorAction Continue -Recurse $temp
# task 2 - determine user info and compare login information