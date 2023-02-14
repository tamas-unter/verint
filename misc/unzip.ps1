Add-Type -AssemblyName System.IO.Compression.FileSystem
$zips=gci *.zip;$i=0;$zips|foreach {$i++;Write-Progress -Activity "unzipping $($_.Name) ($i/$($zips.Count))" -PercentComplete ($i/$zips.Count*100);[System.IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, (pwd).Path);Remove-Item $_.FullName}
