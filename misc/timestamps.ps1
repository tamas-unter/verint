Add-Type -AssemblyName System.IO.Compression.FileSystem
ls -Recurse *.zip|select fullname | %{
	$zip=[System.IO.Compression.ZipFile]::Open($_.fullname, [System.IO.Compression.ZipArchiveMode]::Read);
    $zipname=$_.FullName;
	$zip.Entries|select  LastWriteTime,name|%{
		[pscustomobject]@{timestamp="{0:yyyy-MM-dd HH:mm:ss}" -f $_.lastwritetime; zip=$zipname; file= $_.name
	}}
}|Export-Csv ".\test.csv"
