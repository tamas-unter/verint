$filepattern="int*.log"
$pattern="9137936023080000201"
New-Item -ErrorAction SilentlyContinue -ItemType Directory ".\$pattern"; gci -Filter $filepattern |sls $pattern | Select -Unique Path | Copy-Item -Destination ".\$pattern"