$dirRoot = 'c:\' # Change this as desired 
$filter = '*' # Change this as desired, e.g. *.log or *.txt 
 
function GetFolderSize($path){ 
    $total = (Get-ChildItem $path -ErrorAction SilentlyContinue -filter $filter | Measure-Object -Property length -Sum -ErrorAction SilentlyContinue).Sum 
    if (-not($total)) { $total = 0 } 
    $total 
    } # end function GetFolderSize 
 
# Entry point into script 
$results = @() 
$dirs = Get-ChildItem $dirRoot -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.psIsContainer} 
 
foreach ($dir in $dirs) { 
    
    $childFiles = @(Get-ChildItem $dir.pspath -ErrorAction SilentlyContinue -filter $filter| Where-Object{ -not($_.psIsContainer)}) 
    if ($childFiles) { $filecount = ($childFiles.count)} 
    else                     { $filecount = 0                  } 
 
    $childDirs = @(Get-ChildItem $dir.pspath -ErrorAction SilentlyContinue | Where-Object{ $_.psIsContainer}) 
    if ($childDirs ){ $dircount = ($childDirs.count)} 
    else                    { $dircount = 0                 } 
     
    $result = New-Object psobject -Property @{Folder = (Split-Path $dir.pspath -NoQualifier) 
                                              TotalSize = (GetFolderSize($dir.pspath)) 
                                              FileCount = $filecount; SubDirs = $dircount} 
    $results += $result 
    } # end foreach 
 
$results | Select-Object Folder, TotalSize , FileCount, SubDirs | Sort-Object TotalSize -Descending | Format-Table -auto 
Write-Host "Total of $($dirs.count) directories" 