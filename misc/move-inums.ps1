function get-pathfrominum{
param([string]$inum)
	"g:\calls\$($inum.substring(0,6))\$($inum.substring(6,3))\$($inum.substring(9,2))\$($inum.substring(11,2))\$inum"
}

ls| where name -match "^\d{15}"|
select -ExpandProperty name|%{
	[pscustomobject]@{
		source=$_ 
		dest="$(get-pathfrominum $_.split(".")[0]).$($_.split(".")[1,2] -join ".")"
	}
}|%{
	move-item  -ErrorAction SilentlyContinue $_.source $_.dest
}
