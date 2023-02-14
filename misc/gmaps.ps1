$lat=Read-Host "latitude:"
$lon=Read-Host "longitude:"
start "https://www.google.com/maps/search/?api=1&query=$lat,$lon&zoom=15"
