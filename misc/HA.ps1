$token="eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJjZjU1YTU3Y2E4ZjM0MGQ3YTA0MmNlMzkxYTNlMTYzYyIsImlhdCI6MTYxMjQ1NTkwNywiZXhwIjoxOTI3ODE1OTA3fQ.XI08Lf3oYT4bM6pWn1RSfjRxYpyGkJMknxUEowm1BUo"
$hdrs=@{"Authorization"="Bearer $token";"Content-Type"="application/json"}
$address="http://messenger.local:8123"
$service="/api/states/light.dolgozo_rgb"
curl -Headers $hdrs "$address$service"|select -ExpandProperty Content
