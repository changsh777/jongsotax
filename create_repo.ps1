$token = "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"
$body = '{"name":"macmini-bots","private":true,"description":"Mac Mini 텔레그램 봇 모음"}'
$r = Invoke-RestMethod -Uri "https://api.github.com/user/repos" `
    -Method Post `
    -Headers @{ Authorization = "token $token"; Accept = "application/vnd.github.v3+json" } `
    -ContentType "application/json" `
    -Body $body
Write-Host "생성완료: $($r.html_url)"
