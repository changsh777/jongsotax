$token = "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"
$headers = @{ Authorization = "token $token"; Accept = "application/vnd.github.v3+json" }

# 유저 repos 전체 조회
$repos = Invoke-RestMethod -Uri "https://api.github.com/user/repos?per_page=100" -Headers $headers
$repos | ForEach-Object { Write-Host $_.full_name }
