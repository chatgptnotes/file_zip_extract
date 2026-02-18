# Gmail New Email Notifier
# Checks for new emails and outputs summary for notification

$gogPath = "$env:LOCALAPPDATA\gogcli\gog.exe"
$account = "chatgptnotes@gmail.com"
$stateFile = "C:\Users\NODE08\.openclaw\workspace\gmail-attachments\.last-email-check.txt"

# Get last checked message ID or use time-based
$lastCheck = "15m"
if (Test-Path $stateFile) {
    $lastCheck = (Get-Content $stateFile -Raw).Trim()
}

# Search for new emails
$result = & $gogPath gmail search "newer_than:$lastCheck" --account $account --max 10 --json 2>$null | ConvertFrom-Json

if ($result.threads -and $result.threads.Count -gt 0) {
    Write-Host "NEW_EMAILS_FOUND"
    Write-Host "COUNT:$($result.threads.Count)"
    
    foreach ($thread in $result.threads) {
        $from = $thread.from -replace '<.*>', '' -replace '"', ''
        $subject = $thread.subject
        $id = $thread.id
        $hasAttachment = if ($thread.hasAttachment) { "[ATTACHMENT]" } else { "" }
        
        Write-Host "---"
        Write-Host "ID:$id"
        Write-Host "FROM:$($from.Trim())"
        Write-Host "SUBJECT:$subject $hasAttachment"
    }
}
else {
    Write-Host "NO_NEW_EMAILS"
}

# Update last check time
"15m" | Out-File $stateFile -NoNewline
