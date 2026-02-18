# Gmail Attachment Sync Script
# Downloads attachments from chatgptnotes@gmail.com
# Extracts ZIP files automatically

$gogPath = "$env:LOCALAPPDATA\gogcli\gog.exe"
$account = "chatgptnotes@gmail.com"
$outputDir = "C:\Users\NODE08\.openclaw\workspace\gmail-attachments"
$extractDir = "C:\Users\NODE08\.openclaw\workspace\gmail-attachments\extracted"
$stateFile = "C:\Users\NODE08\.openclaw\workspace\gmail-attachments\.last-sync.txt"

# Create directories if not exists
if (!(Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
if (!(Test-Path $extractDir)) { New-Item -ItemType Directory -Path $extractDir -Force | Out-Null }

# Get last sync time or default to 1 day
$lastSync = "1d"
if (Test-Path $stateFile) {
    $lastSync = Get-Content $stateFile -Raw
    $lastSync = $lastSync.Trim()
}

Write-Host "Checking emails newer than: $lastSync"

# Search for emails with attachments
$searchResult = & $gogPath gmail search "newer_than:$lastSync has:attachment" --account $account --max 50 --json 2>$null | ConvertFrom-Json

$newZipFiles = @()

if ($searchResult.threads -and $searchResult.threads.Count -gt 0) {
    Write-Host "Found $($searchResult.threads.Count) emails with attachments"
    
    foreach ($thread in $searchResult.threads) {
        $msgId = $thread.id
        Write-Host "Processing: $($thread.subject)"
        
        # Get message details
        $msgDetails = & $gogPath gmail get $msgId --account $account --json 2>$null | ConvertFrom-Json
        
        if ($msgDetails.attachments) {
            foreach ($att in $msgDetails.attachments) {
                $filename = $att.filename
                $attId = $att.attachmentId
                $outPath = Join-Path $outputDir $filename
                
                # Skip if already downloaded
                if (Test-Path $outPath) {
                    Write-Host "  Skipping (exists): $filename"
                    continue
                }
                
                Write-Host "  Downloading: $filename"
                & $gogPath gmail attachment $msgId $attId --account $account --out $outputDir --name $filename 2>$null
                
                # Check if it's a ZIP file
                if ($filename -match '\.zip$') {
                    $newZipFiles += @{
                        Path = $outPath
                        Name = $filename
                        From = $msgDetails.headers.from
                        Subject = $msgDetails.headers.subject
                    }
                }
            }
        }
    }
}
else {
    Write-Host "No new emails with attachments"
}

# Extract ZIP files
foreach ($zip in $newZipFiles) {
    $zipPath = $zip.Path
    $zipName = [System.IO.Path]::GetFileNameWithoutExtension($zip.Name)
    $extractTo = Join-Path $extractDir $zipName
    
    Write-Host "ZIP_FOUND: $($zip.Name)"
    Write-Host "FROM: $($zip.From)"
    Write-Host "SUBJECT: $($zip.Subject)"
    
    try {
        # Create extraction folder
        if (!(Test-Path $extractTo)) { New-Item -ItemType Directory -Path $extractTo -Force | Out-Null }
        
        # Extract
        Expand-Archive -Path $zipPath -DestinationPath $extractTo -Force
        
        Write-Host "EXTRACTED_TO: $extractTo"
        Write-Host "CONTENTS:"
        Get-ChildItem -Path $extractTo -Recurse | ForEach-Object {
            $relativePath = $_.FullName.Replace($extractTo, "").TrimStart("\")
            $size = if ($_.PSIsContainer) { "[DIR]" } else { "$([math]::Round($_.Length/1KB, 1)) KB" }
            Write-Host "  - $relativePath ($size)"
        }
    }
    catch {
        Write-Host "ERROR extracting $($zip.Name): $_"
    }
}

# Update last sync time
"30m" | Out-File $stateFile -NoNewline

Write-Host "Sync complete at $(Get-Date)"
