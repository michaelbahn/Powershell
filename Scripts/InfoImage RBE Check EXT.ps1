#Step 8 Verify ICM-RBEEXT share folders older than ½ hour.  
$title = "ICM-RBEEXT share folders older than 30 minutes"
$monitorFolder = '\\Dgvmextrbepd01\RBE01'
 $minutesOld = new-timespan -minutes 30
 $dayOld = new-timespan -days 1
$now = get-date
$dateTimeThresholdLow = $now.AddDays(-2)
$dateTimeThresholdHigh = $now.AddMinutes(-30)
cls
Write-Host $title  $monitorFolder
Write-Host "Check 8 low threshold: " $dateTimeThresholdLow
Write-Host "Check 8 high threshold: "$dateTimeThresholdHigh
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"

$outputFile = Join-Path $settingsPath test.txt
$sender = gc  (join-path $settingsPath Sender.txt)
$recipients = gc  (join-path $settingsPath Recipients.txt)

IF ([string]::IsNullOrWhitespace($sender))
{
    Write-Host "Invalid sender list at " $settingsPath\Sender.txt
    return
} 

IF ([string]::IsNullOrWhitespace($recipients))
{
    Write-Host "Invalid receipient list at " $settingsPath\Recipients.txt
    return
} 
 $monitorFolders  = $null
$monitorFolders  =  Get-ChildItem -Path $monitorFolder -Recurse -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $dateTimeThresholdLow  -and $_.LastWriteTime -le $dateTimeThresholdHigh } | Select-Object FullName, LastWriteTime

if ($monitorFolders  -ne $null)
{
    $monitorFolders | ConvertTo-Html  -Title $title -Property FullName , LastWriteTime | Out-File $outputFile
    $mailBody = Get-Content $outputFile  -Raw
    Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $sender -To $recipients -Subject $title -Body $mailBody -BodyAsHtml
}
else
{
    Write-Host "Check 8 passed: no RBE EXT Folders"
}
