#Delete DIA attachment
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
Write-Host $title
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
$sender = "teamops@edd.ca.gov"
$recipients = "teamops@edd.ca.gov"

$folder = "\\dgvmimgdiapp01\ImportSource\DIAATTACHMT"

Get-ChildItem -Path $folder -File

$files = Get-ChildItem -Path $folder -File 
Write-Log $logFile "Removing $($files.count) files from $($folder)"
Get-ChildItem -Path $folder -File | Remove-Item -Force 
$files = Get-ChildItem -Path $folder -File 
Write-Log $logFile "$($files.count) files now in $($folder)"
$mailBody = Get-Content $logFile

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $title -Body $mailBody.ToString()

