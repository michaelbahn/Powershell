$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$modulePath = "..\Scripts"

#log file settings
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt
$newFileList = @()

#files to push out
#$newFilePath = "D:\iEditor 5.1"
#$newFiles = Get-ChildItem $newFilePath -Recurse 
$newFile = "C:\Users\mbahn-22\Desktop\SKIPS Prod.url"
$iCapturePath = "c$\Users\Public\Desktop"
#$rollbackSuffix = "_08082019"

#get name of file with target list of workstations from workstation-icapture-target.txt
$workstationListName = Get-Content  ("workstation-targets.txt")
$workstations = Get-Content  ($workstationListName)

foreach ($workstation in $workstations) 
{   
    $workstationPath = "\\$($workstation.Trim())\$($iCapturePath)"
     if (! (Test-Path $workstationPath))
    {
        Write-Log $logfile "$($workstationPath) folder does not exist"
    }

    #copy files
    Copy-Item -Path $newFile  -Destination  $workstationPath  -Force
}

    

         