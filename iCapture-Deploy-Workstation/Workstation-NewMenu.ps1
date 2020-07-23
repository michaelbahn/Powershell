cls
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
$newFolderPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs\SKIPS"
#$rollbackSuffix = "_08082019"

#get name of file with target list of workstations from workstation-icapture-target.txt
$workstationListName = Get-Content  ("workstation-targets.txt")
$workstations = Get-Content  ($workstationListName)

foreach ($workstation in $workstations) 
{   
    Write-Log $logfile $workstation
    $workstationPath = "\\$($workstation.Trim())\$($newFolderPath)"

    $newFolder = New-Item -Path $workstationPath -ItemType Directory -Force
    if ($newFolder.Exists)
    {
        Write-Log $logfile "$($workstationPath) created."

        #copy file
        Copy-Item -Path $newFile  -Destination  $workstationPath  -Force
    }

}

    

         