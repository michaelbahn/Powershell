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
$workstationTargetPath = "c$\Users\Public\Desktop"
#$rollbackSuffix = "_08082019"

#get name of file with target list of workstations 
#$workstationListName = Get-Content  
$workstations = Get-Content  .\workstation-training.txt

foreach ($workstation in $workstations) 
{   
    $workstationPath = "\\$($workstation.Trim())\$($workstationTargetPath)"
     if (! (Test-Path $workstationPath))
    {
        Write-Log $logfile "$($workstationPath) folder does not exist"
    }

    #copy files
    Copy-Item -Path $newFile  -Destination  $workstationPath  -Force
}

    
#send email with list of files deployed
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "Workstation Deployment Completed: $($newFileList.Count) files"
#Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

         