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
$newFileName = "DotNetPlugin.dll"
$newFilePath = "\\DGVMICAPIMGDV01\d$\Deployment\CPM\Development\Workstation\bin\dcr"
$newFile = join-path $newFilePath $newFileName 

$iCapturePath = "c$\iCapture\PROD51\bin\dcr"
$rollbackSuffix = "_11182019"

#get name of file with target list of workstations from workstation-icapture-target.txt
$workstationListName = Get-Content  ("workstation-targets.txt")
$workstations = Get-Content  ($workstationListName)

foreach ($workstation in $workstations) 
{   
    $workstationPath = "\\$($workstation.Trim())\$($iCapturePath)"
     if (Test-Path $workstationPath)
    {
        $destinationFileName = join-path $workstationPath $newFileName
        $rollbackFileName = "$($destinationFileName)$($rollbackSuffix)"

        if (Test-Path $destinationFileName)
        {
            Rename-Item -Path $destinationFileName -NewName $rollbackFileName -Force
        }
        else
        {
            Write-Log $logfile "No exisitng file to overwrite at $($destinationFileName)"
        }

        #copy files
        Copy-Item -Path $newFile  -Destination  $workstationPath  -Force
        Write-Log $logfile "$($workstation): $($newFile.Name) copied to to $($workstationPath)"
        $newFileItem = get-ItemProperty $destinationFileName
        $newFileList += $newFileItem
    }
    else
    {
        Write-Log $logfile "$($workstationPath) folder does not exist"
    }

}

#send email with list of files deployed
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "iCapture Workstation Deployment Completed: $($newFileList.Count) files"
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       
    

         