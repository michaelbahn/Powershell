$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$modulePath = "..\Scripts"
$newFileList = @()

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

#file to roll back
$newFilePath = "\\dgvmimgidxpd01\d$\Deployment\DE5617Changes04232019\Workstations\bin"
$newFileName = "ITISL.dll"
$newFile = join-path $newFilePath $newFileName
$rollbackSuffix = "_05032019"

#get version for local path
if (Test-Path "iCaptureVersion.txt")
{
    $iCaptureVersion = Get-Content "iCaptureVersion.txt"
    $iCapturePath = "c$\iCapture\$($iCaptureVersion.Trim())\bin"
}
else
{
   Write-Log $logfile "Error missing $($iCapturePath)"
   return
}

#get name of file with target list of workstations from workstation-icapture-target.txt
$workstationListName = Get-Content  ("workstation-icapture-target.txt")
$workstations = Get-Content  ($workstationListName)

foreach ($workstation in $workstations) 
{   

    $workstationPath = "\\$($workstation.Trim())\$($iCapturePath)"
     if (! (Test-Path $workstationPath))
    {
        Write-Log $logfile "$($workstationPath) folder does not exist"
    }

    #roll back file
    $destinationFileName = join-path $workstationPath $newFileName
    $rollbackFileName = "$($destinationFileName)$($rollbackSuffix)"

     if (Test-Path $rollbackFileName)
    {
        try{
            Remove-Item $destinationFileName -Force
            Rename-Item -Path $rollbackFileName -NewName $destinationFileName  -Force
            Write-Log $logfile "$($workstation): Rollback complete for $($destinationFileName)"
            $newFileItem = get-ItemProperty $destinationFileName
            $newFileList += $newFileItem
        }
        catch
         {
            Write-Log $logfile "Error $($workstation): $($destinationFileName) roll back failed)"
        }
    }
    else
    {
        Write-Log $logfile "Error: $($rollbackFileName) does not exist"
    }

}

#send email with list of files 
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "$($title) Completed: $($newFileList.Count) files"
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

         