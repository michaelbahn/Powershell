$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$newFileList = @()

#log file settings
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt

#read settings
$deployPath = "\\dgvmimgidxpd01\d$\Deployment\DE5617Changes04232019"
$deployPathLength = $deployPath.Length
$rollbackSuffix = "_05032019"
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


#get list of servers
$servers = Get-Content  ("servers.txt")

foreach ($server in $servers) 
{   
    $deployPathServerFiles = "$($deployPath)\$($server.Trim())"
     if ((Test-Path $deployPathServerFiles))
    {
        #Roll back Form files
        $deployServerFormdbFiles= join-path $deployPathServerFiles "FORMDB"

         if ((Test-Path $deployServerFormdbFiles))
        {
            $newFiles = Get-ChildItem $deployServerFormdbFiles -File -Recurse
            foreach ($newFile in $newFiles)
            {
                $sourcePath = $newFile.DirectoryName
                $destinationPath = "\$($sourcePath.Substring($deployPathLength))"
                $destinationFileName = join-path $destinationPath $newFile.Name
                $rollbackFileName = "$($destinationFileName)$($rollbackSuffix)"

                if (Test-Path $destinationPath)
                {
                    if (Test-Path $rollbackFileName)
                    {
                        Remove-Item -Path $destinationFileName  -Force            
                        Rename-Item -Path $rollbackFileName  -NewName  $destinationFileName -Force
                        Write-Log $logfile "Rollback file restored: $($destinationFileName)"
                        $newFileItem = get-ItemProperty $destinationFileName
                        $newFileList += $newFileItem
                    }
                    else
                    {
                        Write-Log $logfile "No rollback file at $($destinationFileName)"
                    }
                }
                else
                {
                    Write-Log $logfile "Error - path does not exist: $($destinationPath) for $newFile"
                }
            }   #end for
        }
        else
        {
            Write-Log $logfile "No FORMDB files for $($server)"
        }

        #rollback bin files
        $deployServerBinFiles= join-path $deployPathServerFiles "bin"

         if ((Test-Path $deployServerBinFiles))
        {
            $newFiles = Get-ChildItem $deployServerBinFiles -File -Recurse
            foreach ($newFile in $newFiles)
            {
                $destinationPath = "\\$($server.Trim())\$($iCapturePath)"
                $destinationFileName = join-path $destinationPath $newFile.Name
                $rollbackFileName = "$($destinationFileName)$($rollbackSuffix)"

                if (Test-Path $destinationPath)
                {
                    if (Test-Path $rollbackFileName)
                    {
                        Remove-Item -Path $destinationFileName  -Force            
                        Rename-Item -Path $rollbackFileName  -NewName  $destinationFileName -Force
                        Write-Log $logfile "Rollback file restored: $($rollbackFileName)"
                        $newFileItem = get-ItemProperty $destinationFileName
                        $newFileList += $newFileItem
                    }
                    else
                    {
                        Write-Log $logfile "No rollback file at $($destinationFileName)"
                    }
                }
                else
                {
                    Write-Log $logfile "Error - path does not exist: $($destinationPath) for $newFile"
                }
            }   #end for
        }
        else
        {
            Write-Log $logfile "No bin files for $($server)"
        }
    }
    else
    {
        Write-Log $logfile "Error $($deployPathServerFiles) folder does not exist"
        break
    }
}   #end for

#send email with list of files deployed
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "$($title) Completed: $($newFileList.Count) files"
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       




         