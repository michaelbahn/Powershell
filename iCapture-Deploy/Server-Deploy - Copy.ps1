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
        #copy Form files
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
                    #make backup of current file
                    if (Test-Path $rollbackFileName)
                    {
                        Write-Log $logfile "$($rollbackFileName) already exists"
                    }
                    elseif (Test-Path $destinationFileName)
                    {
                        Rename-Item -Path $destinationFileName -NewName  $rollbackFileName -Force
                        Write-Log $logfile "Rollback file copied: $($rollbackFileName)"
                    }
                    else
                    {
                        Write-Log $logfile "No exisitng file to overwrite at $($destinationFileName)"
                    }

                    Copy-Item $newFile.FullName $destinationPath -Force            
                    $newFileItem = get-ItemProperty $destinationFileName
                    $newFileList += $newFileItem
                    Write-Log $logfile "$($newFile) copied to $($workstationPath)"            
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

        #copy bin files
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
                    #make backup of current file
                    if (Test-Path $rollbackFileName)
                    {
                        Write-Log $logfile "$($rollbackFileName) already exists"
                    }
                    elseif (Test-Path $destinationFileName)
                    {
                        Rename-Item -Path $destinationFileName -NewName  $rollbackFileName -Force
                        Write-Log $logfile "Rollback file copied: $($rollbackFileName)"
                    }
                    else
                    {
                        Write-Log $logfile "No exisitng file to overwrite at $($destinationFileName)"
                    }

                    Copy-Item $newFile.FullName $destinationPath -Force            
                    $newFileItem = get-ItemProperty $destinationFileName
                    $newFileList += $newFileItem
                    Write-Log $logfile "$($newFile) copied to $($destinationFileName)"            
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
        Write-Log $logfile "Error: $($deployPathServerFiles) folder does not exist"
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

