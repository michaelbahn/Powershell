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
$deployPath = "\d$\Deployment\DE5617_iExportFix_05312019"
$rollbackFolderName = "DE5617_iExportFix_05312019"

#get list of servers
$servers = Get-Content  ("servers.txt")

foreach ($server in $servers) 
{   
    $deployPathServerFiles = "\\$($server.Trim())$($deployPath)"
     if ((Test-Path $deployPathServerFiles))
    {
        #Form files
        $deployServerFormdbFiles= join-path $deployPathServerFiles "FORMDB"

         if ((Test-Path $deployServerFormdbFiles))
        {
        #get rollbackfolder
            $rollbackFolderFormdb = "\\$($server.Trim())\d$\Backup_FORMDB\$($rollbackFolderName)"
             if (!(Test-Path $rollbackFolderFormdb))
            {
                New-Item -Path $rollbackFolderFormdb -ItemType Directory 
                Write-Log $logfile "Created rollback folder $($rollbackFolderFormdb)"
           }

            $newFiles = Get-ChildItem $deployServerFormdbFiles -File -Recurse
            foreach ($newFile in $newFiles)
            {
                $sourcePath = $newFile.DirectoryName
                $destinationPath = $sourcePath.Replace($deployPath, "")
                $destinationFileName = join-path $destinationPath $newFile.Name
                $rollbackPath = $destinationPath.Replace("FORMDB", "d$\Backup_FORMDB\$($rollbackFolderName)")

                if (Test-Path $destinationPath)
                {
                    if ((Test-Path $destinationFileName))
                    {
                    #copy to backup only if backup does not exist
                        if (!(Test-Path $rollbackPath))
                        {
                            New-Item -Path $rollbackPath -ItemType Directory 
                            Copy-Item -Path $destinationFileName -Destination $rollbackPath -Force
                            $newFileItem = get-ItemProperty (join-path $rollbackPath $newFile.Name)
                            $newFileList += $newFileItem
                            Write-Log $logfile "Rollback file $($newFile.Name) copied to $($rollbackPath)"
                        }                    
                        else
                        {
                            Write-Log $logfile "Rollback folder already exists and will not be overwritten: $($rollbackPath)"
                        }

                    }
                    else
                    {
                        Write-Log $logfile "No exisitng file to overwrite at $($destinationFileName)"
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
    }
    else
    {
        Write-Log $logfile "Error: $($deployPathServerFiles) folder does not exist"
    }

}   #end for

#send email with list of files 
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "iCapture Pre-Deployment Backup Completed: $($newFileList.Count) files"    

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

 