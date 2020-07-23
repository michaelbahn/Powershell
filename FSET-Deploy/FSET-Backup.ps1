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
$releaseName = "Edd.Dms.Fset.DatabaseService_05312019"
$releasePath = "\Deployment\$($releaseName)"
$destinationPath = "D$\Program Files\EDD\FSET Database Service"

#get list of servers
$servers = Get-Content  ("servers.txt")

foreach ($server in $servers) 
{   
    $deployPathServerFiles = "\\$($server.Trim())\Deployment\$($releaseName)"
     if ((Test-Path $deployPathServerFiles))
    {
        #get rollbackfolder
        $rollbackFolderName = "\\$($server.Trim())\d$\FSETBackup\$($releaseName)"
            #copy to backup only if backup does not exist
            if (!(Test-Path $rollbackFolderName))
            {
                New-Item -Path $rollbackFolderName -ItemType Directory 
                Write-Log $logfile "Created rollback folder $($rollbackFolderName)"

                $newFiles = Get-ChildItem $deployPathServerFiles -File -Recurse
                foreach ($newFile in $newFiles)
                {
                    $sourcePath = $newFile.DirectoryName
                    $destinationFullPath = "\\$($server.Trim())\$($destinationPath)"
                    $destinationFileName = join-path $destinationFullPath $newFile.Name

                    if (Test-Path $destinationFullPath )
                    {
                        if ((Test-Path $destinationFileName))
                        {
                            Copy-Item -Path $destinationFileName -Destination $rollbackFolderName -Force
                            $newFileItem = get-ItemProperty (join-path $rollbackFolderName $newFile.Name)
                            $newFileList += $newFileItem
                            Write-Log $logfile "Rollback file $($newFile.Name) copied to $($rollbackFolderName)"
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
            Write-Log $logfile "Rollback folder already exists and will not be overwritten: $($rollbackFolderName)"
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
$subject = "FSET Pre-Deployment Backup Completed: $($newFileList.Count) files"    

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

 