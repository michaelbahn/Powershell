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
$deployPath = "\d$\Deployment\DE2503F_ful_06262019"
$rollbackFolderName = "DE2503F_ful_06262019"

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
            $newFiles = Get-ChildItem $deployServerFormdbFiles -File -Recurse
            foreach ($newFile in $newFiles)
            {
                $sourcePath = $newFile.DirectoryName
                $destinationPath = $sourcePath.Replace($deployPath, "")
                $destinationFileName = join-path $destinationPath $newFile.Name

                if (!(Test-Path $destinationPath))
                {
                        New-Item -Path $destinationPath -ItemType Directory                     
                        Write-Log $logfile "Created folder $($destinationPath)"
                }
                Copy-Item -Path $newFile.FullName -Destination $destinationPath -Force
                $newFileItem = get-ItemProperty $destinationFileName
                $newFileList += $newFileItem
                Write-Log $logfile "$($newFile.Name) deployed to $($destinationPath)"

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

#send email with list of files deployed
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "iCapture Deployment Completed: $($newFileList.Count) files"    
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

 