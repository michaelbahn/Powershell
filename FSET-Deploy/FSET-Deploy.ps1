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
$releasePath = "\d$\Deployment\$($releaseName)"
$deploymentDir = "d$\Deployment\$($releaseName)"
$destinationPath = "D$\Program Files\EDD\FSET Database Service"
$service = "FSET Database Service"

#get list of servers
$servers = Get-Content  servers.txt

foreach ($server in $servers) 
{   
    $deployPathServerFiles = "\\$($server.Trim())\$($deploymentDir)"
     if ((Test-Path $deployPathServerFiles))
    {
        $newFiles = Get-ChildItem $deployPathServerFiles -File -Recurse
        foreach ($newFile in $newFiles)
        {
            $destinationFullPath = "\\$($server.Trim())\$($destinationPath)"
            $destinationFileName = join-path $destinationFullPath $newFile.Name

            if (!(Test-Path $destinationFullPath ))
            {
                    New-Item -Path $destinationFullPath -ItemType Directory                     
                    Write-Log $logfile "Created folder $($destinationFullPath)"
            }

        #    $job = get-ciminstance win32_service -filter "Name='$($service)'" -comp $server | Invoke-CimMethod -Name StopService
            Copy-Item -Path $newFile.FullName -Destination $destinationFullPath -Force -ErrorAction Stop
        #    $job = get-ciminstance win32_service -filter "Name='$($service)'" -comp $server | Invoke-CimMethod -Name StartService
            $newFileItem = get-ItemProperty (join-path $destinationFullPath $newFile.Name)
            $newFileList += $newFileItem
            Write-Log $logfile "$($newFile.Name) deployed to $($destinationFullPath)"

        }   #end for
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
$subject = "FSET Deployment Completed: $($newFileList.Count) files"    

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

 