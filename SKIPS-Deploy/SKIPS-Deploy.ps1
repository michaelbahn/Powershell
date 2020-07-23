$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$newFileList = @()
$deploy = $true
$backup = $false

#log file settings
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt

#read settings
$sourcePath = "\\entdbshare1\EntRelease\DMS_DMRC\SKIPS\WebApp"
# do not copy web.config
$destinationPath = "d$\inetpub\wwwroot\SKIPSWebApp"
$backupFolderName = "d$\backup\052922019"

#get list of servers
$servers = Get-Content  ("servers.txt")

foreach ($server in $servers) 
{   
     if (Test-Path $sourcePath)
    {
        $job = get-ciminstance win32_service -filter "Name='IISAdmin'" -comp $server | Invoke-CimMethod -Name 'StopService' 

        $destinationServerPath = "\\$($server)\$destinationPath"
        #backup files
        $backupPath = "\\$($server)\$backupFolderName"

         if (!(Test-Path $backupPath))
        {
            New-Item -Path $backupPath -ItemType Directory
        }

        Copy-Item -Path $destinationServerPath -Destination $backupPath -Recurse -Force
        Remove-Item -Path $destinationServerPath -Recurse -Force
        Copy-Item -Path $sourcePath  -Destination $destinationServerPath -Recurse -Force

        $job = get-ciminstance win32_service -filter "Name='IISAdmin'" -comp $server | Invoke-CimMethod -Name 'StartService' -AsJob

    }
    else
    {
        Write-Log $logfile "Error - source path does not exist: $($sourcePath)"
    }
}   #end for


$subject = $title
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

 