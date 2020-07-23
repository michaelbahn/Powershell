cls
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$mailBody = ""

$emailSender = "teamops@edd.ca.gov"
$emailRecipients  = Get-Content (join-path $settingsPath "emailServicesRecipients.txt")

$serverListPath = join-path $settingsPath 'Services-Test.csv'  #'Services-PFLDIA.csv'

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-Log $logPath $title

#get list of services
$servers = Import-Csv -Path  $serverListPath
Write-Log $logfile "$(get-date) `t Starting Services" 

#loop thru all servers
foreach ($server in $servers)
{
    $job = get-ciminstance win32_service -filter "Name='$($server.Service)'" -comp $server.ServerName | Invoke-CimMethod -Name StartService
    switch ($job.ReturnValue)
    {
        0  {Write-Log $logfile "$(get-date) `t  Service $($server.service) started on $($server.ServerName)"; break;}
        6  {Write-Log $logfile "$(get-date) `t  Service $($server.service) not started on $($server.ServerName)"; break;}
        7  {Write-Log $logfile "$(get-date) `t  Service $($server.service) on $($server.ServerName) did not respond in a timely fashion"; break;}
        8 {Write-Log $logfile "$(get-date) `t  Unknown response from Service $($server.service) on $($server.ServerName)"; break;}
        10  {Write-Log $logfile "$(get-date) `t  Service $($server.service) already started on $($server.ServerName)"; break;}
        default {Write-Log $logfile "$(get-date) `t  Unknown return code $($job.ReturnValue) from Service $($server.service) on $($server.ServerName): $($job.ReturnValue)"; break;}
    }   
}

#Start-Sleep 60

Write-Log $logfile "-----------`t`t`t`t`t --------`t`t ------------------"
Write-Log $logfile "Server  `t`t`t`t`t`t  State `t`t`t Service" 
Write-Log $logfile "-----------`t`t`t`t`t --------`t`t ------------------" 

#check
foreach ($server in $servers)
{
        $status =   get-ciminstance win32_service -filter "Name='$($server.Service)'" -comp $server.ServerName 
        Write-Log $logfile "$($status.PSComputerName)`t$($status.state)`t$($status.name)" 
}

$mailBody = Get-Content $logfile.FullName | Out-String
Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $emailSender -To $emailRecipients -Subject "$($title) Completed" -Body $mailBody
