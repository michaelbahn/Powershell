cls
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$emailSender = "teamops@edd.ca.gov"
$emailRecipients  = Get-Content (join-path $settingsPath "emailServerRebootRecipients.txt")

$serverListPath = join-path $settingsPath 'Server-Reboot.txt'

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$now = get-date -format yyyy-MM-dd-HH-mm
$logFile = "Server-Reboot-" + $now + ".log"
$logFile = Join-Path $logPath $logFile
$logFile = Initialize-Log $logPath $title
Write-Log $logfile "$($title) started: $($now)"
$jobs = @()
$log = $null

#get list of servers to reboot
$servers = Get-Content ($serverListPath)

Foreach ($server in $servers) 
{
    $PingRequest = Test-Connection -ComputerName  $server -Count 2  -Quiet

    if ($PingRequest)
     { 
        try
        {
            $jobs += Restart-Computer -ComputerName $server -AsJob -Force  
            Write-Log $logfile "Restart job created for $($server)"
        }
        catch
        {
            Write-Log $logfile "Restart job failed for $($server): $($_.Exception.Message)"
        }

     }      
    else          
    {
        $log += "$($server): Ping failed"
    }
}

Start-Sleep 300

Foreach ($job in $jobs) 
{
    $log += "$($job.PSBeginTime)  $($job.Location) restart: $($job.State) $($job.PSWEndTime)`r`n" 
}

Write-Log $logfile $log
Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $emailSender -To $emailRecipients -Subject "$($title) Completed" -Body $log


         