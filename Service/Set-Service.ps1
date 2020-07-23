param  (
[string] $serviceAction = $(throw "serviceAction is required"),
[string] $csvFile = $(throw "CSV file name is required"))
    
    $dir = $MyInvocation.MyCommand.Path
    $scriptPath  = Split-Path $dir
    Set-Location  $scriptPath
    $modulePath = "..\Scripts"
    $logPath = "..\Logs"
    $mailBody = ""
    $serviceVerb = $serviceAction.Substring(0, $serviceAction.IndexOf("Service"))

    $emailSender = "teamops@edd.ca.gov"
    $emailRecipients  = Get-Content recipients.txt

    #initialize log file 
    Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
    $logFile = Initialize-Log $logPath $serviceAction

    #get list of services
    $servers = Import-Csv -Path  $csvFile
    Write-Log $logfile $serviceAction

    #loop thru all servers
    foreach ($server in $servers)
    {
        $job = get-ciminstance win32_service -filter "Name='$($server.Service)'" -comp $server.ServerName  | Invoke-CimMethod -Name $serviceAction
        switch ($job.ReturnValue)
        {
            0  {Write-Log $logfile "Service $($server.service) on $($server.ServerName): $($serviceAction) successful"; break;}
            1  {Write-Log $logfile "Service $($server.service) on $($server.ServerName): The request is not supported."; break;}
            2  {Write-Log $logfile "Service $($server.service) on $($server.ServerName): The user did not have the necessary access."; break;}
            3  {Write-Log $logfile "Service $($server.service) on $($server.ServerName): The service cannot be stopped because other services that are running are dependent on it."; break;}
            6  {Write-Log $logfile "Service $($server.service) on $($server.ServerName):  the service has not been started."; break;}
            7  {Write-Log $logfile "Service $($server.service) on $($server.ServerName): The service did not respond in a timely fashion"; break;}
            8 {Write-Log $logfile "Service $($server.service) on $($server.ServerName): Unknown failure"; break;}
            10  {Write-Log $logfile "Service $($server.service) on $($server.ServerName): The service is already running"; break;}
            default {Write-Log $logfile "Service $($server.service) on $($server.ServerName): Unknown return code $($job.ReturnValue)"; break;}
        }   
    }

    Start-Sleep 10

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
    Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $emailSender -To $emailRecipients -Subject "$($serviceAction) Completed" -Body $mailBody
