$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$modulePath = "..\Scripts"

$serverListPath = "Server-List.txt"

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-CSV $scriptPath "$($title)-All.csv"

#get list of services
$servers = Get-Content -Path  $serverListPath

$command = $task = Get-ScheduledTaskInfo -TaskName ZipAndDeleteLogFiles
$task.NextRunTime

#loop thru all servers
Write-CSV $logfile "Server,Service,Log On As,State" 

foreach ($server in $servers)
{
        $services =   get-ciminstance win32_service -comp $server
        
        foreach ($service in $services)
        {
            if (!(($service.StartName.ToUpper() -like "*LOCAL*" ) -or ($service.StartName.ToUpper() -like "*NETWORK*" )))
            {
                Write-CSV $logfile "$($service.PSComputerName),$($service.name),$($service.StartName),$($service.state)"
            }

        }
}


