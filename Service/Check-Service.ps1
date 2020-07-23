cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$modulePath = "..\Scripts"

$serverListPath = "Server-List-All.txt"
##$serverListPath = "Server-List-One.txt"

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-CSV $scriptPath "$($title)-All.csv"

#get list of services
$servers = Get-Content -Path  $serverListPath
$serviceName = "CyveraService"

#loop thru all servers
Write-CSV $logfile "Server,Service,Log On As,State" 

foreach ($server in $servers)
{
        $server = $server.Trim()
        $service =   get-ciminstance win32_service -comp $server -Filter "Name like '$($serviceName)%'"
        
        
        if (!([string]::IsNullOrEmpty($service)))
        {        
            Write-CSV $logfile "$($service.PSComputerName),$($service.name),$($service.StartName),$($service.state)"
        }
        else
        {        
            Write-CSV $logfile "$($server),Missing"
        }

        
}


