cls
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"

$serverListPath = join-path $settingsPath 'Get-Service-List.txt'

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$now = get-date -format yyyy-MM-dd-HH-mm
$logFile = "Get-Service-" + $now
$logFile = Join-Path $logPath $logFile
$logFile = Initialize-Log $logPath $title
$services = @()

#get list of servers
$servers = Get-Content ($serverListPath)

#loop thru all servers
foreach ($server in $servers)
{
    $services += get-ciminstance win32_service -filter "startname <> 'LocalSystem' AND startname <> 'LocalService' AND startname <> 'NetworkService'" -comp $server  | where startname | Select PSComputerName, name,startname,state
    #$services += get-ciminstance win32_service -filter "State = 'Running' AND startname <> 'LocalSystem' AND startname <> 'LocalService' AND startname <> 'NetworkService'" -comp $server  | where startname | Select PSComputerName, name,startname 
}

#write services
foreach ($service in $services)
{
        Write-Log $logfile ",$($service.PSComputerName),$($service.name),$($service.startname),$($service.state)" 
}
