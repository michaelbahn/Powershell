cls
#Initialize settings
$now = get-date
$today = get-date -Format  d
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$logPath = "..\Logs"

#initialize log file 
Import-Module .\Utilities.psm1 -Force 
$logFile = Initialize-Log $logPath $title

$process = Get-Process SolarWinds.ServiceHost.Process
$threshold = '314572800'

$mem = $process.ws

if ($mem -gt $threshold)
{
    Restart-Service -Name SolarWindsAgent64 -Force
    Write-Log $logFile "SolarWindsAgent64 restarted"       
}
else
{
    Write-Log $logFile "Memory check Passed"       
}


