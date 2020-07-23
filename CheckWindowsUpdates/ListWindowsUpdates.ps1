$dir = $MyInvocation.MyCommand.Path
$path  = Split-Path $dir
Set-Location  $path
$scriptPath = Join-Path $path CheckWindowsUpdates.ps1

#get list of servers
$servers = Get-Content -Path  ".\Servers.txt"

#loop thru each error check
foreach ($server in $servers) 
{   
    $updates = Invoke-Command -FilePath $scriptPath -ComputerName $server
    $updates
}