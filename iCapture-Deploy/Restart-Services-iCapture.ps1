cls
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
Invoke-Expression " .\Set-Service.ps1 StopService Services.csv"
Invoke-Expression " .\Set-Service.ps1 StartService Services.csv"
