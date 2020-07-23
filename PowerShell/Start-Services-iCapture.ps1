cls
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
Invoke-Expression " .\Set-Service.ps1 StartService Combined-Prod.csv"
