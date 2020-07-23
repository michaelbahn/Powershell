cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
Invoke-Expression  ".\Windows-Update-History.ps1 Server-List-Error.txt"
