﻿cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
Invoke-Expression " .\Set-Service.ps1 StopService Services.csv"

