$dir = $MyInvocation.MyCommand.Path
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$inputFileName = "$($title).txt"
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"

$inputFilePath = join-path $settingsPath $inputFileName
$outputFilePath = join-path $settingsPath $outputFileName

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-Log $logPath $title

#get list of services
$servers = Get-Content -Path  $inputFilePath

#loop thru all servers
foreach ($server in $servers)
{
    Invoke-Command -ComputerName $server -ScriptBlock {C:\Windows\System32\iisreset.exe }
    

    Write-Log $logfile "$($status.PSComputerName)`t$($status.state)`t$($status.name)" 
    $server.State = $status.State
    $server.StartMode = $status.StartMode
}

 $servers | Export-CSV $outputFilePath -NoTypeInformation