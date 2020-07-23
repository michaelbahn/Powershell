$dir = $MyInvocation.MyCommand.Path
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$inputFileName = "$($title)-Other.csv"
$outputFileName = "$($title)-Results-Other.csv"
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"

$inputFilePath = join-path $settingsPath $inputFileName
$outputFilePath = join-path $settingsPath $outputFileName

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-Log $logPath $title

#get list of services
$servers = Import-Csv -Path  $inputFilePath

#loop thru all servers
foreach ($server in $servers)
{
    $status =   get-ciminstance win32_service -filter "Name='$($server.Service)'" -comp $server.ServerName 
    Write-Log $logfile "$($status.PSComputerName)`t$($status.state)`t$($status.name)" 
    $server.State = $status.State
    $server.StartMode = $status.StartMode
}

 $servers | Export-CSV $outputFilePath -NoTypeInformation