$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = "..\Config"
$modulePath = "..\Scripts"
$importImpressionsScript = join-path $modulePath "Allow-Remote-Only-NLA.ps1"

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

$today = get-date -Format  "MMddyyyy"

#get list of workstations to reboot
$formPaths = Get-Content  (join-path $configPath "quarterChange.txt")

foreach ($formPath in $formPaths) 
{   
    $runtimeISL = join-path $formPath "runtime.isl"
    $destination = "$($runtimeISL)_$($today)"     #change this to before the dot
    Copy-Item -Path $runtimeISL -Destination  $destination -Force
    Write-Log $logfile "$($runtimeISL) copied to $($destination)"

 }


         