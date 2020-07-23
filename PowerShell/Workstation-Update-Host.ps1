$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = "Workstation-Update-Host"
$configPath = "..\Config"
$modulePath = "..\Scripts"
$newHostFile = Get-Content  (join-path $configPath host-preprod.txt)
$hostPath = "c$\WINDOWS\system32\drivers\etc"
$startMenuPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs\REPP-PRE-PROD"
#$startMenuPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs"
$jbShortcut = Join-Path $configPath "REPP-Pre-PROD"

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

#get list of workstations to reboot
$workstations = Get-Content  (join-path $configPath workstation-update.txt)

Foreach ($workstation in $workstations) 
{
    $workstationPath = "\\$($workstation.Trim())\$($hostPath)"
    $workstationHostFile = Join-Path $workstationPath "hosts"

    If (!(test-path $workstationHostFile))   #create log folder if it doesn't exist
    {
        Copy-Item -Path $configPath -Destination $workstationPath -Name "hosts" -Force 
        Write-Log $logfile "$($workstation): host file created."
    }
    
    #append to host file
    Add-Content -Path $workstationHostFile -Value $newHostFile
    Write-Log $logfile "$($workstation): host file updated."

    #copy J&B shortcust
    $workstationStartMenuPath = "\\$($workstation.Trim())\$($startMenuPath)"
    Copy-Item -Path $jbShortcut -Destination $workstationStartMenuPath -Recurse -Force
    Write-Log $logfile "$($workstation): REPP-PRE-PROD added to Programs"
}
         