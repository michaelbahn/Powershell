cls
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$emailSender = "teamops@edd.ca.gov"
$emailRecipients  = Get-Content (join-path $settingsPath "emailServerRebootRecipients.txt")
$getPendingRebootScript = join-path $modulePath "Get-PendingRebootProperty.ps1" 

#list of servers to reboot
$serverListPath = join-path $settingsPath 'Server-Reboot.txt'

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-Log $logPath $title

#get list of servers to reboot
$servers = Get-Content ($serverListPath)

$pendingReboot = @()
Foreach ($server in $servers) 
{
    $reboot = Invoke-Command -FilePath $getPendingRebootScript -ComputerName $server
    if ($reboot) {
        $pendingReboot += $server
        Write-Log $logfile $server
    }
}

$pendingReboot