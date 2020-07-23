$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = "..\Config"
$modulePath = "..\Scripts"

$startMenuPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs"
$iCaptureShortcut = "\\dgvmopspd02\Deploy\iCapture\iCapture-Pre-Prod*"
$iCapturePath = "c$\iCapture"

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

#get list of workstations to update
$workstations = Get-Content  (join-path $configPath "workstation-update.txt")

foreach ($workstation in $workstations) 
{   
    #copy iCapture shortcust
    $workstationStartMenuPath = "\\$($workstation.Trim())\$($startMenuPath)"
    $result = Copy-Item -Path $iCaptureShortcut -Destination $workstationStartMenuPath -Recurse -Force
    Write-Log $logfile "$($workstation): iCapture Pre-Prod shortcuts added"

    #copy icapture folders
    $iCapturePathLocal = join-path "\\$($workstation.Trim())" $iCapturePath
    if (! (Test-Path $iCapturePathLocal))
    {
        New-Item $iCapturePathLocal -ItemType Directory
        Write-Log $logfile "$($workstation): iCapture folder created"
    }
    Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\PRE_PROD50 -Destination $iCapturePathLocal -Recurse -Force
    Write-Log $logfile "$($workstation): PRE_PROD50 copied to $($iCapturePathLocal)"
    Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\PRE_PROD51 -Destination $iCapturePathLocal -Recurse -Force
    Write-Log $logfile "$($workstation): PRE_PROD51 copied to $iCapturePathLocal"   

}
         