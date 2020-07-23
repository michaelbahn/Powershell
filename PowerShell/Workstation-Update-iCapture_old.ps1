$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = "..\Config"
$modulePath = "..\Scripts"
$startMenuPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs\REPP-Pre-Prod"
$jbShortcut = Join-Path $configPath "REPP-Pre-PROD"

$iCapturePath = "c$\iCapture"

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

#get list of workstations
$workstations = Get-Content  (join-path $configPath "workstation-update.txt")

Foreach ($workstation in $workstations) 
{   
    #copy J&B shortcust
    $workstationStartMenuPath = "\\$($workstation.Trim())\$($startMenuPath)"
    $result = Copy-Item -Path $jbShortcut -Destination $workstationStartMenuPath -Recurse -Force
    Write-Log $logfile "$($workstation): J&B_Pre-Prod added to Programs - $($result)"

    #copy icapture folders
    $iCapturePathLocal = join-path "\\$($workstation.Trim())" $iCapturePath
    if (! (Test-Path $iCapturePathLocal))
    {
        New-Item $iCapturePathLocal -ItemType Directory
        Write-Log $logfile "$($workstation): iCapture folder created"
    }
    $result = Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\PRE_PROD50 -Destination $iCapturePathLocal -Recurse -Force
    Write-Log $logfile "$($workstation): PRE_PROD50 copied to iCapture - $($result)"
    $result = Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\PRE_PROD51 -Destination $iCapturePathLocal -Recurse -Force
    Write-Log $logfile "$($workstation): PRE_PROD50 copied to iCapture - $($result)"

    $result = Invoke-Command -ComputerName $workstation -ScriptBlock {reg import "\\Dgvmopspd02\Deploy\iCapture\impression.reg" }
    Write-Log $logfile "$($workstation): impression.reg imported - $($result)"
    Invoke-Command -ComputerName $workstation -ScriptBlock {reg import "\\Dgvmopspd02\Deploy\iCapture\odbc-preprod.reg" }
    Write-Log $logfile "$($workstation): odbc-preprod.reg imported - $($result)"
}
         