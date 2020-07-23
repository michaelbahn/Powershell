$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = Resolve-Path "..\Config"
$modulePath = Resolve-Path "..\Scripts"
$startMenuPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs\REPP-Pre-Prod"
$jbShortcut = Join-Path $configPath "REPP-Pre-PROD"
$regScript = join-path $modulePath "reg_import.ps1"

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
    Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\PRE_PROD50 -Destination $iCapturePathLocal -Recurse -Force
    Write-Log $logfile "$($workstation): PRE_PROD50 copied to iCapture"
    Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\PRE_PROD51 -Destination $iCapturePathLocal -Recurse -Force
    Write-Log $logfile "$($workstation): PRE_PROD50 copied to iCapture"
    Copy-Item -Path \\Dgvmopspd02\Deploy\iCapture\*.reg -Destination $iCapturePathLocal -Force
    Invoke-Command -ComputerName $workstation -ScriptBlock {reg import c:\iCapture\impression.reg" }
    Invoke-Command -ComputerName $workstation -ScriptBlock {reg import c:\iCapture\iStatPP2012.reg" }
    Write-Log $logfile "$($workstation): reg imported"
}
         