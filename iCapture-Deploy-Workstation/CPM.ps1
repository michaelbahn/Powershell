$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"

$iCaptureTestPath = "c$\iCapture\PROD50\bin"
$newFiles = Get-ChildItem "C:\iCapture\PROD50\bin" -File -Recurse 

#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
$logError = Initialize-Log $logPath "$($title)_error"
$logSuccess = Initialize-Log $logPath "$($title)_success"

#get list of workstations
$workstations = Get-Content  (".\cpm_failed.txt")
#$workstations = Get-Content  (".\workstation-icapture-prod.txt")

foreach ($workstation in $workstations) 
{   

    $workstationPath = "\\$($workstation.Trim())\$($iCaptureTestPath)"

    if ((Test-Path $workstationPath))
    {
        foreach ($file in $newFiles) 
        {
            $iCapturePath = $file.DirectoryName.Replace(":", "$")
            $workstationPath = "\\$($workstation.Trim())\$($iCapturePath)"
            #copy file
            Copy-Item -Path $file.FullName -Destination $workstationPath -Force
            if (Test-Path -Path (join-path $workstationPath $newFileName)) 
            {
                Write-Log $logFile "$($workstation): $($file.Name) copied to $($workstationPath)"
            }
            else
            {
                Write-Log $logFile "$($workstation): $($file.Name) not in  $($workstationPath)"
                Write-Log $logError "$($workstation)"
            }
        }
        Write-Log $logSuccess "$($workstation)"

    }
    else
    {
        #New-Item $workstationPath -ItemType Directory
        Write-Log $logFile "can't reach $($workstation)"
        Write-Log $logError "$($workstation)"
    }
}
         