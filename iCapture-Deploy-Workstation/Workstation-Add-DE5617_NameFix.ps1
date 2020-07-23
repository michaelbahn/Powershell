$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

$newFilePath = "\\dgvmimgidxpd01\d$\Deployment\DE5617Changes04192019\Workstations\bin"
$newFileName = "ITISL.dll"
$iCapturePath = "c$\iCapture\PRE_PROD51\bin"
$newFile = join-path $newFilePath $newFileName

#get list of workstations
$workstations = Get-Content  (join-path $configPath "workstation-icapture-trainnig.txt")

foreach ($workstation in $workstations) 
{   

    $workstationPath = "\\$($workstation.Trim())\$($iCapturePath)"
     if (! (Test-Path $workstationPath))
    {
        Write-Log $logfile "$($workstationPath) folder does not exist"
    }

    #copy file
    $localFilePath = join-path $workstationPath $newFileName
    try{
        Rename-Item -Path $localFilePath -NewName "$($localFilePath)_04192019" -Force
        Copy-Item -Path $newFile -Destination $workstationPath  -Force
        Write-Log $logfile "$($workstation): $($newFileName) copied to to $($workstationPath)"
    }
    catch
     {
        Write-Log $logfile "Error: $($workstation): $($newFileName) not in  $($workstationPath)"
    }

}
         