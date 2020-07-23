$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$configPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"

$iCapturePath = "c$\iCapture\PROD51\bin\dcr"
$newFile = (join-path $configPath "Workstations\bin\dcr\DE5617Plugin.dll")
$newFileName = Split-Path $newFile -Leaf
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

#get list of workstations
$workstations = Get-Content  (join-path $configPath "workstation-add-DE5617Plugin.txt")

foreach ($workstation in $workstations) 
{   

    $workstationPath = "\\$($workstation.Trim())\$($iCapturePath)"
     if (! (Test-Path $workstationPath))
    {
        New-Item $workstationPath -ItemType Directory
        Write-Log $logfile "$($workstationPath) folder created if no error above"
    }

    #copy file
    Copy-Item -Path $newFile -Destination $workstationPath -Recurse -Force
    if (Test-Path -Path (join-path $workstationPath $newFileName)) 
    {
        Write-Log $logfile "$($workstation): $($newFileName) copied to $($workstationPath)"
    }
    else
    {
        Write-Log $logfile "Error: $($workstation): $($newFileName) not in  $($workstationPath)"
    }

}
         