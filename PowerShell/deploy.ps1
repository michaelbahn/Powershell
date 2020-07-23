$targetPflServers = @("\\DGVMDRBEPP01", "\\DGVMPRBEPP01")
$targetDiaServers = @("\\DGVMDRBEPP01", "\\DGVMPRBEPP01")
$sourcePflServer = "\\DGVMDRBEPP01"
$sourceDiaServer = "\\DGVMDRBEPP01"
$sysFolder = "d$\Program Files (x86)\UeWI\SYS" 

#backup
$targetServers = $targetPflServers + $targetDiaServers
foreach ($targetServer in $targetServers)
{
    $backupFolder = Join-Path $targetServer "d$\backup"
    if (! (Test-Path $backupFolder))
    {
        New-Item $backupFolder -ItemType Directory
    }

    $targetSysFolder = join-path $targetServer "d$\Program Files (x86)\UeWI\SYS" 
    Copy-Item "$($targetSysFolder)\DIWHITE.frm" $backupFolder  -Force
    Copy-Item "$($targetSysFolder)\DCFORMSD.FRM" $backupFolder  -Force
    Copy-Item "$($targetSysFolder)\PFLWHITE.FRM" $backupFolder  -Force
    Copy-Item "$($targetSysFolder)\rioidx.dat" $backupFolder  -Force
}

$sourceSysFolder = join-path $sourcePflServer "d$\Program Files (x86)\UeWI\SYS" 
foreach ($targetPflServer in $targetPflServers)
{
    $targetSysFolder = join-path $targetPflServer "d$\Program Files (x86)\UeWI\SYS" 
    Copy-Item "$($sourceSysFolder)\DIWHITE.frm" $backupFolder  -Force
    Copy-Item "$($sourceSysFolder)\DCFORMSD.FRM" $backupFolder  -Force
    Copy-Item "$($sourceSysFolder)\PFLWHITE.FRM" $backupFolder  -Force
    Copy-Item "$($sourceSysFolder)\rioidx.dat" $backupFolder  -Force
}
    


Copy-Item "$($sourceSysFolder)\DIWHITE.frm" $targetSysFolder -Force
Copy-Item "$($sourceSysFolder)\DCFORMSD.FRM" $targetSysFolder -Force
Copy-Item "$($sourceSysFolder)\PFLWHITE.FRM" $targetSysFolder  -Force
Copy-Item "$($sourceSysFolder)\rioidx.dat" $targetSysFolder  -Force