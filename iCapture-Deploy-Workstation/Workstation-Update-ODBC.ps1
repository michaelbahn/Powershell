$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$modulePath = "..\Scripts"

$startMenuPath = "c$\ProgramData\Microsoft\Windows\Start Menu\Programs"
$deployPath = "\\dgvmopspd02\Deploy\iCapture"

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

#get list of workstations to reboot
$workstations = Get-Content  ("workstation-update.txt")

foreach ($workstation in $workstations) 
{   
    #copy windows shortcust
    $workstationStartMenuPath = "\\$($workstation.Trim())\$($startMenuPath)"
    $result = Copy-Item -Path "$($deployPath)\iCapture*" -Destination $workstationStartMenuPath -Recurse -Force
    Write-Log $logfile "$($workstation): shortcuts added to $($startMenuPath)"

    #Open HKLM registry hive on target workstation
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $workstation)

    #Add Impression registry on target workstation
    $wow64key = $reg.OpenSubKey("SOFTWARE\Wow6432Node", $true)
    $wow64key.CreateSubKey("Impression Technology")
    $impressionKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology", $true)
    $impressionKey.CreateSubKey("Configuration")
    $configurationKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\Configuration", $true)
    $configurationKey.CreateSubKey("DEFAULT")
    $defaultConfigurationKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\Configuration\DEFAULT", $true)
    $defaultConfigurationKey.SetValue("ConfHost","DGVMLVDIMGPD01,DGVMPFLIMGPD01")
    $defaultConfigurationKey.SetValue("ConfRegSubKey","SOFTWARE\\Impression Technology\\PROD50")
    $configurationKey.CreateSubKey("PRE_PROD50")
    $defaultConfigurationKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\Configuration\PRE_PROD50", $true)
    $defaultConfigurationKey.SetValue("ConfHost","DGVMICAPIMGPP01")
    $defaultConfigurationKey.SetValue("ConfRegSubKey","SOFTWARE\\Impression Technology\\PRE_PROD50")
    $configurationKey.CreateSubKey("PRE_PROD51")
    $defaultConfigurationKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\Configuration\PRE_PROD51", $true)
    $defaultConfigurationKey.SetValue("ConfHost","DGVMIMGIDXPP01")
    $defaultConfigurationKey.SetValue("ConfRegSubKey","SOFTWARE\\Impression Technology\\PRE_PROD50")

    #Add ODBC registry on target workstation
    #verify at "C:\windows\SysWOW64\odbcad32.exe"
    $odbcKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\ODBC\ODBC.INI", $true)

    $odbcKey.CreateSubKey("iStatPP2012")
    $subKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\ODBC\ODBC.INI\iStatPP2012", $true)
    $subKey.SetValue("Driver", "C:\\Windows\\System32\\sqlsrv32.dll")
    $subKey.SetValue("Description", "iStatistics")
    $subKey.SetValue("Server", "DGC1VDBICAPPP01\DGC1VDBICAPPP01")
    $subKey.SetValue("Database", "iStatistics")
    $subKey.SetValue("LastUser", "")
    $subKey.SetValue("Trusted_Connection", "Yes")

    $odbcKey.CreateSubKey("ShadowPP")
    $subKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\ODBC\ODBC.INI\ShadowPP", $true)
    $subKey.SetValue("Driver", "C:\\Windows\\System32\\sqlsrv32.dll")
    $subKey.SetValue("Description", "ShadowPP")
    $subKey.SetValue("Server", "DGC1VDBICAPPP01\DGC1VDBICAPPP01")
    $subKey.SetValue("Database", "iCapture_Shadow")
    $subKey.SetValue("LastUser", "")
    $subKey.SetValue("Trusted_Connection", "Yes")
    $subKey.SetValue("", "")

    $odbcKey.CreateSubKey("ODBC Data Sources")
    $subKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\ODBC\ODBC.INI\ODBC Data Sources", $true)
    $subKey.SetValue("iStatPP2012", "SQL Server")
    $subKey.SetValue("ShadowPP", "SQL Server")

    Write-Log $logfile "$($workstation): registry updated"
    

}
         