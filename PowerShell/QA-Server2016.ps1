cls
#Initialize settings
$now = get-date -format yyyy-MM-dd-HH-mm
$today = get-date -Format  d
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = Resolve-Path "..\Config"
$modulePath = Resolve-Path "..\Scripts"
$logPath = Resolve-Path "..\Logs"
$configFile = join-path $settingsPath "QA-Servers.csv"  #list of servers to QA
$serverReport = join-path $logPath "Server-QA-Report-$($now)"
$serverReportXLS = join-path $logPath "Server-QA-Report-$($now).xlsx"
$serverLocalAdminsCSV = join-path $logPath "Server-Local-Admins.csv"
$serverLocalUsersCSV = join-path $logPath "Server-Local-Users.csv"
$getLocalUsersScript = join-path $modulePath "Get-LocalUsers.ps1"
$getNetIPConfigurationScript = join-path $modulePath "Get-NetIPConfiguration.ps1" 
$getLocalAdministratorsScript = join-path $modulePath "Get-LocalAdministrators.ps1"  
$getSNMPStatusScript = join-path $modulePath "Get-Service-SNMP.ps1"
$getSNMPStartModeScript = join-path $modulePath "Get-SNMP-StartMode.ps1"
$getSNMPRFC1156AgentScript = join-path $modulePath "Get-SNMP-RFC1156Agent.ps1"
$getSNMPTrapDestionationScript = join-path $modulePath "Get-SNMP-Trap-Destionation.ps1"
#Requested registry access is not allowed. --> $getSNMPValidCommunitiesScript =  join-path $modulePath "Get-SNMP-ValidCommunities.ps1"  
$allowRemoteRPCScript = join-path $modulePath "Get-Remote-RPC.ps1"
$AllowRemoteNLAScript = join-path $modulePath "Allow-Remote-Only-NLA.ps1"
$getWindowsUpdateScript = join-path $modulePath "Get-WindowsUpdate.ps1"
$getIE_ESC_AdminScript = join-path $modulePath "Get-IE-ESC-Admin.ps1"
$getIE_ESC_UserScript = join-path $modulePath "Get-IE-ESC-User.ps1"
$getScreenSaverScript =  join-path $modulePath "Get-ScreenSaver.ps1"
$getPowerSleepScript =  join-path $modulePath "Get-PowerSleep.ps1"
$getExplorerSettingsScript =  join-path $modulePath "Get-ExplorerSerttings.ps1"
$serverXLS = Join-Path $logPath "QA-Server.xlsx"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$eventLog = $null
$serverDisk = $null
$serverLocalUsers = $null
$serverLocalAdmins = $null

$eventLogs = @(
    'Application',
    'Security',
    'System')


#get list of servers
$servers = Import-Csv -Path  $configFile

#loop thru each error check
foreach ($server in $servers) 
{   
    $serverWmi = Get-WmiObject win32_networkadapterconfiguration -filter "ipenabled = 'True'" -ComputerName $server.Name
    $serverIPV4 = Get-NetIPAddress -CimSession $server.Name -AddressFamily IPv4 | where { $_.InterfaceAlias -notmatch 'Loopback'}
    $netIP = Invoke-Command -FilePath $getNetIPConfigurationScript -ComputerName $server.Name
    $server.IPAddress = $serverIPV4.IPAddress 
    $server.Domain = $serverWmi.DNSDomain
    $server.Subnet = $serverWmi.IPSubnet[0]
    $server.DefaultGateway = $serverWmi.DefaultIPGateway.Trim('{}') 
    $server.DNS1 = $netIP.DNSServer.ServerAddresses[0]
    $server.DNS2 = $netIP.DNSServer.ServerAddresses[1]
    $server.WINSPrimaryServer = $serverWmi.WINSPrimaryServer
    $server.WINSSecondaryServer = $serverWmi.WINSSecondaryServer
    $server.LMHOSTS_Lookup = $serverWmi.WINSEnableLMHostsLookup
    $server.NetBIOSoverTCPIP = Is-Enabled(Get-WmiObject win32_networkadapterconfiguration -filter "ipenabled = 'True'" -ComputerName $server.Name | select -ExpandProperty Tcpipnetbiosoptions)
    $server.IPversion = $serverIPV4.AddressFamily
    
    $eventLog += get-eventlog -list -ComputerName $server.Name | Where-Object {$eventLogs -contains$_.log} | Select-Object -Property @{Name = 'ServerName'; Expression = {$server.Name}},MaximumKilobytes,OverFlowAction,Log
    $serverDisk += Get-WmiObject win32_logicaldisk -ComputerName $server.Name | Where-Object {$_.DriveType -eq 3} | Select-Object -Property @{Name = 'ServerName'; Expression = {$server.Name}},DeviceID,@{Name="Size-GB";Expression={[math]::round($_.size/1GB)}}
    $serverLocalUsers += Invoke-Command -FilePath $getLocalUsersScript -ComputerName $server.Name | Select-Object -Property @{Name = 'ServerName'; Expression = {$server.Name}},Name,Enabled
    #get-ciminstance win32_useraccount -ComputerName $server.Name | Where-Object  {$_.Domain -eq $server.Name}   

    #local admins
    $serverLocalAdmins += Invoke-Command -FilePath $getLocalAdministratorsScript -ComputerName $server.Name | Select-Object -Property @{Name = 'ServerName'; Expression = {$server.Name}},ObjectClass,Name,PrincipalSource 
    
    #SNMP service
    $server.SNMP_Status = Invoke-Command -FilePath $getSNMPStatusScript -ComputerName $server.Name | Select-Object -ExpandProperty Status
    $server.SNMP_StartupType = Invoke-Command -FilePath $getSNMPStartModeScript -ComputerName $server.Name | Select-Object -ExpandProperty StartMode
    $serverSNMP_Parameters = Invoke-Command -FilePath $getSNMPRFC1156AgentScript -ComputerName $server.Name
    $server.SNMP_Contact = $serverSNMP_Parameters | Select-Object -ExpandProperty sysContact
    $server.SNMP_Location = $serverSNMP_Parameters | Select-Object -ExpandProperty sysLocation
    $server.SNMP_TrapDestination = Invoke-Command -FilePath $getSNMPTrapDestionationScript -ComputerName $server.Name | Select-Object -ExpandProperty 1
    $server.RemoteDesktopEnabled = Invoke-Command -FilePath $allowRemoteRPCScript -ComputerName $server.Name 
    $server.AllowRemoteOnlyNLA = Invoke-Command -FilePath $AllowRemoteNLAScript -ComputerName $server.Name 
    $server.WindowsUpdate = Invoke-Command -FilePath $getWindowsUpdateScript -ComputerName $server.Name 
    $server.IE_ESC_Admin = Invoke-Command -FilePath $getIE_ESC_AdminScript -ComputerName $server.Name 
    $server.IE_ESC_User = Invoke-Command -FilePath $getIE_ESC_UserScript -ComputerName $server.Name 
    $server.PowerSleepTimeout = Invoke-Command -FilePath $getPowerSleepScript  -ComputerName $server.Name 
    # current user settings need a login
    #$server.ScreenSaverTimeout = Invoke-Command -FilePath $getScreenSaverScript -ComputerName $server.Name 
    #$explorerSettings = Invoke-Command -FilePath $getExplorerSettingsScript -ComputerName $server.Name 
    #$server.AlwaysShowMenus = $explorerSettings.AlwaysShowMenus 
    #$server.DisplayFileSizeFolderTips  = $explorerSettings.DisplayFileSizeFolderTips 
    #$server.HiddenFilesAndFolders = $explorerSettings.HiddenFilesAndFolders 
    #$server.HideEmptyDrives = $explorerSettings.HideEmptyDrives 
    #$server.HideFileExtensions = $explorerSettings.HideFileExtensions 
    #$server.AlwaysShowIcons = $explorerSettings.AlwaysShowIcons 
    #$server.DisplayFullPathInTitleBar = $explorerSettings.DisplayFullPathInTitleBar 

}

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$xl.DisplayAlerts = $false

$sheet = @($eventLog, $serverDisk, $serverLocalUsers, $serverLocalAdmins, $servers)
$label = @("Event Logs", "Disks", "Local Users", "Local Admins", "Servers")

$firstSheet = $true

for ($index=0; $index -lt $sheet.Count; $index++) 
{

    $serverReportCSV = "$($serverReport)-$($index).csv"

    if (Test-Path ($serverReportCSV))
    {
        Remove-Item $serverReportCSV
    }

    $sheet[$index] | Export-Csv -path $serverReportCSV -NoTypeInformation

    
    if ($index -eq 0)   #first time thru loop, create workbook
    {
        $wbMain = $xl.Workbooks.Open($serverReportCSV) 
        $wbMain.SaveAs($serverReportXLS, 51)
        $wbMain.ActiveSheet.Name = $label[$index]        
    }
    else    #add sheet to workbook
    {
        #$wsMain = $wbMain.worksheets.add()
        $wsMain = $wbMain.ActiveSheet
        $wbTemp = $xl.Workbooks.Open($serverReportCSV) 
        $wsTemp = $wbTemp.ActiveSheet
        $wsTemp.copy($wsMain)
        $wbTemp.Close
        $wbMain.ActiveSheet.Name = $label[$index]        
        $wbMain.Save
    }

}

$wbMain.Save
$wbMain.Close
$xl.Quit

#$serverLocalUsers | Export-Csv -path $serverLocalUsersCSV -NoTypeInformation
#$serverLocalAdmins | Export-Csv -path $serverLocalAdminsCSV -NoTypeInformation


