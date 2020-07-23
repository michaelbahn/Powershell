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
$configFile = join-path $settingsPath "Server-List.txt"  #list of servers to QA
$serverReport = join-path $logPath "Audit-Server-Report-$($now)"
$serverReportXLS = join-path $logPath "Audit-Server-Report-$($now).xlsx"
$getFolderPermissionsScript =  join-path $modulePath "Get-FolderPermissions.ps1"
$serverXLS = Join-Path $logPath "Audit-Server.xlsx"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

#get list of servers
$servers = Get-Content $configFile

$userRightsReport  = $null
#loop thru each error check
$userRightsAll = @() 
foreach ($server in $servers) 
{   
    $session = New-PSSession $server
    $shares = net view $server /all | select -Skip 7 | ?{$_ -match 'disk*'} | %{$_ -match '^(.+?)\s+Disk*'|out-null;$matches[1]}

    foreach ($share in $shares) 
    {
        $shareAccess = Invoke-Command -Session $session -ScriptBlock {param($p1) Get-SmbShareAccess -Name $p1} -ArgumentList $share
        $userRights = $shareAccess | Select-Object -Property @{Name = 'Server'; Expression={$server}}, @{Name = 'Share'; Expression={$share}}, @{Name = 'Account'; Expression = {$_.AccountName.ToString()}} 
        $userRightsAll += $userRights | Where-Object {!(($_.Account -like "BUILTIN\*") -or ($_.Account -like "NT AUTHORITY\*")) }

        #$userRights += $shareAccess | Select-Object -Property @{Name = 'Server'; Expression={$server}}, @{Name = 'UserID'; Expression = {$_.AccountName.ToString().Substring(11)}} | Where-Object {$_.AccountName -like "EDD_Domain*"} 
        
    }
    Remove-PSSession -Session $session
    Clear-Variable $session
}

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$xl.DisplayAlerts = $false

$sheet = @($userRightsAll)

$serverReportCSV = "$($serverReport)-$($index).csv"

    if (Test-Path ($serverReportCSV))
    {
        Remove-Item $serverReportCSV
    }

$sheet | Export-Csv -path $serverReportCSV -NoTypeInformation
    
$wbMain = $xl.Workbooks.Open($serverReportCSV) 
    $wbMain.SaveAs($serverReportXLS, 51)
    $wbMain.Save
    $wbMain.Close($false)
    $xl.Quit()
