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
$configFile = join-path $settingsPath "Audit-Server.csv"  #list of servers to QA
$serverReport = join-path $logPath "Audit-Server-Report-$($now)"
$serverReportXLS = join-path $logPath "Audit-Server-Report-$($now).xlsx"
$getFolderPermissionsScript =  join-path $modulePath "Get-FolderPermissions.ps1"
$serverXLS = Join-Path $logPath "Audit-Server.xlsx"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

#get list of servers
$servers = Import-Csv -Path  $configFile

$userRightsReport  = $null
#loop thru each error check
foreach ($server in $servers) 
{   
    $acl = get-acl $server.Path | Select-Object Access
    $userRights = $acl.Access | Where-Object {$_.IdentityReference -like "EDD_Domain*"} | Sort-Object -Property IdentityReference
    $userRightsReport += $userRights | Select-Object -Property @{Name = 'Path'; Expression={$server.Path}}, @{Name = 'UserID'; Expression = {$_.IdentityReference.ToString().Substring(11)}}, FileSystemRights, IsInherited
}

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$xl.DisplayAlerts = $false

$sheet = @($userRightsReport)

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
    

 $sender = Get-Content  (join-path $settingsPath Sender.txt)
$recipients = Get-Content  (join-path $settingsPath recipients.txt)
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject "BRS Server File Permissions"  -Attachments $serverReportXLS
