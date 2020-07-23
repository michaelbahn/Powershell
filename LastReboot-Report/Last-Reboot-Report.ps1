cls
#Initialize settings
$now = get-date -format yyyy-MM-dd-HH-mm
$today = get-date -Format  d
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = Resolve-Path "..\Scripts"
$logPath = Resolve-Path "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath "$($title)_$($serverListFile)"

#get list of servers
$servers = Get-Content $serverListFile
$sender = "teamops@edd.ca.gov"
$recipients = Get-Content recipients.txt

$report = @(0)

foreach ($server in $servers) 
{
    Write-Log $logFile $server
    $rebootTime = Get-EventLog -LogName System -Newest 30000 | Where-Object {$_.EventID -eq 6005} | Select-Object -First 1 Time | Add-Member -MemberType AliasProperty -Name "Server" -Value $server -PassThru 
    #Get-EventLog -LogName System -Newest 30000 | Where-Object {$_.EventID -in (6005,6006,6008,6009,1074,1076)}

    if ($rebootTime.Length -eq 0)
    {
        Write-Log $logFile $Error[0] 
    }
    else 
    {
        report += $rebootTime
    }

}

$reportSorted = $report | Sort-Object Time
$reportSorted | ConvertTo-Html  -Title "Last Reboot Report" -Property "Last Restart Time", Server | Out-File "LastRebootReport.htm" 

$MailBody = Get-Content "LastRebootReport.htm"  -Raw

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject 'Last Reboot Report' -Body $MailBody -BodyAsHtml