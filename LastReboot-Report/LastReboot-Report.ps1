cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
    $modulePath = "..\Scripts"
    $logPath = "..\Logs"    #initialize log file 
    Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
    $logFile = Initialize-Log $logPath $serviceAction

$servers = gc  Server-List-All.txt
$sender = gc  sender.txt
$recipients = gc recipients.txt

IF ([string]::IsNullOrWhitespace($servers))
{
    Write-Host "Invalid server list " 
    return
} 

IF ([string]::IsNullOrWhitespace($sender))
{
    Write-Host "Invalid sender list "
    return
} 

IF ([string]::IsNullOrWhitespace($recipients))
{
    Write-Host "Invalid receipient list"
    return
} 
 
$report = foreach ($server in $servers) 
{
    Write-Log $logFile $server
    try 
    {
        Get-CimInstance -ClassName win32_operatingsystem  -Property *  -ComputerName $server | select csname , lastbootuptime | Add-Member -MemberType AliasProperty -Name "Last Restart Time" -Value lastbootuptime  -PassThru | Add-Member -MemberType AliasProperty -Name Server -Value csname -PassThru
    }
    catch   
    {
        Write-Log $logFile $Error[0] 

    }
}

$reportSorted = $report | Sort-Object lastbootuptime
$reportSorted | ConvertTo-Html  -Title "Last Reboot Report" -Property Server , "Last Restart Time" | Out-File "LastRebootReport.htm" 

$MailBody = Get-Content "LastRebootReport.htm"  -Raw

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject 'Last Reboot Report' -Body $MailBody -BodyAsHtml