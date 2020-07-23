cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath

$servers = gc  servers.txt
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
 
$report = foreach ($server in $servers) {
Get-CimInstance -ClassName win32_operatingsystem  -Property *  -ComputerName $server | select csname , lastbootuptime | Add-Member -MemberType AliasProperty -Name "Last Restart Time" -Value lastbootuptime  -PassThru | Add-Member -MemberType AliasProperty -Name Server -Value csname -PassThru                                                                                                                                                                                                                                                                                             
}

$reportSorted = $report | Sort-Object lastbootuptime
$reportSorted | ConvertTo-Html  -Title "Last Reboot Report" -Property Server , "Last Restart Time" | Out-File "LastRebootReport.htm" 

$MailBody = Get-Content "LastRebootReport.htm"  -Raw

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject 'Last Reboot Report' -Body $MailBody -BodyAsHtml