cls
$scriptPath = Get-Location
$settingsPath = "..\scripts"

$outputFile = Join-Path $settingsPath test.htm 

$servers = gc  (join-path $settingsPath Server-reboot.txt)
$sender = gc  (join-path $settingsPath Sender.txt)
$recipients = gc  (join-path $settingsPath Recipients.txt)

IF ([string]::IsNullOrWhitespace($servers))
{
    Write-Host "Invalid server list at " $settingsPath\Server-reboot.txt
    return
} 

IF ([string]::IsNullOrWhitespace($sender))
{
    Write-Host "Invalid sender list at " $settingsPath\Sender.txt
    return
} 

IF ([string]::IsNullOrWhitespace($recipients))
{
    Write-Host "Invalid receipient list at " $settingsPath\Recipients.txt
    return
} 
 
$report = foreach ($server in $servers) {
#Get-CimInstance -ClassName win32_operatingsystem  -Property *  -ComputerName $server | select csname , lastbootuptime 
Get-CimInstance -ClassName win32_operatingsystem  -Property *  -ComputerName $server | select csname , lastbootuptime | Add-Member -MemberType AliasProperty -Name "Last Restart Time" -Value lastbootuptime  -PassThru | Add-Member -MemberType AliasProperty -Name Server -Value csname -PassThru                                                                                                                                                                                                                                                                                             
}

$reportSorted = $report | Sort-Object lastbootuptime
$reportSorted | ConvertTo-Html  -Title "Last Reboot Report" -Property Server , "Last Restart Time" | Out-File $outputFile

$MailBody = Get-Content $outputFile  -Raw

Send-MailMessage -SmtpServer '151.143.2.2' -From $sender -To $recipients -Subject 'Last Reboot Report' -Body $MailBody -BodyAsHtml