
$servers = gc  'D:\Powershell-Production-scripts\Scripts\Server-reboot.txt'

$report = foreach ($server in $servers) {
Get-CimInstance -ClassName win32_operatingsystem  -Property *  -ComputerName $server | select csname, lastbootuptime
}

$report | ConvertTo-Html  | Out-File D:\Script\test.htm 

$MailBody = Get-Content D:\Script\test.htm  -Raw

Send-MailMessage -SmtpServer '151.143.2.2' -From 'Astika.Kishore@edd.ca.gov' -To 'Astika.Kishore@edd.ca.gov' -Subject 'Reboot Script Results' -Body $MailBody -BodyAsHtml


