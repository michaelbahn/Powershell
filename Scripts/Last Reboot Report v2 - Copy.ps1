$scriptPath = 'D:\Powershell-Production-scripts\Scripts'
$outputFile = Join-Path $scriptPath test.htm 
$servers = gc  (join-path $scriptPath Server-reboot.txt)

$report = foreach ($server in $servers) {
Get-CimInstance -ClassName win32_operatingsystem  -Property *  -ComputerName $server | select csname , lastbootuptime 
}

$report | ConvertTo-Html  | Out-File $outputFile

$MailBody = Get-Content $outputFile  -Raw
$recipients = @('Michael.Bahn@edd.ca.gov', 'Michael.Cave@edd.ca.gov')
Send-MailMessage -SmtpServer '151.143.2.2' -From 'Michael.Bahn@edd.ca.gov' -To $recipients -Subject 'Last Reboot Report' -Body $MailBody -BodyAsHtml