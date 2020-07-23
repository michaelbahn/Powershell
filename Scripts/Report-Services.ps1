$computers = Get-Content "C:\Utility\SQL-Server.txt"
$last_boot = Get-WmiObject win32_operatingsystem -ComputerName $computers |
Select csname, @{LABEL = 'LastBootUpTime'; EXPRESSION = {$_.ConverttoDateTime($_.lastbootuptime)}} |
ConvertTo-HTML -Fragment

$result = @()
ForEach ($objitem in $computers) {
  $result += get-service -ComputerName $objitem -Name "*sql*" 
 }

$result = $result | Select-Object MachineName,Status,Name,DisplayName | ConvertTo-HTML -Fragment

##<p>Check the message to be sure the services are started</p>
##<p> </p>
##<p>Services</br>

$BodyText = @"

$result
</p>
<p> </p>
<p>Servers were last rebooted</br>
$last_boot
</p>
"@

$BodyText | Out-File D:\Temp\body.txt

$params = @{
    From = 'Astika.Kishore@edd.ca.gov'
    To = 'teaminf@edd.ca.gov'
    Subject = 'SQL Reboot Script Results'
    Body = $BodyText
    SMTPServer = '151.143.2.2'
}

Send-MailMessage @params -BodyAsHTML
