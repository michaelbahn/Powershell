$computers = Get-Content "D:\Scripts\test files\Dev-test2.txt"
$last_boot = Get-WmiObject win32_operatingsystem -ComputerName $computers |
Select csname, @{LABEL = 'LastBootUpTime'; EXPRESSION = {$_.ConverttoDateTime($_.lastbootuptime)}} | ConvertTo-HTML -Fragment

$result = @()
##ForEach ($computer in $computers) {
##$result += get-service -ComputerName $objitem -Name "*sql*" 

##$result = $result | Select-Object MachineName,Status,Name,DisplayName | ConvertTo-HTML -Fragment
##}
$BodyText = @"
<p>Check the message to be sure the services are started</p>
<p> </p>
<p>Services</br>
$result
</p>
<p> </p>
<p>Computers were last rebooted</br>
$last_boot
</p>
"@

$BodyText | Out-File D:\Temp\body.txt

$params = @{
    From = 'Kevin.Garcia@edd.ca.gov'
    To = 'kevin.garcia@edd.ca.gov'
    Subject = 'Reboot Script Results'
    Body = $BodyText
    SMTPServer = '151.143.2.2'
}

Send-MailMessage @params -BodyAsHTML
