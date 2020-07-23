$computers = Get-Content "U:\PS\input\servers.txt"
Foreach($computer in $computers){
try {
 $Networks = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -ComputerName $Computer -ErrorAction Stop
 } catch {
  Write-Verbose "Failed to Query $Computer. Error details: $_"
  continue
}
foreach($Network in $Networks) {
  $DNSServers = $Network.DNSServerSearchOrder
  $NetworkName = $Network.Description
 If(!$DNSServers) {
  $PrimaryDNSServer = "Notset"
  $SecondaryDNSServer = "Notset"
 } elseif($DNSServers.count -eq 1) {
  $PrimaryDNSServer = $DNSServers[0]
  $SecondaryDNSServer = "Notset"
 } else {
  $PrimaryDNSServer = $DNSServers[0]
  $SecondaryDNSServer = $DNSServers[1]
}
}
$OutputObj = New-Object -Type PSObject
$OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $computer.ToUpper()
$OutputObj | Add-Member -MemberType NoteProperty -Name PrimaryDNSServers -Value $PrimaryDNSServer
$OutputObj | Add-Member -MemberType NoteProperty -Name SecondaryDNSServers -Value $SecondaryDNSServer
$OutputObj
}


