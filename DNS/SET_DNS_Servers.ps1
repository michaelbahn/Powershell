$computer = Get-Content "U:\PS\input\servers.txt"
Foreach($computer in $computer){
$NICs = Get-WMIObject Win32_NetworkAdapterConfiguration -computername $computer | where{$_.IPEnabled -eq “TRUE”}

Foreach($NIC in $NICs) {
$DNSServers = “151.143.56.202",”151.143.101.7" # set dns servers here
 $NIC.SetDNSServerSearchOrder($DNSServers)
 $NIC.SetDynamicDNSRegistration(“TRUE”)
}
 }