cls
$events = Get-EventLog -ComputerName  DGVMWEBFSTCPP01 -LogName Security  -newest 5 | Select-Object -Property *
$events