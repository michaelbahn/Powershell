

$computer= $env:computername

get-Service -ComputerName $computer | where {$_.name -like "Unisys*"} | Start-Service

cls