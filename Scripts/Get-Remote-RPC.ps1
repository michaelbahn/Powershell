$ts = Get-ItemProperty  -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server"

switch ($ts.AllowRemoteRPC) 
{
    0 {return "Disabled"; break}
    1 {return "Enabled"; break}
    default {return "Unknown"; break}
}