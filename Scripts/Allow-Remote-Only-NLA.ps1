$rdp = Get-ItemProperty  -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"

switch ($rdp.SecurityLayer) 
{
    0 {return "Disabled"; break}
    2 {return "Enabled"; break}
    default {return "Unknown"; break}
}