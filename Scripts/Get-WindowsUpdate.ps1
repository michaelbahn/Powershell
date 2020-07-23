$au = Get-ItemProperty  -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

switch ($au.AUOptions) 
{
    2 {return "Notify for download"; break}
    3 {return "Download updates only"; break}
    4 {return "Download and schedule"; break}
    default {return "Unknown"; break}
}