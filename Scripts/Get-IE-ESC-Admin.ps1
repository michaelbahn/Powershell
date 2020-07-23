$escAdmin = Get-ItemProperty  -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"

switch ($escAdmin.IsInstalled) 
{
    0 {return "Disabled"; break}
    1 {return "Enabled"; break}
    default {return "Unknown"; break}
}