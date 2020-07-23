$desktop = Get-ItemProperty  -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop"


if (Get-Member -inputobject $desktop -name "ScreenSaveActive" -Membertype Properties) 
{
	if ($desktop.ScreenSaveActive -eq 1) 
	{
		return $desktop.ScreenSaveTimeOut
	}
	else
	{
		return 0
	}
}
else
{
	return $null
}
