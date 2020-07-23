$explorerReg = Get-ItemProperty  -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"


$explorer = New-Object -TypeName psobject 
Add-Member -InputObject $explorer -MemberType NoteProperty -Name AlwaysShowMenus -Value ""
Add-Member -InputObject $explorer -MemberType NoteProperty -Name DisplayFileSizeFolderTips -Value ""
Add-Member -InputObject $explorer -MemberType NoteProperty -Name HiddenFilesAndFolders -Value ""
Add-Member -InputObject $explorer -MemberType NoteProperty -Name HideEmptyDrives -Value ""
Add-Member -InputObject $explorer -MemberType NoteProperty -Name HideFileExtensions -Value ""
Add-Member -InputObject $explorer -MemberType NoteProperty -Name AlwaysShowIcons -Value ""
Add-Member -InputObject $explorer -MemberType NoteProperty -Name DisplayFullPathInTitleBar -Value ""

if (Get-Member -inputobject $explorerReg -name "AlwaysShowMenus" -Membertype Properties) 
{
    if ($explorerReg.AlwaysShowMenus -eq 1) 
    {
        $explorer.AlwaysShowMenus = "Checked"
    }
    else
    {
            $explorer.AlwaysShowMenus = "Unchecked"
    }
}
else
{
        $explorer.AlwaysShowMenus = "Unchecked"
}

if (Get-Member -inputobject $explorerReg -name "FolderContentsInfoTip" -Membertype Properties) 
{
    if ($explorerReg.FolderContentsInfoTip -eq 1) 
    {
        $explorer.DisplayFileSizeFolderTips = "Checked"
    }
    else
    {
            $explorer.DisplayFileSizeFolderTips = "Unchecked"
    }
}
else
{
        $explorer.DisplayFileSizeFolderTips = "Checked"
}



if ($explorerReg.Hidden -eq 1) 
{
    $explorer.HiddenFilesAndFolders = "Show"
}
else
{
    $explorer.HiddenFilesAndFolders = "Hide"
}


switch ($explorerReg.HideDrivesWithNoMedia)
{
    0 {$explorer.HideEmptyDrives = "Unchecked"; break;}
    1 {$explorer.HideEmptyDrives = "Checked"; break;}
    default {$explorer.HideEmptyDrives = "Unknown"; break;}
}

switch ([int] $explorerReg.HideFileExt)
{
    0 {$explorer.HideFileExtensions = "Unchecked"; break;}
    1 {$explorer.HideFileExtensions = "Checked"; break;}
    default {$explorer.HideFileExtensions = "Unknown"; break;}
}

switch ([int] $explorerReg.IconsOnly)
{
    0 {$explorer.AlwaysShowIcons = "Unchecked"; break;}
    1 {$explorer.AlwaysShowIcons = "Checked"; break;}
    default {$explorer.AlwaysShowIcons = "Unknown"; break;}
}

$cabinetState = Get-ItemProperty  -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CabinetState"

switch ([int] $cabinetState.FullPath)
{
    0 {$explorer.DisplayFullPathInTitleBar = "Unchecked"; break;}
    1 {$explorer.DisplayFullPathInTitleBar = "Checked"; break;}
    default {$explorer.DisplayFullPathInTitleBar = "Unknown"; break;}
}

return $explorer
