cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$activeDirectory = @()
$usernames = @("MCave-22", "MCave", "MBahn", "MBahn-22", "DWomack", "DWomack-22", "RDuda", "RDuda-22")
foreach ($username in $usernames) 
{
    #$activeDirectory += Get-ADPrincipalGroupMembership $username | select -Property name, @{Name = "UserID";Expression = {$($username)}} | sort name
    $activeDirectory += Get-ADPrincipalGroupMembership $username | select -Property @{Name = "ADGroup";Expression = {"$($_.name)"}}, @{Name = "UserID";Expression = {$($username)}} | sort name
}
Export-CSV -InputObject $activeDirectory -Path ".\ad.csv"