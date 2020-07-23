$acl = get-acl \\DGVMAPPTAXPD01\H$\Shares\DMS\Auditlog  | Select-Object Access
$userRights = $acl.Access | Where-Object {$_.IdentityReference -like "EDD_Domain*"}
$userRightsReport = $userRights | Select-Object -Property FileSystemRights, IsInherited, @{Name = 'UserID'; Expression = {$_.IdentityReference.ToString().Substring(11)}}
return $userRightsReport