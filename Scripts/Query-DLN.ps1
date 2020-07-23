param  ([string] $dln = $(throw "DLN is required"), [string] $sqlInstance = $(throw "SQL Instance is required"))
$query = "SELECT dln FROM [BINFO_IDX].[dbo].[ATTRIBUTES] WHERE dln like '$($dln)%'"
$resultSet =  Invoke-Sqlcmd -Query $query -ServerInstance $sqlInstance 
return $resultSet