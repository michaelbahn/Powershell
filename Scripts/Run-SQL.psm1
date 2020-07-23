function Query-DLN-Count ($sqlServer, $database, $dlnSearch)
{
    $query = “SELECT OBJECT_ID, dln FROM [$($database)].[dbo].[ATTRIBUTES] WHERE dln like '$($dlnSearch)%'"

    $dataSet = Invoke-Sqlcmd -ServerInstance $sqlServer -Query $query
    
    if ($dataSet -ne $null)
    {
        return $dataSet.Count
    }
    else
    {
        return $null
    }

}