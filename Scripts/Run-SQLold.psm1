function Run-SQL ($serverInstance, $dlnSearch)
{
    $slashPos = $parm1.IndexOf("\")
    $dataSource = $parm1.Substring(1, $slashPos -1)
    $database = $parm1.Substring($slashPos +1)
    $connectionString = “Server=$dataSource;Database=$database;Integrated Security=True;”

    $query = “SELECT OBJECT_ID, dln FROM [INFOIDX].[dbo].[ATTRIBUTES] WHERE dln like '$($dlnSearch)%'"

    $remoteScript = "$dataSet = Invoke-Sqlcmd -ServerInstance $serverInstance -Query $query"
    Invoke-Command -ScriptBlock $remoteScript -ComputerName $dataSource -ArgumentList $dataSet, $serverInstance, $query        
    return $dataSet.Count
}