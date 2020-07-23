function Select-Counts ($sqlServer,$days)
{
    $rows = @()
    foreach ($day in $days)
    {
        $dateFormattedForTableName =  $day.ToString("yyyyMMdd")
        $dateFormatForReport =  $day.ToString("M/d/yy")
        $rows += @([PSCustomObject]@{Query ="SELECT count(*) as cnt FROM [dbo].[z_iddates_$($dateFormattedForTableName)] ";  Caption = "$($dateFormatForReport) Table Count"; RecordCount= 0})
    }

    $rows += @([PSCustomObject]@{Query = "SELECT count(*) as cnt FROM [dbo].[Report_Data] ";  Caption = "Report Data Count"; RecordCount= 0} )
    $rows += @([PSCustomObject]@{Query = "SELECT count(*) as cnt FROM [dbo].[Report_Data]  Where NEW_WORKSET = '-~-' or WORKSET = '-~-'" ;  Caption = "Workset: -~-"; RecordCount= 0} )
    $useDB = "use Report; "
    foreach ($row in $rows)
    {
        $query = $useDB
        $query += $row.Query
        $result = Invoke-Sqlcmd -ServerInstance $sqlServer -Query $query
        $row.RecordCount = $result.cnt -as [int]
    }
    
    return $rows
}
