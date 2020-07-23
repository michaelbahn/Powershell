function Get-DLN-Prefix ($reconFirstRow)
{
    $dlnFirst = $reconFirstRow.Substring($reconFirstRow.Length -16)
    return $dlnFirst.Substring(0, 13)
}

function Get-DLN-Count ($reconLast)
{
    [int] $dlnCount = 0
    switch ($reconLast.Substring(0,4).ToUpper())
    {
        "VERI" {
                        $startDLNindex = 33
                        $endDLNindex = $reconLast.Length
                        break
                   }
       "ALL " {
                        $startDLNindex = 4 
                        $endDLNindex = $reconLast.IndexOf("DLNs")
                        break;
                }
        "ONLY " {
                        $startDLNindex = $reconLast.IndexOf("of")  + 2
                        $endDLNindex = $reconLast.IndexOf("DLNs")
                        break
                       }
                
        default {
                        Write-Host "Unexpected count row: $($reconLast)" 
                        return 0
                    }
    }
    [int] $dlnCount = $reconLast.Substring($startDLNindex, $endDLNindex - $startDLNindex - 1)
    return $dlnCount
}
$reconDir =Get-ChildItem -Path  "\\dgvmappuipd01\dms\rbeben\recon" -File -ErrorAction SilentlyContinue

foreach ($reconFileName in $reconDir)
{
                $reconFileData = Get-Content $reconFileName.FullName
                $recon = $reconFileData-Split [Environment]::NewLine

                $lineCount = $recon.Count - 1
                $dlnSearch = Get-DLN-Prefix ($recon[0])  #DLN is in first line
                Write-Host "$($reconFileName) DLN: $($dlnSearch)"
                [int] $dlnCount = Get-DLN-Count ($recon[$lineCount])  #DLN count is in count line
                Write-Host "$($reconFileName) Count: $($dlnCount)"
                if ($dlnCount -eq 0)
                    {$dlnCount = $lineCount}
}
