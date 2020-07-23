function Get-SearchPath ($pathSeftToXprc, $mainframeFile)
{
    if ($mainframeFile.IndexOf("WINONVAL") -ge 0)
    {
        $searchPath = Join-Path $pathSeftToXprc "BWSLoadErrors"
    }
    elseif ($mainframeFile.IndexOf("DEFWILD") -ge 0)
    {
        $searchPath = Join-Path $pathSeftToXprc "DefAddIn"
    }
    elseif ($mainframeFile.IndexOf("MCERRORS") -ge 0)
    {
        $searchPath = Join-Path $pathSeftToXprc "MassChangeError"
    }
    elseif ($mainframeFile.IndexOf("MCTRANS") -ge 0)
    {
        $searchPath = Join-Path $pathSeftToXprc "MassChangeTran"
    }
    elseif ($mainframeFile.IndexOf("REJECT") -ge 0)
    {
        $searchPath = Join-Path $pathSeftToXprc "WIReject"
    }
    else
    {
        $searchPath = $pathSeftToXprc 
    }

    return $searchPath
}