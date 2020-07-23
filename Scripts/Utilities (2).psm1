function Initialize-Log ($logPath, $logFilePrefix)
{
    If (!(test-path $logPath))   #create log folder if it doesn't exist
    {
          New-Item -ItemType Directory -Force -Path $logPath
    }

    #create new log file
    $now = get-date -format yyyy-MM-dd-HH-mm
    $logFilePath = "$($logFilePrefix)-$($now).log"
    $logFilePath = Join-Path $logPath  $logFilePath 
    
	If ((test-path $logFilePath))   #create log file if it doesn't exist
    {
		$logFile = Get-Item  $logFilePath 
    }
    else
    {
		$logFile = New-Item $logFilePath -ItemType File
    }

    return $logFile
}

#function to write to log file
function Write-Log($logFileObject, [string] $logText)
{
        try
        {
            $logTextwithDate = "$(get-date)`t$($logText)"
            $logTextwithDate | Out-File $logFileObject.FullName -Append 
            Write-Host $logTextwithDate
        }
        catch
        {
            Write-Host "Write-Log error: $($logFileObject.FullName)"
        }
}

#function to test settings files
function Test-Files ($logFile, [string[]]$settingFiles)
{
    $returnValue = $true
    foreach ($settingFile in $settingFiles)
    {
        If (!(test-path $settingFile))  
        {
              Write-Log($logFile, "$($settingFile) not found.")
              $returnValue = $false
              break
        }        
    }

    return $returnValue

}


#converted number to enabled/disabled
function Is-Enabled ($number)
{
        If ($number -eq 1)   
        {
          return "Enabled"
        }        
        else
        {
          return "Disabled"
        }        


}

function Get-DLN-Prefix ($recon)
{
    $reconRow = $recon[0]

    #check for Linking text 
    $indexLinking = $reconRow.IndexOf("Linking")
    if ($indexLinking -lt 0)
    {
        $indexDLN = $reconRow.IndexOf("DLN")
        $dln = $reconRow.Substring($indexDLN+4).Trim()
        $dlnSearch = $dln.SubString(0, $dln.Length - 3)    #truncate right 3 characters for search
    }
    else    #for Linking Batch, get DLN from last row
    {
        $reconRow = $recon[$recon.Count - 1]     #last row
        $indexStart = $reconRow.LastIndexOf("\") + 1
        $indexEnd = $reconRow.IndexOf(" ", $indexStart) 
        $dln = $reconRow.Substring($indexStart, $indexEnd - $indexStart)    
        $dlnSearch = $dln -replace "\D+"          #remove alpha characters
        $year = (get-date).Year.ToString()
        $dlnSearch = $dlnSearch.Replace($year, $year.Substring(2))
    }

    return $dlnSearch 
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

function Is-Numeric ($Value) {
    return $Value -match "^[\d\.]+$"
}

function AddTabToLines ( [string[]]$stream) {
    $newStream = ""
    for($i=0;$i -lt $stream.count;  $i++)
    {
        $newStream += "$($stream[$i])`t`r`n"
    }
        
    return $newStream
}

#returns true if at least two parameters are greater than zero
function AtLeastTwo($count1, $count2, $count3) 
{
    $a = ($count1 -gt 0)
    $b = ($count2 -gt 0)
    $c = ($count3 -gt 0)

    if (($a -and $b) -or ($b -and $c) -or ($a -and $c)) 
    {
        return $true;
    }
    else
    {
        return $false;
    }
}

function  Update-HTML-File ($fileHTMLtemplate, $token, $newText)
{
    $content = Get-Content $fileHTMLtemplate
    $newContent = $content.Replace($token, $newText) | out-String
    return $newContent 
}

function Move-OrCopyItem ($sourceFolder, $destinationFolder)
{
    if (Test-Path $destinationFolder)
    {
        Copy-Item $sourceFolder $destinationFolder -Force
        Remove-Item $sourceFolder -Recurse -Force 
    }
    else
    {
        Move-Item $sourceFolder $destinationFolder -Force
    }
}
