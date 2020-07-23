function Initialize-Log ($logPath, $logFilePrefix)
{
    If (!(test-path $logPath))   #create log folder if it doesn't exist
    {
          New-Item -ItemType Directory -Force -Path $logPath
    }

    #create new log file
    $now = get-date -format yyyy-MM-dd-HH-mm
    $logFile = "$($logFilePrefix)-$($now).log"
    $logFile = Join-Path $logPath $logFile
    New-Item $logFile -ItemType File
    return $logFile
}

#function to write to log file
function Write-Log($logFileObject, [string] $logText)
{
        $logText | Out-File $logFileObject.FullName -Append 
        Write-Host $logText
}

#function to test settings files
function Test-Files ($logFile, [string[]]$settingFiles)
{
    $returnValue = $true
    foreach ($settingFile in $settingFiles)
    {
        If (!(test-path $settingFile))   #create log folder if it doesn't exist
        {
              $returnValue = $false
              break
        }        
    }

    return $returnValue

}
