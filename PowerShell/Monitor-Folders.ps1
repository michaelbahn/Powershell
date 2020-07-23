cls
$title = "Monitor-Folders"

#Initialize settings
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$outputFile = Join-Path $settingsPath test.txt
$now = get-date


#log file settings
Import-Module (Join-Path $modulePath Write-Log.psm1) -Force
$logFile = Initialize-Log $logPath $title

#email settings
$sender = Get-Content  (join-path $settingsPath Sender.txt)
$recipients = Get-Content (join-path $settingsPath Recipients.txt)
If ([string]::IsNullOrWhitespace($sender))
{
    Write-Log $logFile "Invalid sender list at $($settingsPath)\Sender.txt"
    return
} 
If ([string]::IsNullOrWhitespace($recipients))
{
    Write-Log $logFile  "Invalid receipient list at $($settingsPath)\Recipients.txt" 
    return
} 

#Get monitor settings
$monitorSettings = Import-Csv -Path  (join-path $settingsPath Monitor-Errors.csv) 

#loop thru each check
foreach ($monitor in $monitorSettings) 
{   
     #ItemType specifies files or folders
    switch ($monitor.ItemType.Trim())  
    {
        "Directory"
            {$gciItemType = "-Directory"; break}
        "File"
            {$gciItemType = "-File"; break}
        default
            {$gciItemType = ""; break}
    }
        
    #if date/time thresholds are specified, build the Where clause
    $dateTimeThresholdHigh = $null
    $dateTimeThresholdLow = $null
    
    if ($monitor.dayThresholdLow -ne 0)
    {
        $dateTimeThresholdLow = $now.AddDays(0 - $monitor.dayThresholdLow)
        Write-Log $logFile "$($monitor.Title) low threshold: $($dateTimeThresholdLow)"
    }

    if ($monitor.minuteThresholdLow	-ne 0)
    {
        $dateTimeThresholdLow = $now.AddMinutes(0 - $monitor.minuteThresholdLow)
        Write-Log $logFile "$($monitor.Title) low threshold: $($dateTimeThresholdLow)"
    }

    if ($monitor.dayThresholdHigh	-ne 0)
    {
        $dateTimeThresholdHigh= $now.AddDays(0 - $monitor.dayThresholdHigh)
        Write-Log $logFile "$($monitor.Title) high threshold: $($dateTimeThresholdHigh)"
    }

    if ($monitor.minuteThresholdHigh	-ne 0)
    {
        $dateTimeThresholdHigh= $now.AddMinutes(0 - $monitor.minuteThresholdHigh)
        Write-Log $logFile "$($monitor.Title) high threshold: $($dateTimeThresholdHigh)"
    }

    $gciWhere = ""
    if (($dateTimeThresholdHigh -ne $null) -or  ($dateTimeThresholdLow -ne $null))
    {
        $gciWhere = "| Where-Object {"

        if ($dateTimeThresholdLow -ne $null) 
        {
            $gciWhere += "`$_.LastWriteTime -ge `$dateTimeThresholdLow "
        }

        if ($dateTimeThresholdHigh -ne $null)
        {
            if ($dateTimeThresholdLow -ne $null)   #if both thresholds set add an AND
            {
                $gciWhere += "  -and "
            }

            $gciWhere += "`$_.LastWriteTime -le `$dateTimeThresholdHigh "
        }
        #end of where
        $gciWhere += "} "
    }
    else
    {
        $gciWhere = ""
    }

    #assemble Get-ChildItem command to execute
    #example: Get-ChildItem -Path $monitor.Folder -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $dateTimeThresholdLow  -and $_.LastWriteTime -le $dateTimeThresholdHigh } | Select-Object FullName, LastWriteTime
    $monitorFolders  = $null
    $gciCommand = "`$monitorFolders  = Get-ChildItem -Path `$monitor.Folder $($gciItemType) -Force -ErrorAction SilentlyContinue $($gciWhere)  | Select-Object FullName, LastWriteTime"
    Invoke-Expression $gciCommand 

    if ($monitorFolders  -ne $null)
    {
        Write-Log $logFile "$($monitor.Title) found at $($monitor.Folder)"
        $monitorFolders | ConvertTo-Html  -Title $title -Property FullName , LastWriteTime | Out-File $outputFile
        $mailBody = Get-Content $outputFile  -Raw
        Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $sender -To $recipients -Subject $monitor.Title -Body $mailBody -BodyAsHtml
    }
    else
    {
        Write-Log $logFile "$($monitor.Title) check passed."
    }
}