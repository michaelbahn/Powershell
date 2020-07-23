cls
#Initialize settings
$now = get-date
$today = get-date -Format  d
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$monitorErrorsHTML = Join-Path $logPath monitor-errors.html   #temp file used for email
$monitorSettingsPath = join-path $settingsPath "$($title).csv"  #config file with files/folders to monitor
$emailSenderPath = join-path $settingsPath emailMonitorSender.txt
$emailRecipientsPath  = join-path $settingsPath emailMonitorRecipients.txt

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-Log $logPath $title

#verify settings files exist
if (!(Test-Files $logFile, $monitorSettingsPath $emailSenderPath $emailRecipientsPath ))
{
    return
}

#get email settings
$emailSender = Get-Content  $emailSenderPath 
$emailRecipients = Get-Content $emailRecipientsPath 

#get list of errors to monitor
$monitorSettings = Import-Csv -Path  $monitorSettingsPath

#initiazlize list of alerts
$alerts = $null
$errorFileList = $null

#loop thru each error check
foreach ($monitor in $monitorSettings) 
{   
    #Check StartTime and StopTime to see if it's time to check for this error  
    $WithinStartStopTime = $true

    if ($monitor.StartTime.Trim() -ne "")
    {
        $startTime = Get-Date -Date "$($today) $($monitor.StartTime)"
        if ($startTime -gt $now)
        {
            Write-Log $logFile "Skipping $($monitor.Title) waiting for start time: $($startTime)"
            $WithinStartStopTime = $false
        }
    }

    if ($monitor.StopTime.Trim() -ne "")
    {
        $stopTime = Get-Date -Date "$($today) $($monitor.StopTime)"
        if ($stopTime -lt $now)
        {
            Write-Log $logFile "Skipping $($monitor.Title) - past stop time: $($stopTime)"
            $WithinStartStopTime = $false
        }
    }

    if  ($WithinStartStopTime)
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
        
         #FileExtensionFilter specifies filter
         if ($monitor.FileExtensionFilter.Trim() -ne "")
        {
            $gciFilter = " -Filter *.$($monitor.FileExtensionFilter.Trim()) "
        }
        else
        {
            $gciFilter = ""
        }

         #FileExcludeFilter specifies filter
         if ($monitor.FileExcludeFilter.Trim() -ne "")
        {
            $gciExclude = " -exclude *.$($monitor.FileExcludeFilter.Trim()) "
        }
        else
        {
            $gciExclude = ""
        }
    
    
        #if date/time thresholds are specified, build the Where clause
        $dateTimeThresholdHigh = $null
        $dateTimeThresholdLow = $null
    
        if ((Is-Numeric($monitor.dayThresholdLow)) -and ($monitor.dayThresholdLow -ne 0))
        {
            $dateTimeThresholdLow = $now.AddDays(0 - $monitor.dayThresholdLow)
        }

        if ((Is-Numeric($monitor.minuteThresholdLow)) -and ($monitor.minuteThresholdLow	-ne 0))
        {
            $dateTimeThresholdLow = $now.AddMinutes(0 - $monitor.minuteThresholdLow)
        }

        if ((Is-Numeric($monitor.dayThresholdHigh)) -and ($monitor.dayThresholdHigh	-ne 0))
        {
            $dateTimeThresholdHigh= $now.AddDays(0 - $monitor.dayThresholdHigh)
        }

        if ((Is-Numeric($monitor.minuteThresholdHigh)) -and ($monitor.minuteThresholdHigh	-ne 0))
        {
            $dateTimeThresholdHigh= $now.AddMinutes(0 - $monitor.minuteThresholdHigh)
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
        $monitorErrors  = $null
        $gciCommand = "`$monitorErrors  = Get-ChildItem -Path $($monitor.Folder) $($gciItemType) $($gciFilter) $($gciExclude) -Force -ErrorAction SilentlyContinue $($gciWhere)  | Select-Object Name, LastWriteTime"
        Invoke-Expression $gciCommand 

        if ($monitorErrors  -ne $null)
        {
            $alert = "$($monitor.Title) found at $($monitor.Folder)`t`r`n"
            $errorFiles = Out-String -InputObject $monitorErrors -stream
            Write-Log $logFile $alert
            Write-Log $logFile $errorFiles
            [string] $errorFileOutput = AddTabToLines $errorFiles
            $alerts += $alert
            $errorFileList += $errorFileOutput
            $errorFileList += "`t`r`n"
        }
        else
        {
            Write-Log $logFile "Passed: $($monitor.Title)"
        }
    }
}

#if there were alerts, send email
if ($alerts -ne $null)
{
    $alerts += $errorFileList
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $emailSender -To $emailRecipients -Subject "$($title) Alert" -Body $alerts
}
else
{
    Write-Log $logFile "Finished"
}
