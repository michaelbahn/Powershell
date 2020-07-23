﻿cls
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
$fileAlerts = $null

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

         #build FileName filter
         $filterRaw = $monitor.FileName.Trim()
         $filterIndex = $filterRaw.IndexOf("%")
         
         if ($filterIndex -ge 0)
         {
             $dateFormat = $monitor.DateFormat.Trim()
             if ( $dateFormat -ne "")
            {
                $dateValue = get-date -format $dateFormat
            }
            else
            {
                $fileAlerts = "Missing DateFormat: $($monitor.Title)"
                Write-Log $logFile $fileAlerts
            }
            $filter = $filterRaw.Replace("%", $dateValue)      
        }
        else
        {
            $filter = $filterRaw
        }

        $gciFilter = " -Filter $($filter)"

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
        $monitorFiles  = $null
        $gciCommand = "`$monitorFiles  = Get-ChildItem -Path $($monitor.Folder) $($gciFilter) -Force -ErrorAction SilentlyContinue $($gciWhere)  | Select-Object FullName, LastWriteTime, Length"
        Invoke-Expression $gciCommand 

        if ($monitorFiles  -eq $null)
        {
            $alert = "$($monitor.Title):`t`r`nFolder:`t$($monitor.Folder)`t`r`nFile not found:`t$($monitor.FileName)`t`r`n"
            $fileAlerts += $alert
            Write-Log $logFile $alert
        }
        else
        {
            if ([string]::IsNullOrWhiteSpace($monitor.SizeLimitMB))
            {
                Write-Log $logFile "Passed: $($monitor.Title)"
            }
            else
            {
                [double] $fileSize = $monitorFiles[0].length / 1000000

                if ($fileSize -ge $monitor.SizeLimitMB)
                {
                    $fileAlerts += "$($monitor.Title): `t$($monitor.FileName) size of $($fileSize) exceeds size limit of $($monitor.SizeLimitMB)MB`r`n"
                    Write-Log $logFile $fileAlerts                       
                }
                else
                {
                    Write-Log $logFile "Passed: $($monitor.Title)"       
                }
                
            }

        }
    }
}

#if there were problem batches, send email
if ($fileAlerts -ne $null)
{
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $emailSender -To $emailRecipients -Subject "$($title) Alert" -Body $fileAlerts
}
else
{
    Write-Log $logFile "Finished"
}