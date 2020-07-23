cls
$scriptPath = Split-Path $MyInvocation.MyCommand.Path
Set-Location $scriptPath
$newLine = "`r`n"
$settingsPath = "..\Config"

#email settings

$recipientsTo = $null
$recipientsCC = $null
$sender = Get-Content  (join-path $settingsPath SenderPurgeReport.txt)
$recipients = Get-Content  (join-path $settingsPath RecipientsPurgeReport.txt)
Foreach ($recipient in $recipients) 
{
    if ($recipient.Substring(0,3) -eq "cc:")
        {$recipientsCC += $recipient.Substring(3) }
    else
        {$recipientsTo += $recipient}
}
 
#log file settings
$logPath = "..\Logs"
$logFile = "BenefitsPurggeReport-" + (get-date -format yyyy-MM-dd-HH-mm) + ".log"
$logFile = Join-Path $logPath $logFile

#find Benefits Purge log file 
$purgeReportPath = "\\Dgvmbrptpd01\benefitsexe\Log"
$purgeReportFileName = "BenefitsPurge"
$todaysDate = (Get-Date)
$dateFormatted = $todaysDate.ToString("yyyyMMdd")
$purgeReportFileName += $dateFormatted
$purgeReportFileName += ".log"
$purgeReportFileName = Join-Path $purgeReportPath $purgeReportFileName

$subject = "Benefits Purge Report for " + $todaysDate.ToString("MM/dd/yy")

$textData = Get-Content $purgeReportFileName

$mailBody  =  "Please see attached $($subject) $($newLine)$($newLine)"
$mailBody += "`r`n"
$mailBody  +=  "Summary Information from report:"
$mailBody += "`r`n"
$ignore = $true

Foreach ($line in $TextData) 
{   #don't include date/time stamp
    $line = $line.Substring(26) 

    if ($line.IndexOf("Total number of documents deleted") -ge 0)
    {
        $ignore = $false
    }
    if ($line.IndexOf("Processing Purge Type") -ge 0) 
    {
        $mailBody += "`r`n"
        $ignore = $false
    }
    if ($line.IndexOf("***********************************") -ge 0) 
    {
        $ignore = $true
    }
    if ($line.IndexOf("registry settings") -ge 0  ) 
    {
        $ignore = $true
    }
    

    if (!$ignore)
    {    
         if (![string]::IsNullOrWhiteSpace($line))
        {
            $mailBody += $line
            $mailBody += "`r`n"
        }
    }
}

 Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $sender -To $recipientsTo -Cc $recipientsCC -Subject $subject -Body $mailBody -Attachments $purgeReportFileName
