cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"

Import-Module .\Alert-SEFT-functions.psm1 -Force
Import-Module (Join-Path $modulePath Utilities.psm1) -Force


$logFile= Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt
#$recipients = "michael.bahn@edd.ca.gov"

$pathAppLogFile = "\\dgvmappuipd01\DMS\DE1101CZ-SEFT-Logs"
$dateFormatFileSuffix =  (get-date).ToString("yyyyMMdd")
# for testing $dateFormatFileSuffix =  ((get-date).AddDays(-2)).ToString("yyyyMMdd")
$appLog = "DMSGFtpAppLogFile$($dateFormatFileSuffix).log"
$appLogFileName = "$($pathAppLogFile)\$($appLog)"
$oldAppLogFileName = "$($logPath)\$($appLog)"
$mailHeader = "<a href=$($appLogFileName)>$($appLog):<br /></a>"

if (Test-Path $appLogFileName)
{
    Write-Log $logFile "SEFT App Log file found at $($appLogFileName)"
    
    $appErrors = Get-Content $appLogFileName
    
    #check to see if errors have already been reported
    if (Test-Path $oldAppLogFileName)
    {    
        $oldAppErrors = Get-Content $oldAppLogFileName
        $newErrorCount = $appErrors.Count - $oldAppErrors.Count
        $startLine = $oldAppErrors.Count   #zero based
    }
    else
    {
        $newErrorCount = $appErrors.Count 
        $startLine = 0
    }
            

    if ($newErrorCount -gt 0)
    {    
        $newErrorMessageExists = $true
        if ($appErrors.Count -eq 1)
        {
            $emailMessage = $appErrors.ToString()

        }
        else
        {
            $emailMessage = ""
            for ($i=$startLine;$i -lt $newErrorCount; $i++)
            {
                $appError = $appErrors[$i].ToString()
                Write-Log $logFile "App Error: $($appError)"
                $emailMessage = "$($emailMessage)$($appError)<br />"
                Write-Log $logFile "Email Message: $($emailMessage)"
            }
        }

    }
    else
    {
        $newErrorMessageExists = $false
    }


    if ($newErrorMessageExists)
    {
        Write-Log $logFile $emailMessage
        $subject = "SEFT DE5617 Upload Alert"

        Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -BodyAsHtml  "$($mailHeader)$($emailMessage)"

        Copy-Item -Path $appLogFileName -Destination $logPath -Force
    }
    else
    {
        Write-Log $logFile "No new error messages"
    }

}
else
{
    Write-Log $logFile "No App Log file found at $($appLogFileName)."
}

