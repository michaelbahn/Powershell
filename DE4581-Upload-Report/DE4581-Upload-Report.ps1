cls
$scriptPath = Split-Path $MyInvocation.MyCommand.Path
Set-Location $scriptPath
$newLine = "`r`n"

#log file settings
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
#email
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt

#find Benefits Purge log file 
$reportPath = "\\dgvmappuipd01\DMS\ICM-IEMUI\Reports\REPORTS_PROD"
$reportName = "DE4581UploadReport"
$todaysDate = (Get-Date)
$dateFormatted = $todaysDate.ToString("yyyyMMdd")

$reportFileName = "$($reportName).$($dateFormatted).TXT"
$reportPathFileName = Join-Path $reportPath  $reportFileName
$dateFormattedEmail = $todaysDate.ToString("MM/dd/yyyy")
$subject = "ICM-IEMUI Report $($todaysDate.ToString($dateFormattedEmail))"

$scanners = @("SCANNER4", "SCANNER5", "SCANNER6", "SCANNER9")
$textData = Get-Content $reportPathFileName -Raw
$lastIndexFound = 0

foreach ($scanner in $scanners)
{
    $indexScanner = $textData.LastIndexOf($scanner)
    Write-Log $logfile "$($scanner) index: $($indexScanner)"

    if ($indexScanner -gt $lastIndexFound)
    {
        $lastIndexFound = $indexScanner
    }
}

if ($lastIndexFound -gt 0)
{
    $indexBatchTime = $textData.IndexOf("Create time", $lastIndexFound) + 13
    $indexBatchNumber = $textData.IndexOf("DE4581", $lastIndexFound) + 7
    $batchNumber = $textData.Substring($indexBatchNumber, 12)
    $batchTime = $textData.Substring($indexBatchTime, 8)
 
    $mailBody  =  "Last batch $($batchNumber) processed at $batchTime$($newLine)$($newLine)"
     Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -Attachments $reportPathFileName 
     Write-Log $logfile "Email sent: $($mailBody)"
}    
else
{
       Write-Log $logfile "No user scanner entries in $($textData)"
}    

