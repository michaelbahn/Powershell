#reads a log file and converts to html page and opens in a web browser
cls
$scriptPath = Split-Path $MyInvocation.MyCommand.Path
Set-Location $scriptPath
$logPath = "..\Logs"
$logFileName = "BenefitsPurge"
$todaysDate = (Get-Date)
$dateFormatted = $todaysDate.tostring("yyyyMMdd")
$logFileName += $dateFormatted
$logFileName += ".log"
$logFileName = join-path $logPath $logFileName
$reportFileName = "report.html"
$reportFilePath = "..\Output"
$reportFileName = join-path $reportFilePath $reportFileName

function TextToHtml($logFileName, $reportCaption)
{
    $TextData = Get-Content $logFileName

    $lineData =  [string]::Empty
    Foreach ($line in $TextData) {
      $lineData += $line
    }

    Get-Content $logFileName | ConvertTo-HTML -Property @{Label=$reportCaption;Expression={$_}} | Out-File $reportFileName
   }

function EmailReport($reportFileName, $mailSubject)
{
$scriptPath = Get-Location
$settingsPath = "..\Config"

$sender = gc  (join-path $settingsPath Sender.txt)
$recipients = gc  (join-path $settingsPath Recipients.txt)

IF ([string]::IsNullOrWhitespace($sender))
{
    Write-Host "Invalid sender list at " $settingsPath\Sender.txt
    return
} 

IF ([string]::IsNullOrWhitespace($recipients))
{
    Write-Host "Invalid receipient list at " $settingsPath\Recipients.txt
    return
} 
 
$mailBody = Get-Content $reportFileName  -Raw

Send-MailMessage -SmtpServer 'smtp.edd.ca.gov' -From $sender -To $recipients -Subject $mailSubject -Body $mailBody -BodyAsHtml
}

TextToHtml $logFileName "Benefits InfoImage Purge Log - $todaysDate" $reportFileName

$mailSubject = "Benefits Purge Report for " + $todaysDate.ToString("MM/dd/yy")
EmailReport $reportFileName $mailSubject