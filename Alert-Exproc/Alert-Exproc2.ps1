﻿cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"

#Import-Module .\Alert-SEFT-functions.psm1 -Force
Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = "teaminf@edd.ca.gov"
$recipients = Get-Content  .\recipients.txt

$xprcLogPath = "\\DGVMAPPTAXPD01\DMs\EXCPPROC\Logs" 
$dateFormatFileName =  (get-date).ToString("yyMMdd")
$xprcLogFileName = "$($xprcLogPath)\$($dateFormatFileName).Log"

function removeLogNoise ($inputString)
{
    $newString = "  $($inputString.SubString(2, $inputString.Length - 2))"
    #$linePrefix = $newString.SubString(0, 81)
    $linePrefix = "$($xprcLogFileName):"
    $inf = "INF "
    $info = "INFO:"
    $removeTokens = {$linePrefix, $inf, $info}
    foreach ($removeToken in $removeTokens)
    {
        while ($newString.IndexOf($removeToken) -ge 0)
        {
            $newString = $newString -replace [regex]::escape($removeToken), ''
        }
    }
    return $newString
}

$xprcMessages = @()
$errorMessages = @()
$errorMessage = ""
$xprcMessage = ""
$xprcError = $false
$keepTrying = $true
$xprcFileExists = $false
$numberTries = 0
$tryLimit = 30
$secondsWait = 60

while ($keepTrying -and ($numberTries -lt $tryLimit))
{
    if (Test-Path $xprcLogFileName )
    {
        $xprcFileExists = $true
        $keepTrying = $false
        $xprcMessage = "EXPROC log file found at $($xprcLogFileName)"
        Write-Log $logFile $xprcMessage 
        $xprcMessages += $xprcMessage
    }
    else
    {
        Write-Log $logFile "EXPROC log file not found at $($xprcLogFileName), will try again in $($secondsWait.ToString()) seconds."
        Start-Sleep -s $secondsWait 
        $numberTries++
    }
}

if ($xprcFileExists)
{                    
    $token = "SESSION SUMMARY"
    $xprcLogFile = Get-Content $xprcLogFileName | Out-String
    $indexSessionSummary = $xprcLogFile.IndexOf($token)

    if ($indexSessionSummary -ge 0)
    {
        $report = Select-String $token -path $xprcLogFileName -Context 0,10 -SimpleMatch
        Write-Log $logFile $report
        $mailBody = removeLogNoise $report.ToString()    
        $xprcError = $false
    }
    else
    {
        $errorMessage = "Missing xprc log file $xprcLogFile$($dateFormatFileSuffix).Log at $($xprcLogFileName)"
        $errorMessages += $errorMessage
        Write-Log $logFile $errorMessage
        $xprcError = $true
    }
}
else
{
        $xprcError = $true
}

    
if ($xprcError)
{
     $mailBody = $errorMessages | Out-String
     $mailBody = "Log Folder: $($xprcLogPath)`r`n`r`n$($mailBody)"
     $subject = "EXPROC log error for $((get-date).ToString("M/d/yy"))"
}
else
{
     $mailBody = "Log File: $($xprcLogFileName)`r`n$($mailBody)"
     $subject = "EXPROC processing for $((get-date).ToString("M/d/yy"))"
}

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $mailBody
