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

$pathSeftFile = "\\dgvmappentpd01\DMS\SEFT\SEFTLogs"
$pathSeftToXprc = "\\dgvmappentpd01\DMS\SEFT-XPRC2\BWS"
$pathOutOfBalance = "\\dgvmappentpd01\DMS\SEFT-XPRC2\BWS\OutofBalance"
$dateFormatFileSuffix =  (get-date).ToString("yyyyMMdd")
# for testing $dateFormatFileSuffix =  ((get-date).AddDays(-2)).ToString("yyyyMMdd")
$seftFileName = "$($pathSeftFile)\DMSGFtpAppLogFile$($dateFormatFileSuffix).Log"

$seftMessages = @()
$errorMessages = @()
$errorMessage = ""
$seftMessage = ""
$seftError = $false
$keepTrying = $true
$seftFileExists = $false
$numberTries = 0
$tryLimit = 20

$outOfBalanceFiles = Get-ChildItem $pathOutOfBalance -File
if ($outOfBalanceFiles.Count -gt 0)
{
    $outOfBalanceMessage = $outOfBalanceFiles -join ' <br /> '
    $subject = "Quarterly Out of Balance files have arrived"
    $recipients = "michael.bahn@edd.ca.gov"
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -BodyAsHtml $outOfBalanceMessage
}

while ($keepTrying -and ($numberTries -lt $tryLimit))
{
    if (Test-Path $seftFileName)
    {
        $seftFileExists = $true
        $keepTrying = $false
        $seftMessage = "SEFT log file found at $($seftFileName)"
        Write-Log $logFile $seftMessage 
        $seftMessages += $seftMessage
    }
    else
    {
        Write-Log $logFile "SEFT log file found at $($seftFileName), will try again in 30 seconds."
        Start-Sleep -s 30
        $numberTries++
    }
}

if ($seftFileExists)
{                    
    $seftLogFile = Get-Content $seftFileName -Raw
    $indexSeft = $seftLogFile.IndexOf("Successful")

    if ($indexSeft -ge 0)
    {
        $seftMessage = $seftLogFile.Substring(0, $indexSeft + 11)
        $seftMessages += $seftMessage
        Write-Log $logFile $seftMessage
    }
    else
    {
        $errorMessage = "Missing close quote: $($seftLogFile)"
        $errorMessages += $errorMessage
        Write-Log $logFile $errorMessage
        $seftError = $true
    }
}
else
{
    $errorMessage = "Missing SEFT log file DMSGFtpAppLogFile$($dateFormatFileSuffix).Log at $($pathSeftFile)"
    $errorMessages += $errorMessage
    Write-Log $logFile $errorMessage
    $seftError = $true
}

if (!$seftError)
{
    $token = "Successfully Deleted the file - '"
    $mainframeFiles = @()
    $endOfFile = $false

    while (!$endOfFile)
    {    
        $indexSuccess = $seftLogFile.IndexOf($token)
        if ($indexSuccess -ge 0)
        {
            $newSeftLogFile = $seftLogFile.Substring($indexSuccess+$token.Length)
            $indexCloseQuote = $newSeftLogFile.IndexOf("'")
            if ($indexCloseQuote -ge 0)
            {
                $mainframeFiles += $newSeftLogFile.Substring(0, $indexCloseQuote)
            }

            $seftLogFile = $newSeftLogFile 
        }
        else
        {
            $endOfFile = $true
        }
    }

    #find files
    foreach ($mainframeFile in $mainframeFiles)
    {
        #get the correct folder for the file
        $searchPath = Get-SearchPath $pathSeftToXprc $mainframeFile
        $searchResults = Get-ChildItem -Name "$($mainframeFile)*" -Path $searchPath -File 
        #should have two hits
        if ($searchResults.Count -ne 2)
        {
            if ($searchResults.Count -le 1)
            {
                if ($searchResults.Count -eq 1)
                {
                    if ($searchResults.Name.IndexOf(".STS") -lt 0)
                    {
                        $errorMessage = "Missing STS file at $($searchResults.Directory)"
                    }
                    else
                    {
                        $errorMessage = "Missing file at $($searchResults.Directory)"
                    }
                }
                else  #zero
                {
                        $errorMessage = "Missing files for $($mainframeFile) at $($pathSeftToXprc)"
                }
            }
            else  #more than 2?
            {
                $errorMessage = "Too many file matches ($($searchResults.Count)) for $($mainframeFile) at $($pathSeftToXprc)"
            }
                            
            $errorMessages += $errorMessage
            Write-Log $logFile $errorMessage
        }
        else
        {
            $seftMessage = "$($mainframeFile) files found: $($searchResults[0]) and $($searchResults[1]) "        
            $seftMessages += $seftMessage
            Write-Log $logFile $seftMessage
        }
    }   #for
} #if

if ($errorMessages.Count -gt 0)
{
     $mailBody = $errorMessages | Out-String
     if ($seftMessages.Count -gt 0)
     {
        $mailBody = "$($seftMessages | Out-String)`r`n$($mailBody)"
     }
     $subject = "SEFT Processing for $((get-date).ToString("M/d/yy"))"
}
else
{
     $mailBody = $seftMessages | Out-String    
     $subject = "SEFT Successful Processing for $((get-date).ToString("M/d/yy"))"
}

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $mailBody
