cls
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

$taxLogPath = "\\DGVMTRPTPD01\d$\DE9ADJIIF\Logs" 
$dateFormatted =  (get-date).ToString("MMddyyyy")
$taxLogFile =  "DMSIIFDE9ADJTransLog$($dateFormatted).log"
$taxLogFileName = join-path $taxLogPath $taxLogFile

$taxMessages = @()
$errorMessages = @()
$errorMessage = ""
$taxError = $false
$keepTrying = $true
$taxFileExists = $false
$numberTries = 0
$tryLimit = 2
$secondsWait = 20
$service = "DMSIIFDE9ADJExtractWinService"
$server = "DGVMTRPTPD01"

while ($keepTrying -and ($numberTries -lt $tryLimit))
{
    if (Test-Path $taxLogFileName )
    {
        $taxFileExists = $true        $keepTrying = $false
        $taxMessage = "DE9 log file found at $($taxLogFileName)"
        Write-Log $logFile $taxMessage 
    }
    else
    {
        Write-Log $logFile "DE9 log file not found at $($taxLogFileName), will restart service."
        $numberTries++
        $job = get-ciminstance win32_service -filter "Name='$($service)'" -comp $server  | Invoke-CimMethod -Name StopService
        Start-Sleep -s $secondsWait
        $job = get-ciminstance win32_service -filter "Name='$($service)'" -comp $server  | Invoke-CimMethod -Name StartService        Start-Sleep -s $secondsWait

    }
}

if (!($taxFileExists))
{                    
    $errorMessage = "Missing Tax DE9 Adjustment log file $($taxLogFileName)"
    Write-Log $logFile $errorMessage
    $subject = "Tax DE9 Adjustment process did not run for $((get-date).ToString("M/d/yyyy"))"
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $errorMessage
}

