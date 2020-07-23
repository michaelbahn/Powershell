cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt

$pathPflFtp = "\\evsappsshrpd01\pflftp_team$"
$pathPflFtp = Join-Path  $pathPflFtp  "DONE"
$dateFormatted =  $(get-date).ToString("MMddyyyy")
$cdia = 0
$bdia = 1
$fileNames = @("$($pathPflFtp)\PFL$($dateFormatted)C*.DIA", "$($pathPflFtp)\PFL$($dateFormatted)B*.DIA")
$missingFile = $false
$errorMessages = @()

if (!(Test-Path $pathPflFtp))
{
            $errorMessages += "Unable to access PFL-DIA folder at $($pathPflFtp)"
}
else
{

    foreach ($fileName in $fileNames)
    {
        if (!(Test-Path $fileName))
        {
            $fileSearch = $fileName.Substring($pathPflFtp.Length+1)
            $searchResults = Get-ChildItem -Name $fileSearch -Path $pathPflFtp   -File -Recurse
        
            if ($searchResults -eq $null)
            {
                $errorMessage = "File $($fileSearch) was not found:`r`n$($pathPflFtp)`r`n"
                $errorMessages += $errorMessage
                Write-Log $logFile $errorMessage
            }
            else
            {
                $errorMessage = "File $($searchResults.FullName) is not in DONE folder:`r`n$($pathPflFtp)`r`n"
                $errorMessages += $errorMessage
                Write-Log $logFile $errorMessage 
            }        
        }
        else
        {
            Write-Log $logFile "File exists: $($filename)"
        }        

    }

}

if (($errorMessages.Length -gt 0) -and ((get-date).DayOfWeek -ne "Monday"))  
{
    $errorMessages += "`r`n"
    $mailBody = $errorMessages | Out-String
    $subject = "No C or B File Present for $((get-date).ToString("M/d/yy")) in PFL-DIA Done folder"  
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $mailBody 
}

