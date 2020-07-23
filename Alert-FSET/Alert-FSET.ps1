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

$pathFsetIcm = "\\dgvmfstwebpd01\H$\Shares\Team\iCapture-FSET_ICM\Error"
$errorExtension =  "*.ERR"

$errorFile = ""
$errorFiles = @()
$errorMessages = @()
$errorMessage = ""

$errorFiles = Get-ChildItem -Path $pathFsetIcm -Filter $errorExtension
if ($errorFiles.Length -gt 0)
{
    $fsetError = $true
    
    foreach ($errorFile in $errorFiles)
    {
        Write-Log $logFile $errorMessage
        $lengthBusinessName = String.Length("business_name") + 1
        $errorFixed = $false
        $rows = Get-Content $errorFile
        foreach ($row in $rows)
        {
            $indexBusinessName = $row.IndexOf("business_name")
            $indexEOL = $row.LastIndexOf("|")

            if ($indexBusinessName -ge 0)
            {
                if ($indexEOL -gt 60)
                {
                    Write-Log $logFile "Business name exceeds length limit: $($row)"
                    $row = "$($row.Substring(0,59))|"
                    Write-Log $logFile "Business name truncated: $($row)"
                    $errorFixed = $true
                }
                else
                {
                    Write-Log $logFile "Business name OK: $($row)"
                }
            }
        }

        if ($errorFixed) 
        {
            Set-Content $errorFile $rows
            $errorMessage = "Fixed FSET file:`t$($errorFile.FullName)`r`n"
            Write-Log $logFile $errorMessage
            $errorMessages  += $errorMessage 
        }
        else
        {
            $errorMessage = "FSET error file needs to be fixed:`t$($errorFile.FullName)`r`n"
            $errorMessage  += $errorMessage 
        }            
    }
}

else
{
    $fsetError = $false
    Write-Log $logFile "No error files found at $($pathFsetIcm)."
}


if ($errorMessages.Count -gt 0)
{
     $mailBody = $errorMessages | Out-String
     if ($FSETMessages.Count -gt 0)
     {
        $mailBody = "$($FSETMessages | Out-String)`r`n$($mailBody)"
     }
     $subject = "FSET Processing for $((get-date).ToString("M/d/yy"))"
}
else
{
     $mailBody = $FSETMessages | Out-String    
     $subject = "FSET Successful Processing for $((get-date).ToString("M/d/yy"))"
}

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $mailBody
