#Create ACK for RBE Batch with single SUBMIT file
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = ".\Logs"
$fsetSQLconnection = "dgvmsqlfstpd01\DGC1VDBFSETPD02"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
Import-Module .\FSET-SQL.psm1 -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt

$htmlHeader = Get-Content htmlHeader.txt
$htmlTableHeaderTransmission = "<table><tr><th>Confirmation #</th><th>Status</th><th>Postmark</th><th>Duration (sec)</th></tr>"
$htmlTableHeaderElements = "<table><tr><th>Confirmation #</th><th>Status</th><th>Modified Date</th><th>Batch</th><th>File</th></tr>"
$htmlTableFooter = "</table>"
$htmlFooter = "</body></html>"

#check for errrs database
$transmissionErrors =Query-Transmission-Status $fsetSQLconnection

if ($transmissionErrors -ne  $null)
{
    $message = "$($htmlHeader)`r`n$($htmlTableHeaderTransmission)`r`n"
     
    foreach ($transmissionError in $transmissionErrors)
    {
        $message +=  "<tr><td>$($transmissionError.TransmissionConfirmationNumber)</td><td>$($transmissionError.TransmissionStatusDescription)</td><td>$($transmissionError.PostmarkedDate)</td><td>$($transmissionError.ProcessingTime)</td></tr>`r`n"
    }
   $message += "$($htmlTableFooter)<br><h3>Transmission Elements</h3>`r`n" 
   $message += "$($htmlTableHeaderElements)`r`n"

    #get transmission elements for each transmission error
    foreach ($transmissionError in $transmissionErrors)
    {
        $transmissionElements =Query-Transmission-Elements $fsetSQLconnection $transmissionError.TransmissionID
        foreach ($transmissionElement in $transmissionElements)
        {        
            $message += "<tr><td>$($transmissionElement.ConfirmationNumber)</td><td>$($transmissionElement.StatusCode)</td><td>$($transmissionElement.ModifiedDate)</td><td>$($transmissionElement.BatchNumber)</td><td>$($transmissionElement.ErrorFilePath)</td></tr>`r`n"
        }
    }
   $message += "$($htmlTableFooter)`r`n" 
   $message += "$($htmlFooter)`r`n" 

    $subject = "$($title) Alert"
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $message -BodyAsHtml
}  
else
{
    Write-Log $logFile "No FSET errors"
}
