  cls
#Initialize settings
$newLine = "`r`n"
$now = get-date
$today = get-date -Format  d
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"
#html
$htmlHeader = Get-Content htmlHeader.txt
$htmlTableHeaderFixed = "<table><tr><th>Errors Fixed</th><th>Date/Time</th></tr>"
$htmlTableHeaderTried = "<table><tr><th>Errors Not Fixed</th><th>Original Date/Time</th></tr>"
$htmlTableFooter = "</table>"
$htmlFooter = "</body></html>"

#get email settings
$sender = Get-Content  sender.txt
$recipients = Get-Content recipients.txt

#initialize log file 
Import-Module (Join-Path $modulePath Utilities.psm1) -Force 
$logFile = Initialize-Log $logPath $title

$fsetPath = Get-ContentWithComments fset-path.txt
$fsetErrorArchive = "$($fsetPath)\ErrorArchive"
if (!(Test-Path $fsetErrorArchive)) 
{
    New-Item -Path $fsetErrorArchive -ItemType Directory -Force
}

$fsetErrors = Get-ChildItem -path $fsetPath -Filter *.ERR*90* -File

if ($fsetErrors.Count -gt 0)
{
    #initiazlize list of errors
    $fsetErrorRecord = New-Object –TypeName PSObject –Prop (@{ 'FullName'=$null;'Name'=$null;'CreationTime' = $null;})
    $errorsFixed = @()
    $errorsTried = @()
    $message = @()
    $errorMessages = @()

    foreach ($fsetError in $fsetErrors)
    {
        $archivePath = Join-Path $fsetErrorArchive $fsetError.Name
        if (!(Test-Path $archivePath)) 
        {
            #record file info for email
            $fsetErrorRecord.FullName = $archivePath
            $fsetErrorRecord.CreationTime = $fsetError.CreationTime
            $fsetErrorRecord.Name = $fsetError.Name
            #convert ERR file name to xml file name
            $errIndex = $fsetError.Name.IndexOf(".ERR")
            $newName = $fsetError.Name.Substring(0, $errIndex) + ".xml"
            $newFullName = join-path $fsetPath $newName
            if (!(Test-Path $newFullName))
            {
                Write-Log $logFile "Copying $fsetError.FullName to $($archivePath)" 
                try 
                {
                    Copy-Item -Path $fsetError.FullName $archivePath
                    Write-Log $logFile "Renaming $($fsetError.FullName) to $($newFullName)" 
                    Rename-Item -Path $fsetError.FullName -NewName $newFullName -ErrorAction Stop
                    $errorsFixed += $fsetErrorRecord
                }
                catch 
                {
                    $errorMessage = "Renaming failed: $($fsetError.FullName) to $($newFullName)" 
                    $errorMessages += "<br><p>$($errorMessage)</p>"
                    Write-Log $logFile $errorMessage 
                 }
            }
             else
            {
                $errorMessage = "Unable to rename $($fsetError.Name) at $($fsetPath).  Duplicate xml file exists: $($newName).  " 
                $errorMessages += "<br><p>$($errorMessage)</p>"
                Write-Log $logFile $errorMessage 
            }
        }
        else
        {
            $archivedFile = Get-ItemProperty $archivePath
            $fsetErrorRecord.FullName = $fsetError.FullName
            $fsetErrorRecord.CreationTime = $archivedFile.CreationTime
            $fsetErrorRecord.Name = $fsetError.Name
            $errorsTried += $fsetErrorRecord
            Write-Log $logFile "Previuosly tried fix:  $($archivedFile.FullName)" 
        }
     }  #for

#build html
    if ($errorsFixed.Count -gt 0) 
    {
        $message = "$($htmlHeader)$($newLine)$($htmlTableHeaderFixed)"  

        foreach ($errorFixed in $errorsFixed)
        {
            $message += "<tr><td>$(Convert-FilePathToHTML $errorFixed)</td><td>$($errorFixed.CreationTime)</td></tr>"
        }
        
        $message += $htmlTableFooter
    }

    if ($errorsTried.Count -gt 0) 
   {
        if ($errorsFixed.Count -eq 0) 
        {
            $message = "$($htmlHeader)$($htmlTableHeaderTried)"    
        }
        else
        {
            $message += "<br>$($htmlTableHeaderTried)"  
        }

        foreach ($errorTried in $errorsTried)
        {
            $message += "<tr><td>$(Convert-FilePathToHTML $errorTried)</td><td>$($errorTried.CreationTime)</td></tr>"
        }

        $message += $htmlTableFooter
        $message += "<br><p>Errors not fixed have previously gone through the fix process.</p>"  
    }  
    if ($errorMessages.Count -gt 0)
    {
        $message += $errorMessages
    }
    $message += $htmlFooter
    $subject = "$($title) Alert"
     Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $message -BodyAsHtml
     Write-Log $logFile "Sent email to $($recipients):"
 }
else
{
    Write-Log $logFile "Finished: no FSET errors"
}

