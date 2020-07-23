cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)

#production
$scriptLogDir = "d:\Logs"
$iifLogDir = "d:\temp"
$iifLogDir = "d:\temp"


Import-Module (Join-Path $scriptPath Utilities.psm1) -Force
$logFile= Initialize-Log $scriptLogDir $title
$lastErrorFile = Initialize-Log $scriptPath "LastError"
$htmlFileName =  join-path $scriptPath "$($title).htm"

$sender = "teaminf@edd.ca.gov"
$recipients = "teaminf@edd.ca.gov"
$recipients = "michael.bahn@edd.ca.gov"

$date = Get-Date

########## HTML Header ##########
Function writeHtmlHeader  
{  
    param($htmlFileName)  
    $date = Get-Date  
    Add-Content $htmlFileName "<html>" 
    Add-Content $htmlFileName "<head>" 
    Add-Content $htmlFileName "<META HTTP-EQUIV=refresh CONTENT=15>"
    Add-Content $htmlFileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
    Add-Content $htmlFileName '<title>ZIPWrap Service Monitor - Goethe</title>' 
    add-content $htmlFileName '<STYLE TYPE="text/css">' 
    add-content $htmlFileName  "<!--" 
    add-content $htmlFileName  "td {" 
    add-content $htmlFileName  "font-family: Tahoma;" 
    add-content $htmlFileName  "font-size: 11px;" 
    add-content $htmlFileName  "border-top: 1px solid #999999;" 
    add-content $htmlFileName  "border-right: 1px solid #999999;" 
    add-content $htmlFileName  "border-bottom: 1px solid #999999;" 
    add-content $htmlFileName  "border-left: 1px solid #999999;" 
    add-content $htmlFileName  "padding-top: 0px;" 
    add-content $htmlFileName  "padding-right: 0px;" 
    add-content $htmlFileName  "padding-bottom: 0px;" 
    add-content $htmlFileName  "padding-left: 0px;" 
    add-content $htmlFileName  "}" 
    add-content $htmlFileName  "body {" 
    add-content $htmlFileName  "margin-left: 5px;" 
    add-content $htmlFileName  "margin-top: 5px;" 
    add-content $htmlFileName  "margin-right: 0px;" 
    add-content $htmlFileName  "margin-bottom: 10px;" 
    add-content $htmlFileName  "" 
    add-content $htmlFileName  "table {" 
    add-content $htmlFileName  "border: thin solid #000000;" 
    add-content $htmlFileName  "}" 
    add-content $htmlFileName  "-->" 
    add-content $htmlFileName  "</style>" 
    Add-Content $htmlFileName "</head>" 
    Add-Content $htmlFileName "<body>" 
    
    add-content $htmlFileName  "<table width='100%'>" 
    add-content $htmlFileName  "<tr bgcolor='#D0E6FF'>" 
    Add-Content $htmlFileName  "<td></td>"
    add-content $htmlFileName  "<td width='40%' height='25' align='center'>" 
    add-content $htmlFileName  "<font face='tahoma' color='#003399' size='4'><strong><a href='\\DGVMWEBDIAPD01\IIFECMLogs'>IIF timeouts</strong></font>" 
    add-content $htmlFileName  "</td>" 
    Add-Content $htmlFileName  "<td></td>"
    Add-content $htmlFileName  "</tr>" 

   
 }  
    
 ########## Table Header ##########
 Function writeTableHeader  
 {  
 param($htmlFileName)  
    
 Add-Content $htmlFileName "<tr bgcolor=#D0E6FF>" 
 Add-Content $htmlFileName "<td width='25%' align='center'><strong>Date Time</strong></td>" 
 Add-Content $htmlFileName "<td width='15%' align='center'><strong>User</strong></td>" 
 Add-Content $htmlFileName "<td width='20%' align='center'><strong>DLN</strong></td>" 
 Add-Content $htmlFileName "<td width='40%' align='center'><strong>Message</strong></td>" 
 Add-Content $htmlFileName "</tr>" 
 }  
    
########## HTML Footer ########## 
Function writeHtmlFooter  
 {  
 param($htmlFileName) 
 #Add-Content $htmlFileName "<table width='100%'><tbody>" 
 #Add-Content $htmlFileName "<tr bgcolor='#D0E6FF'>" 
 #Add-Content $htmlFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>End Report</strong></font></td>" 
 #Add-Content $htmlFileName "</tr>"    
 Add-Content $htmlFileName "</body>" 
 Add-Content $htmlFileName "</html>" 
 }  
    

############## Fix a Batch  ###############
function FixBatch ($batch, [ref]$fixMessage) 
{  

    Copy-Item -Path $batch.FullName -Destination $workingFolder -Force -Recurse
    $batchWorkingFolder = Split-Path $batch.FullName -leaf
    $batchWorkingPath = Join-Path $workingFolder $batchWorkingFolder
    $errorTextFile = "$($batchWorkingPath)\error.txt"
    
    $okToMoveBatch = $false

    #if can't find error file then leave
    if (!(Test-Path $errorTextFile))
    {
        $log = "Can't find error file $($errorTextFile)"
        $fixMessage.Value = "$($log) <br>"
        Write-Log $logFile $fixMessages 
        return 
    }

    $errorTexts = Get-Content -Path $errorTextFile

    foreach ($errorText in $errorTexts)
    {
        $endOfInvalidFileTypeMessage = "is an invalid file types"
        $zipWrapErrorString = "ZipWrap-Error\"

        $indexEndOfInvalidFileTypeMessage = $errorText.IndexOf($endOfInvalidFileTypeMessage)
        if ($indexEndOfInvalidFileTypeMessage -ge 0)
        {
            $errorText = $errorText.Substring(0,$indexEndOfInvalidFileTypeMessage - 1)

            $indexZipWrapError = $errorText.IndexOf($zipWrapErrorString)
            $completePath = $errorText.SubString($indexZipWrapError+ $zipWrapErrorString.Length)
            $files = $completePath.Split("\")
            $doneFolder = join-path $sidesErrorDir $files[0] 
            $zipBatch = Join-Path $batchWorkingPath "$($files[1]).zip"
            
            #verify zip file exists
            if (!(Test-Path $zipBatch))
            {
                $log = "Cannot find $($zipBatch) in $($doneFolder) "
                Write-Log $logFile "$($log)"
                $fixMessage.Value += "$($log) <br> "
                $okToMoveBatch = $false
                return
            }

            #unzip files unless they are already unzipped
            $tempWorkingFolder = Join-Path $workingFolder $files[1]
            if (!(Test-Path $tempWorkingFolder))
            {
                Expand-Archive -LiteralPath $zipBatch -DestinationPath $tempWorkingFolder
            }

            $fileToConvert = Join-Path $tempWorkingFolder $files[2]
            $csvFileIndex = $fileToConvert.IndexOf(".csv")
            $rtfFileIndex = $fileToConvert.IndexOf(".rtf")
            $pdfFileName = ""

            try 
            {
                if ($csvFileIndex -ge 0)
                {
                    ConvertCsvToPDF $fileToConvert ([ref] $pdfFileName)
                    $pdfFileLeaf = Split-Path -Path $pdfFileName -Leaf
                    $log = "Converted $($files[2]) to $($pdfFileLeaf) "
                    Write-Log $logFile "$($log)"
                    $fixMessage.Value += "$($log) <br> "
                    $okToMoveBatch = $true

                }
                elseif ($rtfFileIndex -ge 0)
                {
                    ConvertRtfToPDF $fileToConvert ([ref] $pdfFileName) 
                    $log = "Converted $($files[2]) to $($pdfFileName) "
                    Write-Log $logFile $log
                    $fixMessage.Value += "$($log) <br> "
                    $okToMoveBatch = $true
                }
                else
                {
                    $fixMessage.Value += "Unable to fix <br> "
                    Write-Log $logFile "Unable to fix: $($errorText)"
                }


                if (Test-Path $pdfFileName)
                {
                    Set-Location $batch.FullName
                    $zipBatchPath = join-path $batch.FullName $zipBatchPath
              
                    Remove-Item -Path $fileToConvert -Force 
                    Move-Item -Path $pdfFileName -Destination $tempWorkingFolder -Force 
                    #create zip file
                    $newZip = "$($tempWorkingFolder).zip"
                    Compress-Archive -Path "$($tempWorkingFolder)\*.*" -DestinationPath $newZip 
                    Copy-Item -Path $newZip -Destination $batch.FullName -Force 
                    Remove-Item -Path $newZip -Force                   #remove temp zip that was copied to batch
                }
                else
                {
                    $log = "Unable to create PDF for: $($files[2])"
                    $fixMessage.Value += "$($log) <br> "
                    Write-Log $logFile $log 
                    $okToMoveBatch = $false
                }

            }
		    catch 
            {
			    $errorObject = $_
                $log = "Error with $($pdfFileName) in $($newZip)`t`r`n$($errorObject)"
                Write-Log $logFile $log
                $fixMessage.Value += "$($log) <br> "
                $okToMoveBatch = $false
		    }


        }  #if

    } #for

    if ($okToMoveBatch)
    {    
        Remove-Item -Path "$($batch.FullName)\error.txt"
    }
}  
 

########################### MAIN ###########################

$errorCount = 0
$lastError = get-content $lastErrorFile
$lastErrorDateTimeString = get-content $lastErrorFile
[datetime]$lastErrorDateTime = $lastErrorDateTimeString

$iifLog = get-content $iifLogFile

for ($i=0;$i -lt $iifLog.count;$i++)
{
    if ($iifLog[$i].IndexOf("|") -ge 0)
    {
        $iifLogRow = ($iifLog[$i]).split("|")

        $dateTime = $iifLogRow[0]
        $userName = $iifLogRow[1]
        $dln = $iifLogRow[3]
        $message = $iifLogRow[4]

        if ([datetime]$dateTime -gt $lastErrorDateTime)
        {
            $errorCount =+ $errorCount

            Add-Content $htmlFileName "<tr>" 
 	        Add-Content $htmlFileName "<td width='25%' align='left'>$($dateTime)</td>" 
            Add-Content $htmlFileName "<td width='15%' align='left'>$($userName)</td>" 
   	        Add-Content $htmlFileName "<td width='20%' align='left'>$($dln)</td>" 
   	        Add-Content $htmlFileName "<td width='40%' align='left'>$($message)</td>" 
            Add-Content $htmlFileName "</tr>" 
        }
    }

}








$sidesErrorFiles = Get-ChildItem -Path $sidesErrorDir -Directory
$sidesErrorCount = $sidesErrorFiles.Count
Write-Log $logFile "SIDES Error count at $($sidesErrorDir) is $($sidesErrorCount)"

if($sidesErrorCount -le 0)
{
    return
}

$objExcel = New-Object -ComObject excel.application
$cleanupFiles = @()
$fixedBatchCount = 0

#clean up temp directory
Get-ChildItem -Path $workingFolder -File | Remove-Item -Force -Recurse
Get-ChildItem -Path $workingFolder -Directory | Remove-Item -Force -Recurse

#create html file in temp directory
New-Item -itemtype file -Path $htmlFileName -Force
writeHtmlHeader $htmlFileName
writeTableHeader $htmlFileName 


foreach ($batch in $sidesErrorFiles)
{

    Add-Content $htmlFileName "<tr>" 
 	Add-Content $htmlFileName "<td width='20%' align='left'><a href='file:$($batch.FullPath)'>$($batch.Name)</a></td>" 

    #if can't find error file then leave
    $errorTextFile = "$($batch.FullName)\error.txt"
    if (!(Test-Path $errorTextFile))
    {
        $errorMessages += "Can't find error file $($errorTextFile)"
        Write-Log $logFile $errorMessages 
        $fixMessage = "Missing error file"
    }
    else
    {
        #backup files that will be updated
        if (!(Test-Path "$($sidesErrorBackup)\$($batch.Name)"))
        {
            Copy-Item $batch.FullName $sidesErrorBackup -Recurse
        }

        #read error text file
        $errorMessages = @()
        $fixMessages = @()
        $errors = Get-Content -Path $errorTextFile
        foreach ($errorMessage in $errors)
        {
            $errorMessages += "$($errorMessage) <br>"
        }

        #Fix Batch
        Write-Log $logFile "Fix Batch $($batch.FullName)"
        $fixMessage = ""
            
        $okToMoveBatch = FixBatch $batch ([ref]$fixMessage)
            
        if ($okToMoveBatch)
        {
            #move item in two steps
            Copy-Item $batch.FullName $sidesInputDir  -Recurse

            $cleanupFiles += $batch.FullName
            $fixedBatchCount += 1
        }
                        
    } #else

    Add-Content $htmlFileName "<td width='40%' align='left'>$($errorMessages)</td>" 
   	Add-Content $htmlFileName "<td width='40%' align='left'>$($fixMessage)</td>" 
    Add-Content $htmlFileName "</tr>" 

}  #foreach

if ($objWord -ne $null)
{
    $objWord.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWord)
}

if ($objExcel -ne $null)
{
    $objExcel.Quit
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
}


#remove fixed files from Zip Error folder
foreach ($cleanupFile in $cleanupFiles)
{
    Remove-Item $cleanupFile -Force -Recurse
    Write-Log $logFile "Cleaned up: $($cleanupFile)"
}

#alert if fixed files are still in Zip Error folder
foreach ($cleanupFile in $cleanupFiles)
{
    if (test-path ($cleanupFile.FullName))
    {
        $cleanupError = "Folder not cleaned up: $($cleanupFile.FullName)"
        Write-Log $logFile $cleanupError
        Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To "michael.bahn@edd.ca.gov" -Subject $cleanupError  -Body $cleanupError 
    }
}

writeHtmlFooter  $htmlFileName
$mailBody = Get-Content $htmlFileName  -Raw

$unfixedBatchCount = $sidesErrorCount - $fixedBatchCount

if(($unfixedBatchCount -le 0) -and ($sidesErrorCount -gt 1))
{
    $emailSubject = "Fix-SIDES fixed $($sidesErrorCount) ZipWrap Batches" 
}
elseif (($unfixedBatchCount -le 0) -and ($sidesErrorCount -eq 1))
{
    $emailSubject = "Fix-SIDES fixed $($sidesErrorCount) ZipWrap Batch"
}
elseif (($unfixedBatchCount -gt 0) -and ($unfixedBatchCount -eq 1))
{
    $emailSubject = "SIDES ZipWrap Errors: $($unfixedBatchCount) SIDES batch in ZipWrap-Error folder" 
}
else
{
    $emailSubject = "SIDES ZipWrap Errors: $($unfixedBatchCount) SIDES batches in ZipWrap-Error folder" 
}

    
Write-Log $logFile $logMessage
     	

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $emailSubject  -Body $mailBody -BodyAsHtml
 

