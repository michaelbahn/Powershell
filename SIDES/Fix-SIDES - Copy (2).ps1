cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$objWord = $null
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)

#EXCEL Constants
$xlLandscape = 2
$xlPaperLegal = 5
$xlFixedFormatQuality = 0 
$xlIncludeDocProperties = $false 
$xlIgnorePrintAreas = $false 
$xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]

#production
$sidesInputDir = "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\Input"
$sidesErrorDir = "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\ZipWrap-Error"
$sidesErrorBackup = "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\Input-error"

#testing
#$sidesInputDir = "d:\Input"
#$sidesErrorDir = "d:\ZipWrap-Error"
#$sidesErrorBackup = "d:\Input-error"

$workingFolder = "D:\SIDES-Temp"
$htmlFileName =  join-path $workingFolder "SIDES.htm"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile= Initialize-Log $logPath $title

$sender = "teaminf@edd.ca.gov"
$recipients = "teaminf@edd.ca.gov"
#$sender = "michael.bahn@edd.ca.gov"
#$recipients = "michael.bahn@edd.ca.gov"


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
    add-content $htmlFileName  "<td width='45%' height='25' align='center'>" 
    add-content $htmlFileName  "<font face='tahoma' color='#003399' size='4'><strong><a href='\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\ZipWrap-Error'>SIDES ZipWrap Errors </strong></font>" 
    add-content $htmlFileName  "</td>" 
    Add-Content $htmlFileName  "<td></td>"
    Add-content $htmlFileName  "</tr>" 
    #add-content $htmlFileName  "</table>" 

    #Add-Content $htmlFileName "<tr bgcolor='#D0E6FF'>" 
    #Add-Content $htmlFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>  </strong></font></td>" 
    #Add-Content $htmlFileName "</tr>" 
   
 }  
    
 ########## Table Header ##########
 Function writeTableHeader  
 {  
 param($htmlFileName)  
    
 Add-Content $htmlFileName "<tr bgcolor=#D0E6FF>" 
 Add-Content $htmlFileName "<td width='20%' align='center'><strong>Batch</strong></td>" 
 Add-Content $htmlFileName "<td width='45%' align='center'><strong>Error</strong></td>" 
 Add-Content $htmlFileName "<td width='35%' align='center'><strong>Fix</strong></td>" 
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
    
############## RTF ###############
function ConvertRtfToPDF ($rtfFile, [ref] $pdfFileName)
{
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word") | Out-Null
    
    $pdfFile = Get-ChildItem $rtfFile | % {$_.BaseName}
    $pdfFileName.Value = "$($workingFolder)\$($pdfFile).pdf"
    
    if ($objWord -eq $null)
    {
        $objWord = New-Object –comobject Word.Application
    }

    $objWord.Visible = $true  
    $objDocument= $objWord.documents.open($rtfFile)
    $objDocument.SaveAs($pdfFileName.Value, 17)
    
    $objDocument.Close()
}

############## CSV ###############
function ConvertCsvToPDF ($csvFile, [ref] $pdfFileName)
{
    $pdfFile = Get-ChildItem $csvFile | % {$_.BaseName}
    $pdfFileName.Value = "$($workingFolder)\$($pdfFile).pdf"

    $objExcel.visible = $true
    $workbook = $objExcel.workbooks.open($csvFile, 3)
    $worksheet = $workbook.Worksheets.Item(1)
    # Auto-Sizing Columns / Rows
    $worksheet.UsedRange.Columns.Autofit() | Out-Null
    #PRINT SETUP TO FIT DATA TO ONE PAGE
    $worksheet.PageSetup.Zoom = $false 
    $worksheet.PageSetup.Orientation = $xlLandscape
    $worksheet.PageSetup.Draft = $false
    $worksheet.PageSetup.PaperSize = $xlPaperLegal    
    $worksheet.PageSetup.LeftMargin = $objExcel.InchesToPoints(0.25)
    $worksheet.PageSetup.RightMargin = $objExcel.InchesToPoints(0.25)
    $worksheet.PageSetup.TopMargin = $objExcel.InchesToPoints(0.25)
    $worksheet.PageSetup.BottomMargin = $objExcel.InchesToPoints(0.25)
    $worksheet.PageSetup.FitToPagesWide = 1
    $worksheet.PageSetup.FitToPagesTall = 9999

    $workbook.Saved = $true
    $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfFileName.Value, $xlQuality, $xlIncludeDocProperties, $xlIgnorePrintAreas)
    $objExcel.Workbooks.close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)

    return $pdfFileName
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
            }


        }  #if

    } #for

    Remove-Item -Path "$($batch.FullName)\error.txt"
}  
 

########################### MAIN ###########################
#Get count of files in directory $sidesErrorDir

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

    Add-Content $htmlFileName "<td width='45%' align='left'>$($errorMessages)</td>" 
   	Add-Content $htmlFileName "<td width='35%' align='left'>$($fixMessage)</td>" 
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
 

