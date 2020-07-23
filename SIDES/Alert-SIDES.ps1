cls
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$sidesInputDir = "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\Input\"
$sidesErrorDir = "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\ZipWrap-Error"
$sidesErrorBackup = "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\Input-error"

$nonftpfile = ""

$htmlFileName =  "SIDES.htm"
#Import-Module .\Alert-SEFT-functions.psm1 -Force
Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = "teaminf@edd.ca.gov"
$recipients = "teaminf@edd.ca.gov"
#$recipients = "michael.bahn@edd.ca.gov"
$workingFolder = "d:\temp"

New-Item -itemtype file -Path $htmlFileName -Force
$date = Get-Date


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
    add-content $htmlFileName  "<td width='75%' height='25' align='center'>" 
    add-content $htmlFileName  "<font face='tahoma' color='#003399' size='4'><strong><a href='\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\ZipWrap-Error'>SIDES ZipWrap Errors </strong></font>" 
    add-content $htmlFileName  "</td>" 
    add-content $htmlFileName  "</tr>" 
    #add-content $htmlFileName  "</table>" 

    #Add-Content $htmlFileName "<tr bgcolor='#D0E6FF'>" 
    #Add-Content $htmlFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>  </strong></font></td>" 
    #Add-Content $htmlFileName "</tr>" 
   
 }  
    
 # Function to write the HTML Header to the file  
 Function writeTableHeader  
 {  
 param($htmlFileName)  
    
 Add-Content $htmlFileName "<tr bgcolor=#D0E6FF>" 
 Add-Content $htmlFileName "<td width='25%' align='center'><strong>Batch</strong></td>" 
 Add-Content $htmlFileName "<td width='75%' align='center'><strong>Error</strong></td>" 
 Add-Content $htmlFileName "</tr>" 
 }  
    
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
    

function ConvertRtfToPDF ($csvFile)
{
    return $pdfFile
}

function ConvertCsvToPDF ($csvFile)
{
    $xlLandscape = 2
    $xlPaperLegal = 5

    $pdfFile = gci $csvFile | % {$_.BaseName}
    $pdfFile = "$($workingFolder)\$($pdfFile).pdf"
    $xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]
    $objExcel.visible = $true
    $workbook = $objExcel.workbooks.open($csvFile, 3)
    $worksheet = $workbook.Worksheets.Item(1)
    # Auto-Sizing Columns / Rows
    $worksheet.UsedRange.Columns.Autofit() | Out-Null
    #PRINT SETUP TO FIT DATA TO ONE PAGE
    $worksheet.PageSetup.Orientation = $xlLandscape
    $worksheet.PageSetup.Draft = $false
    $worksheet.PageSetup.PaperSize = $xlPaperLegal    
    $worksheet.PageSetup.LeftMargin = $objExcel.InchesToPoints(0.25)
    $worksheet.PageSetup.RightMargin = $objExcel.InchesToPoints(0.25)
    $worksheet.PageSetup.FitToPagesWide = 1
    $worksheet.PageSetup.FitToPagesTall = 9999

    $workbook.Saved = $true
    $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfFile)
    $objExcel.Workbooks.close()
    return $pdfFile
}

function FixBatch ($batch) 
{  

    Copy-Item -Path $batch.FullName -Destination $workingFolder -Force -Recurse
    $batchWorkingFolder = Split-Path $batch.FullName -leaf
    $batchWorkingPath = Join-Path $workingFolder $batchWorkingFolder
    $errors = Get-Content -Path "$($batchWorkingPath)\error.txt"

    foreach ($errorMessage in $errors)
    {
        $endOfInvalidFileTypeMessage = "is an invalid file types"
        $zipWrapErrorString = "ZipWrap-Error\"

        $indexEndOfInvalidFileTypeMessage = $errorMessage.IndexOf($endOfInvalidFileTypeMessage)
        if ($indexEndOfInvalidFileTypeMessage -ge 0)
        {
            $errorText = $errorMessage.Substring(0,$indexEndOfInvalidFileTypeMessage - 1)

            $indexZipWrapError = $errorText.IndexOf($zipWrapErrorString)
            $completePath = $errorText.SubString($indexZipWrapError+ $zipWrapErrorString.Length)
            $files = $completePath.Split("\")
            $doneFolder = join-path $sidesErrorDir $files[0] 
            $zipBatch = Join-Path $batchWorkingPath "$($files[1]).zip"
            $tempWorkingFolder = Join-Path $workingFolder $files[1]
            Expand-Archive -LiteralPath $zipBatch -DestinationPath $tempWorkingFolder

            $csvFileIndex = $errorText.IndexOf(".csv")
            $rtfFileIndex = $errorText.IndexOf(".rtf")

            if ($csvFileIndex -ge 0)
            {
                $csvFileName = "$($errorText.Substring(0,$csvFileIndex - 1)).csv"
                $csvFile = Join-Path $tempWorkingFolder $files[2]
                $pdfFile = ConvertCsvToPDF $csvFile

            }
            elseif ($rtfFileIndex -ge 0)
            {
                $rtfFileName = "$($errorText.Substring(0,$rtfFileIndex - 1)).rtf"
                $rtfFile = Join-Path $tempFolder $files[2]
                $pdfFile = ConvertRtfToPDF $rtfFile
            }
            else
            {
                $pdfFile = $null
            }

            if (Test-Path $pdfFile)
            {
                Set-Location $batch.FullName
                $zipBatchPath = join-path $batch.FullName $zipBatchPath
              
                Remove-Item -Path $csvFile -Force 
                Move-Item -Path $pdfFile -Destination $tempWorkingFolder -Force 
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

if($sidesErrorCount -le 0)
{
	Write-Log $logFile "SIDES Dir count is $($sidesErrorCount)"
    return
}
else
{
    if($sidesErrorCount -eq 1)
    {
        $logMessage = "ZipWrap Error: $($sidesErrorCount) SIDES batch in ZipWrap-Error folder." 
    }
    else
	{
        $logMessage = "ZipWrap Errors: $($sidesErrorCount) SIDES batches in ZipWrap-Error folder." 
    }
    
    Write-Log $logFile $logMessage
    writehtmlheader $htmlFileName
    writeTableHeader $htmlFileName  	
   $objExcel = New-Object -ComObject excel.application
   Set-Location $workingFolder
    Remove-Item "$($workingFolder)\*.*"  -Force -Recurse


	foreach ($batch in $sidesErrorFiles)
    {
        Add-Content $htmlFileName "<tr>" 
 	    Add-Content $htmlFileName "<td width='25%' align='left'><a href='file:$($batch.FullPath)'>$($batch.Name)</a></td>" 
        $errors = Get-Content -Path "$($batch.FullName)\error.txt"
        $errorMessages = ""
        foreach ($errorMessage in $errors)
        {
            $errorMessages += "$($errorMessage) <br>"
        }

   	    Add-Content $htmlFileName "<td width='75%' align='left'>$($errorMessages)</td>" 
 	    Add-Content $htmlFileName "</tr>" 
        Copy-Item $batch.FullName $sidesErrorBackup -Recurse
        Write-Log $logFile "Fix Batch $($batch.FullName)"
        FixBatch $batch
        Move-Item $batch.FullName $sidesInputDir  -Force 
        if (!(Test-Path "$($sidesInputDir)\$($batch.FullName)"))
        {
            Write-Log $logFile "Failed moving fixed Batch $($batch.FullName) to $($sidesInputDir)"
        }
    }
    objExcel.Quit
    writeHtmlFooter  $htmlFileName

    $mailBody = Get-Content $htmlFileName  -Raw
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $logMessage  -Body $mailBody -BodyAsHtml
}
 

