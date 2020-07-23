#Create ACK for RBE Batch with single SUBMIT file
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$messageRBE = $null
$messageDDF = $null
$countRBE = 0
$countDDF = 0
$excelReportPath = "d:\Reports"
$excelReportFile = Join-Path $excelReportPath "DailyFix.xlsx"
$excelReportToday = Join-Path $excelReportPath "DailyFix$(get-date -Format "yyyyMMdd").xlsx"
#$yesterday = get-date.AddDays(-1) -Format "yyyyMMdd"
#$excelReportToday = Join-Path $excelReportPath "DailyFix$($yesterday).xlsx"
  
$excelChart = Join-Path $excelReportPath "chart.jpg"
if (test-path ($excelChart)) {Remove-Item $excelChart -Force}
$chartOutputType = "JPG"

# report data starts midnight for current date
$todayMidnight = Get-Date -Hour 0 -Minute 00 -Second 00
$today = get-date -UFormat "%A %b %e, %Y"
$subject = "Daily Report for Automated RBE Batch Fixes: $($today)"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  (join-path $scriptPath sender.txt)
$recipients = Get-Content  (join-path $scriptPath recipients.txt)

#read Fix-RBE settings file
$configFileDDF = join-path $scriptPath "Fix-DDF.txt"
$configFileRBE = join-path $scriptPath "Fix-RBE.csv"
if (!(Test-Files $logFile, $configFileDDF, $configFileRBE))    #check for settings file
{
    return
}
$ddfDirs = Get-Content -Path  ($configFileDDF)   
$rbeDirs = Import-Csv -Path  ($configFileRBE)   

#open Daily Report xlsx
#to run from task scheduler you need to create C:\Windows\SysWOW64\config\systemprofile\desktop
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $false
$xl.DisplayAlerts = $false
#Write-Log $logFile "Opened Com Object: $($xl.Name)"
$workbook = $xl.Workbooks.Open($excelReportFile)
$worksheet = $workbook.sheets.item("Data")
#Write-Log $logFile "Opened excel $($workbook.Name) sheet $($worksheet.Name)"

#go to row for new date 
$endOfSheet = $false
$row = 2
while (!($endOfSheet))
{
    $reportDate = $WorkSheet.Range("A$($row)").Text
    if  ([string]$reportDate -as [DateTime])  
    {
        $row++
    }
    else
    {
        $endOfSheet = $true
    }
}

#loop through RBE folders specified in  settings file
foreach ($ddfDir in $ddfDirs)
{
    $ddfRootFolder = Split-Path -path $ddfDir -parent
    $ddfFixFolderPath = Join-Path $ddfRootFolder "DDF-Fix"   #folder for Fix-DDF
    $ddfFixArchivePath = Join-Path $ddfRootFolder "DDF-Fix-Archive"   #archive folder for Fix-DDF

#get all batches fixed by Fix-DDF
    $batchFoldersDDF = Get-ChildItem -Path $ddfFixFolderPath -Directory -ErrorAction SilentlyContinue

    if ($batchFoldersDDF -eq $null)
    {
        Write-Log $logFile "No Fix-DDF batches at $($ddfFixFolderPath)"
    }

#loop through folders 
    foreach ($batchFolderDDF in $batchFoldersDDF)
    {
        $countDDF += 1
        $batchName = "$($ddfDir)`t$($batchFolderDDF.Name)"
        $messageDDF += "<li>$($batchName)</li>`r`n"
        Write-Log $logFile "Batch #$($countDDF):`t$($batchName)"
        Move-OrCopyItem $batchFolderDDF.FullName $ddfFixArchivePath
    }
}

#loop through RBE folders specified in  settings file
foreach ($rbeDir in $rbeDirs)
{
    $rbeRootFolder = Split-Path -path $rbeDir.RBEPath -parent
    $rbeFixFolderPath = Join-Path $rbeRootFolder "RBE-Fix"   #folder for Fix-RBE
    $rbeArchiveFolderPath = Join-Path $rbeRootFolder "RBE-Fix-Archive"   #archive folder for Fix-RBE

#get all batches fixed by Fix-RBE   
    $batchFoldersRBE = Get-ChildItem -Path $rbeFixFolderPath -Directory -ErrorAction SilentlyContinue 

    if ($batchFoldersRBE -eq $null)
    {
        Write-Log $logFile "nothing processed at $($rbeFixFolderPath)"
    }

#loop through folders 
    foreach ($batchFolderRBE in $batchFoldersRBE)
    {
        $countRBE += 1
        $batchName = "$($rbeDir.RBEpath)`t$($batchFolderRBE.Name)"
        $messageRBE += "<li>$($batchName)</li>`r`n"
        Write-Log $logFile "Batch #$($countRBE):`t$($batchName)"
        Move-OrCopyItem $batchFolderRBE.FullName $rbeArchiveFolderPath
    }
}
    
#if there were problem batches, send email
if (($countDDF -ge 1) -or ($countRBE -ge 1l))
{
    $message = "<!DOCTYPE html><html><body>`r`n"
    $worksheet.Range("A$($row)").NumberFormat = "M/dd/yy"
    $worksheet.Range("A$($row)").Value = get-date -Format "M/dd/yy"
    $worksheet.Range("B$($row)").NumberFormat = "0"
    $worksheet.Range("B$($row)").Value = $countRBE.ToString()
    $worksheet.Range("C$($row)").NumberFormat = "0"
    $worksheet.Range("C$($row)").Value = $countDDF.ToString()
    $worksheet.Range("D$($row)").NumberFormat = "0"
    $worksheet.Range("D$($row)").Value = ($countDDF + $countRBE).ToString()
    $range = $worksheet.Range("A1:D$($row)") 
    Write-Log $logFile "Export Chart:`t$($excelChart)"

    #$chart =  $worksheet.ChartObjects(1).Chart    
    foreach ($chart in $worksheet.ChartObjects([System.Type]::Missing))  
        {
            $chart.Activate 
            #$chart.setSourceData($range)
            $chart.Chart.Export($excelChart, $chartOutputType, $false)
            Write-Log $logFile "Export Chart:`t$($excelChart)"
        }    

    $return = $workbook.SaveAs($excelReportFile)
    $workbook.SaveAs($excelReportToday)
    $workbook.Close
    $xl.Quit
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)


    if ($countRBE -ge 1)
    {
        if ($countRBE -gt 1)
            {$plural = "es"}
        else
            {$plural = ""}

        $message += "<h2>Fix-RBE (create ACK file): $($countRBE) batch$($plural) fixed during $($today)`r`n</h2><ul>$($messageRBE)</ul>`r`n"
    }

    if ($countDDF -ge 1)
    {
        if ($countDDF -gt 1)
            {$plural = "es"}
        else
            {$plural = ""}

        $message += "<h2>Fix-DDF (multiple image types): $($countDDF) batch$($plural) fixed during $($today)`r`n</h2><ul>$($messageDDF)</ul>`r`n"
    }

    #$message += "<img src=d:\reports\chart.jpg></body></html>"
    $message += "<img src=cid:chart.jpg></body></html>"

    Write-Log $logFile "emailing $($recipients)"
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $message -BodyAsHtml -Attachments $excelChart
    Start-Sleep -s 60
    Stop-Process -Name EXCEL -Force
}
else
{
    Write-Log $logFile "Nothing created today"
}

