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
$excelReportPath = "d:\Reports\DailyFix.xlsx"
$excelReportToday = "d:\Reports\DailyFix$(get-date -Format "yyyyMMdd").xlsx"

# report data starts midnight for current date
$todayMidnight = Get-Date -Hour 0 -Minute 00 -Second 00
$today = get-date -UFormat "%A %b %e, %Y"
$subject = "Daily Report for Automated RBE Batch Fixes: $($today)"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  Sender.txt
$recipients = Get-Content  emailRecipients.txt

#read Fix-RBE settings file
$configFileDDF = join-path $settingsPath "Fix-DDF.txt"
$configFileRBE = join-path $settingsPath "Fix-RBE.csv"
if (!(Test-Files $logFile, $configFileDDF, $configFileRBE))    #check for settings file
{
    return
}
$ddfDirs = $configFileDDF
$rbeDirs = $configFileRBE

#open Daily Report xlsx
$xl = New-Object -COM "Excel.Application"
$xl.Visible = $false
$xl.DisplayAlerts = $false
$workbook = $xl.Workbooks.Open($excelReportPath)
$worksheet = $workbook.sheets.item("Data")

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
        $messageDDF += "$($batchName)`t`r`n"
        Write-Log $logFile "Batch #$($countDDF):`t$($batchName)"
        #Move-Item $batchFolderDDF $ddfFixArchivePath
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
        $messageRBE += "$($batchName)`t`r`n"
        Write-Log $logFile "Batch #$($countRBE):`t$($batchName)"
        #Move-Item $batchFolderRBE $rbeArchiveFolderPath
    }
}
    
#if there were problem batches, send email
$message = ""
if (($countDDF -ge 1) -or ($countRBE -ge 1l))
{
    $worksheet.Range("A$($row)").NumberFormat = "M/dd/yyyy"
    $worksheet.Range("A$($row)").Value = get-date -Format "M/dd/yyyy"
    $worksheet.Range("B$($row)").NumberFormat = "0"
    $worksheet.Range("B$($row)").Value = $countRBE.ToString()
    $worksheet.Range("C$($row)").NumberFormat = "0"
    $worksheet.Range("C$($row)").Value = $countDDF.ToString()
    $worksheet.Range("D$($row)").NumberFormat = "0"
    $worksheet.Range("D$($row)").Value = ($countDDF + $countRBE).ToString()

    #$worksheet.ChartObjects(1).Activate 
    #$chart =  $worksheet.ChartObjects(1).Chart

    $workbook.SaveAs($excelReportPath)
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

        $message += "Fix-RBE (create ACK file): $($countRBE) batch$($plural) fixed during $($today):`t`r`n`t`r`n$($messageRBE)`r`n"
    }

    if ($countDDF -ge 1)
    {
        if ($countDDF -gt 1)
            {$plural = "es"}
        else
            {$plural = ""}

        $message += "Fix-DDF (multiple image types): $($countDDF) batch$($plural) fixed during $($today):`t`r`n`t`r`n$($messageDDF)`r`n"
    }



    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $message -Attachments $excelReportPath
}
else
{
    Write-Log $logFile "Nothing created today"
}
