#Create ACK for RBE Batch with single SUBMIT file
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = Resolve-Path ("..\Config")
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$messageRBE = $null
$messageDDF = $null
$countRBE = 0
$countDDF = 0
$excelReportPath = "d:\Reports\DailyFix.xlsx"

# report data starts midnight for current date
$todayMidnight = Get-Date -Hour 0 -Minute 00 -Second 00
$today = get-date -UFormat "%A %b %e, %Y"
$subject = "Daily Report for Automated RBE Batch Fixes: $($today)"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  (join-path $settingsPath Sender.txt)
$recipients = Get-Content  (join-path $settingsPath emailRBE.txt)

#read Fix-RBE settings file
$configFileDDF = join-path $settingsPath "Fix-DDF.txt"
$configFileRBE = join-path $settingsPath "Fix-RBE.csv"
if (!(Test-Files $logFile, $configFileDDF, $configFileRBE))    #check for settings file
{
    return
}
$ddfDirs = Get-Content -Path  ($configFileDDF)   
$rbeDirs = Import-Csv -Path  ($configFileRBE)   

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$xl.DisplayAlerts = $false
$workbook = $Excel.Workbooks.Open()
$worksheet = $workbook.sheets.item(0)
    
#loop through RBE folders specified in  settings file
foreach ($ddfDir in $ddfDirs)
{
    $ddfRootFolder = Split-Path -path $ddfDir -parent
    $ddfFixFolderPath = Join-Path $ddfRootFolder "DDF-Fix"   #archive folder for Fix-DDF

#get all batches fixed by Fix-DDF
    $datessDDF = Get-ChildItem -Path $ddfFixFolderPath -Directory -ErrorAction SilentlyContinue | Group-Object {$_.LastWriteTime.Date} | Select-Object Count, Name

    if ($datessDDF -eq $null)
    {
        Write-Log $logFile "No Fix-DDF batches at $($ddfFixFolderPath)"
    }

#loop through dates 
    foreach ($dateDDF in $datesDDF)
    {

        $match = $false
        $endOfSheet = $false
        for ($row = 2, !($endOfSheet -or $match), $row++)
        {
            $reportDate = $WorkSheet.Range("A$($row)").Text


        }


        $countDDF += 1
        $batchName = "$($ddfDir)`t$($batchFolderDDF.Name)"
        $messageDDF += "$($batchName)`t`r`n"
        Write-Log $logFile "Batch #$($countDDF):`t$($batchName)"
    }
}

#loop through RBE folders specified in  settings file
foreach ($rbeDir in $rbeDirs)
{
    $rbeRootFolder = Split-Path -path $rbeDir.RBEPath -parent
    $rbeFixFolderPath = Join-Path $rbeRootFolder "RBETEMP2"   #archive folder for Fix-RBE

#get all batches fixed by Fix-RBE   
    $batchFoldersRBE = Get-ChildItem -Path $rbeFixFolderPath -Directory -ErrorAction SilentlyContinue | Group-Object {$_.LastWriteTime.Date}

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
    }
}
    
#if there were problem batches, send email
$message = ""
if (($countDDF -ge 1) -or ($countRBE -ge 1l))
{
    if ($countDDF -ge 1)
    {
        if ($countDDF -gt 1)
            {$plural = "es"}
        else
            {$plural = ""}

        $message += "Fix-DDF (multiple image types): $($countDDF) batch$($plural) fixed during $($today):`t`r`n`t`r`n$($messageDDF)`r`n"
    }

    if ($countRBE -ge 1)
    {
        if ($countRBE -gt 1)
            {$plural = "es"}
        else
            {$plural = ""}

        $message += "Fix-RBE (create ACK file): $($countRBE) batch$($plural) fixed during $($today):`t`r`n`t`r`n$($messageRBE)`r`n"
    }


    

#    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $message
}
else
{
    Write-Log $logFile "Nothing created today"
}
