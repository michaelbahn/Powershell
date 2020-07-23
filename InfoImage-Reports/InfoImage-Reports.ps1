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
Import-Module .\InfoImage-Report-SQL.psm1 -Force

#email settings
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt
$recipientsTax = Get-Content  recipientsTax.txt
$subjectTax = "MIS $((get-date).DayOfWeek) Morning Counts"

$databaseProduction = "DGC1VDBRPTPD01\DGC1VDBRPTPD01"
$databaseTME = "DMGOVIRDBPD11\DMGOVIRDBPD11"
$pathInfoImageWebService = "\\DGVMTRPTPD01\InfoImageWebLog"
$tmeLogPathExtract = "\\dgvmtrptpd01\d$\TME\Logs\TaxMISExtractConsoleLogs"
$csvSettings  = "LogFiles.csv"
$LogFileSettingWorkset = @("Workset Update", "Total no. of files processed successfully", "\\dgvmtrptpd01\WorksetUpdate\log", "\\dgvmtrptpd01\d$\TME\Logs\WorkSetUpdateLog")

$htmlHeader = Get-Content htmlHeader.txt
$htmlTableHeader = "<table><tr><th></th><th>Production (DGC1VDBRPTPD01)</th><th>TME (DMGOVIRDBPD11)</th></tr>"
$htmlTableFooter = "</table>"
$htmlFooter = "</body></html>"

function Compare-LogFiles ([DateTime] $day)
{
    $dateFormatReport =  $day.ToString("M/d/yyyy")
    $dateFormatFileSuffix =  $day.ToString("yyyyMMdd")
    $htmlLogTable = @("<table><tr><th>Console Log file</th><th>$($dateFormatReport)</th><th>Production</th><th>TME</th></tr>")

    $logFileSettings = Import-Csv -Path $csvSettings    

    foreach ($logFileSetting in $logFileSettings)
    {
        $column1caption = $logFileSetting.column1
        $column2caption = $logFileSetting.column2
        $subSection = $logFileSetting.subsection
        $prodLogPath = $logFileSetting.prodLogPath
        $tmeLogPath = $logFileSetting.tmeLogPath
       
        switch ($column1caption)
        {
            "Tax Reports" 
            {
                $prodLogFileName = "$($prodLogPath)\DMS_TRCCApplication$($dateFormatFileSuffix).log"
                $tmeLogFileName = "$($tmeLogPath)\DMS_TRCCApplication$($dateFormatFileSuffix).log"
                if ((Test-Path $prodLogFileName) -and (Test-Path $tmeLogFileName))
                {
                    $prodLog = Get-Content $prodLogFileName -Raw
                    $tmeLog = Get-Content $tmeLogFileName -Raw
    
                    $prodSubSectionIndex = $prodLog.IndexOf($subSection)
                    $tmeSubSectionIndex = $tmeLog.IndexOf($subSection)
                    $subProdLog = $prodLog.Substring($prodSubSectionIndex)
                    $subTmeLog = $tmeLog.Substring($tmeSubSectionIndex)
                    $prodIndexValue = $subProdLog.IndexOf($column2caption)
                    $tmeIndexValue = $subTmeLog.IndexOf($column2caption)
                    $subProdLogValue = $subProdLog.Substring($prodIndexValue)
                    $subTmeLogValue = $subTmeLog.Substring($tmeIndexValue)
                    $indexProdLogEnd = $subProdLogValue.IndexOf("Log Entry")
                    #if you hit end of file, then stop
                    if ($indexProdLogEnd  -lt 0) 
                    { 
                        $indexProdLogEnd = $subProdLogvalue.Length - 1
                    }
                    $indexTmeLogEnd = $subTmeLogValue.IndexOf("Log Entry")
                    if ($indexTmeLogEnd  -lt 0) 
                    { 
                        $indexTmeLogEnd = $subTmeLogvalue.Length - 1
                    }

                    $prodLogValue = $subProdLogValue.Substring($column2caption.Length, $indexProdLogEnd - $column2caption.Length)
                    $tmeLogValue = $subTmeLogValue.Substring($column2caption.Length, $indexTmeLogEnd - $column2caption.Length)

                    #convert row to HTML
                    $column2caption = $column2caption.Replace("=", $subSection)
                    $htmlLogTable  +=  "<tr><td>$($column1caption)</td><td>$($column2caption)</td><td>$($prodLogValue)</td><td>$tmeLogValue</td></tr>`r`n"
                    break
                }  
            }  #tax reports

            "SA Console" 
            {
                $prodLogFileName = "$($prodLogPath)\SAConsoleAPI$($dateFormatFileSuffix).log"
                $tmeLogFileName = "$($tmeLogPath)\SAConsoleAPI_$($dateFormatFileSuffix).log"
                if ((Test-Path $prodLogFileName) -and (Test-Path $tmeLogFileName))
                {
                    $prodLog = Get-Content $prodLogFileName -Raw
                    $tmeLog = Get-Content $tmeLogFileName -Raw
    
                    $prodIndexValue = $prodLog.IndexOf($column2caption)
                    $tmeIndexValue = $tmeLog.IndexOf($column2caption)
                    $subProdLog = $prodLog.Substring($prodIndexValue)
                    $subTmeLog = $tmeLog.Substring($tmeIndexValue)
                    $indexProdLogEnd = $subProdLog.IndexOf((get-date).ToString("yyyy"))
                    $indexTmeLogEnd = $subTmeLog.IndexOf((get-date).ToString("yyyy"))
                    $prodLogValue = $subProdLog.Substring($column2caption.Length, $indexProdLogEnd - $column2caption.Length)
                    $tmeLogValue = $subTmeLog.Substring($column2caption.Length, $indexTmeLogEnd - $column2caption.Length)

                    #convert row to HTML
                    $column2caption = $column2caption.Replace(":", "")
                    $htmlLogTable  +=  "<tr><td>$($column1caption)</td><td>$($column2caption)</td><td>$($prodLogValue)</td><td>$tmeLogValue</td></tr>`r`n"
                    break
                }   #if
            } #SA Console
            "Workset Update" 
            {

                $prodLogFileName = "$($prodLogPath)\WorksetUpdate$($dateFormatFileSuffix).log"
                $tmeLogFileName = "$($tmeLogPath)\WorksetUpdateLog_$($dateFormatFileSuffix).log"

                if ((Test-Path $prodLogFileName) -and (Test-Path $tmeLogFileName))
                {
                    $tmeLog = Get-Content $tmeLogFileName -Raw
                    #missing space in TME log
                    if ($column2caption -eq "Total no. workitems processed in all worksets:")
                    {
                        $column2search = "Total no.workitems processed in all worksets:"
                    }
                    else
                    {
                        $column2search = $column2caption
                    }

                    $tmeIndexValue = $tmeLog.IndexOf($column2search)
                    $subTmeLog = $tmeLog.Substring($tmeIndexValue)
                    $indexTmeLogEnd = $subTmeLog.IndexOf((get-date).ToString("yyyy"))
                    $tmeLogValue = $subTmeLog.Substring($column2search.Length, $indexTmeLogEnd - $column2search.Length)
    
                    $prodLog = Get-Content $prodLogFileName -Raw
                    $prodLogInt = 0
                    for ($i=1; $i -le 2; $i++)
                    {
                        $prodIndexValue = $prodLog.IndexOf($column2caption)
                        $subProdLog = $prodLog.Substring($prodIndexValue)
                        $indexProdLogEnd = $subProdLog.IndexOf((get-date).ToString("yyyy"))
                        $prodLogValue = $subProdLog.Substring($column2caption.Length, $indexProdLogEnd - $column2caption.Length)
                        $prodLogInt +=   $prodLogValue -as [int]
                        $prodLog = $subProdLog.Substring($column2caption.Length)  #get next total
                    }
                    #convert row to HTML
                    $column2caption = $column2caption.Replace(":", "")
                    $htmlLogTable  +=  "<tr><td>$($column1caption)</td><td>$($column2caption)</td><td>$($prodLogInt)</td><td>$tmeLogValue</td></tr>`r`n"
                    break                
                }   #if
            } #Workset Update
        }  #switch
    }#for

    $htmlLogTable += "$($htmlTableFooter)`r`n" 
    return $htmlLogTable
}

#get days to check
$yesterday = (get-date).AddDays(-1)
$days = @()

#add extra days for week-end
switch ($yesterday.DayOfWeek)
{
    #"Monday" {$days += $yesterday.AddDays(-3); $days += $yesterday.AddDays(-2); $days += $yesterday.AddDays(-1); break}
    "Sunday" {$days += $yesterday.AddDays(-2); $days += $yesterday.AddDays(-1); break}
    "Saturday" {$days += $yesterday.AddDays(-1); break}
}

#default is to check yesterday
$days += $yesterday

#get table counts
$balance = $true
 
$resultProduction = Select-Counts  $databaseProduction $days
$resultTME = Select-Counts  $databaseTME $days

#iterate through rows and generate html 
$resultRowCount = [int] $resultProduction.Count
$reportRow = ""
$balance = $true
for ($i = 0; $i -lt $resultRowCount; $i++)
{
    #if any of the counts don't match we are are out of balance
    if ($resultProduction[$i].RecordCount -ne $resultTME[$i].RecordCount)
    {
        $balance = $false
    }

    #convert output row to HTML
    $reportRow +=  "<tr><td><b>$($resultProduction[$i].Caption)</b></td><td>$($resultProduction[$i].RecordCount)</td><td>$($resultTME[$i].RecordCount)</td></tr>`r`n"
}

if ($balance)
{
    $report =  "<p>The MIS counts balance this morning. </p>"
}
else
{
    $report =  "<p>The MIS counts do not balance this morning. </p>"
}
$report +=  "$($htmlHeader)`r`n$($htmlTableHeader)"
$report += $reportRow
$report += "$($htmlTableFooter)`r`n" 

#InfoImage web Service html
$pathInfoImageWebService = "\\DGVMTRPTPD01\d$\TME\Logs\InfoImageWebLog"
$htmlListHeader = "<ul style='list-style-type:disc;'><b>"
$htmlListFooter = "</ul></b>"
$report += "$($htmlListHeader)"
$htmlLogFiles = @()

#check InfoImage web Service

foreach ($day in $days)
{
    $dateFormatForReport =  $day.ToString("M/d/yy")
    $webServiceText = "InfoImage Web Service $($dateFormatForReport): "

    #open log file
    $dateFormattedForTableName =  $day.ToString("yyyyMMdd")
    $logfile = "$($pathInfoImageWebService)\SAConsoleAPI_$($dateFormattedForTableName).log"
    if (test-path ($logFile))
    {        
        $webServiceLog = Get-Content $logFile -Raw
        if ($webServiceLog.IndexOf("Logged in successfully") -ge 0)
        {
            if ($webServiceLog.IndexOf("Logged off Successfully") -ge 0)
            {
                $webServiceLink = "<a href=$($logFile)>Successful</a>"
                $htmlLogFiles += Compare-LogFiles $day
            }
            else
            {
                $webServiceLink = "<a href=$($logFile)>Missing Logoff in Log File</a>"
            }
        }
        else
        {
            $indexError = $webServiceLog.IndexOf("Error Message: ")
            if ($indexError -ge 0)
            {
                $indexErrorStart = $indexError+15 #move to start of error message
                $errorStart = $webServiceLog.Substring($indexErrorStart)
                $indexErrorEnd = $errorStart.IndexOf((get-date).ToString("yyyy"))
                $errorMessage = $errorStart.Substring(0, $indexErrorEnd)
                $webServiceLink = "<a href=$($logFile)>Error: $errorMessage</a>"
            }
            else
            {
                $webServiceLink = "<a href=$($logFile)>Missing Login in Log File</a>"
            }
        }
    }
    else
    {
        $webServiceLink = "<a href=$($pathInfoImageWebService)>Error: No Log File</a>"
    }

    $report += "<li>$($webServiceText)$($webServiceLink)</li>"

}

$report += "$($htmlListFooter)"

#send basic report to Tax group
$reportTax = $report + "$($htmlFooter)"
$subject = "INFOIMAGE REPORT FOR $((get-date).ToString("M/d/yy"))"
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipientsTax -Subject $subjectTax  -Body $report -BodyAsHtml

if ($htmlLogFiles -ne $null)
{
    $report += "$($htmlLogFiles)"
}

$report += "$($htmlFooter)"

$subject = "INFOIMAGE REPORT FOR $((get-date).ToString("M/d/yy"))"
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject  -Body $report -BodyAsHtml
