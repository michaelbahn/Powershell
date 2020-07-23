#Create ACK for RBE Batch with single SUBMIT file
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
Write-Host $title
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$rbeProblemBatch = $null
$now = get-date
$twentyMinutesOld = $now.AddMinutes(-20)    #threshold for old file

Import-Module (Join-Path $modulePath Utilities.psm1) -Force
Import-Module .\Run-SQL.psm1 -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content  recipients.txt

#read settings file
$configFile = "$($title).csv"
if (!(Test-Files $logFile, $configFile))    #check for settings file
{
    return
}
$rbeDirs = Import-Csv -Path  ($configFile)   

#loop through RBE folders specified in  settings file
foreach ($rbeDir in $rbeDirs)
{
    $tempFolder = Split-Path -path $rbeDir.RBEPath -parent
    $tempFolder = Join-Path $tempFolder  "RBE-Fix"   #temp folder to move problem batches

#look for old SUBMIT.SYS file    
    $oldSubmitFiles = Get-ChildItem -Path $rbeDir.RBEPath -File -Filter "SUBMIT.STS" -Recurse -ErrorAction SilentlyContinue # | Where-Object LastWriteTime -lt $twentyMinutesOld

    if ($oldSubmitFiles -eq $null)
    {
        Write-Log $logFile "$($rbeDir.RBEPath) passed old SUBMIT.STS check"
    }

#loop through folders that contain old SUBMIT.SYS file
    foreach ($oldSubmitFile in $oldSubmitFiles)
    {
        $batchPath = $oldSubmitFile.PSParentPath.Trim()
        $batchCount =  Get-ChildItem -Path $batchPath -File -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        if ($batchCount -eq 1 )     #only one file in batch
        {
                $batchName = Split-Path $batchPath -Leaf
                $reconFileName = join-path $rbeDir.ReconPath "$($batchName).RSL"

                if (!(Test-Path -Path $reconFileName)) 
                {
                    $message = "Problem with $batchPath`r`n" 
                    $message += "$($reconFileName) recon file does not exist in $($rbeDir.ReconPath)" 
                    Write-Log $logFile $message
                    $rbeProblemBatch += $message
                    $title += " problem"
                    break
                }

                #get batch count
                $reconFileData = Get-Content $reconFileName
                $recon = $reconFileData-Split [Environment]::NewLine
                $lineCount = $recon.Count - 1

                $dlnSearch = Get-DLN-Prefix ($recon)  #Get DLN to query database
                [int] $dlnCount = Get-DLN-Count ($recon[$lineCount])  #DLN count is in count line

                #verify in database
                #$queryCount  = Query-DLN-Count $rbeDir.SQLPath $rbeDir.Database $dlnSearch

#                if ($queryCount -ne  $dlnCount)
#                {
#                   $message = "$($reconFileName) recon batch file count does not match query count in $($rbeDir.SQLPath)`r`n" 
#                   $message +="Batch $($batchPath)`r`n" 
#                   $message +="Recon batch file count: $($dlnCount)`r`n" 
#                   $message += "Query count: $($queryCount)" 
#                    Write-Log $logFile  $message
#                    $rbeProblemBatch += $message
#                    $title += " problem"
#                    break
#                }

                #archive RBE folder and create ACK
                if (!(Test-Path $tempFolder))
                {
                    New-Item $tempFolder -ItemType Folder
                }

                 Move-Item $batchPath -destination $tempFolder   #move old batch
                 #Remove-Item $batchPath -Force  -Recurse #delete old batch

                #create new ACK File in RBE folder
                $newAckFileName = "$($batchName).ACK"

                if (!(Test-Path "$($rbeDir.RBEPath)\$($newAckFileName)"))
                {
                    try{
                        New-Item -path $rbeDir.RBEPath -ItemType File  -Name $newAckFileName  -ErrorAction Stop
                        $message = "ACK created: $(Join-Path $rbeDir.RBEPath $newAckFileName)`r`n"
                        Write-Log $logFile $message
                        # $rbeProblemBatch += $message
                    }
                    catch
                    {
                        $message = "Error creating ACK file $($newAckFileName) at $($rbeDir.RBEPath)`r`n"
                        Write-Log $logFile $message
                        $rbeProblemBatch += $message
                    }
                }
                else
                {
                    $message = "ACK file $($newAckFileName) already exists in $($rbeDir.RBEPath)`r`n"
                    Write-Log $logFile $message
                }

        }
        else
        {
                    $message = "$($batchPath) has old SUBMIT.STS but other files are present`r`n"
                    Write-Log $logFile $message
        }
    }
}
    
#if there were problem batches, send email
if ($rbeProblemBatch -ne $null)
{
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $title -Body $rbeProblemBatch 
}
else
{
    Write-Log $logFile "No RBE Folders with old SUBMIT.STS"
}
