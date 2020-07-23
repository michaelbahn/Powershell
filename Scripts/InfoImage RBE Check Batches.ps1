#Step 4 RBE Check Benefits
$title = "RBE Check Benefits"
cls
Write-Host $title
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$title = "RBE-Check-Batches"

Import-Module (Join-Path $modulePath Write-Log.psm1) -Force
Import-Module (Join-Path $modulePath Run-SQL.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  (join-path $settingsPath Sender.txt)
$recipients = Get-Content  (join-path $settingsPath Recipients.txt)
$rbeDirs = Import-Csv -Path  (join-path $settingsPath RBE-Dirs-Prod.csv) 

$now = get-date
$thirtyMinutesOld = $now.AddMinutes(-30)

foreach ($rbeDir in $rbeDirs)
{
    $tempFolder = Split-Path -path $rbeDir.RBEPath -parent
    $tempFolder = Join-Path $tempFolder  "RBETEMP"   #temp folder to move problem batches
    
    $rbeBatches = Get-ChildItem -Path $rbeDir.RBEPath   -Directory -ErrorAction SilentlyContinue | Where-Object LastWriteTime -lt $thirtyMinutesOld
    foreach ($batch in $rbeBatches)
    {
        $batchPath = Join-Path $rbeDir.RBEPath $batch.Name
        $batchCount =  Get-ChildItem -Path $batchPath -File -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $batchFileName = Get-ChildItem -Path $batchPath -File -Force -ErrorAction SilentlyContinue -Filter "SUBMIT.STS" | Select-Object -ExpandProperty Name
        if ($batchCount -eq 1 )     #only one file in batch
        {
            if ($batchFileName -eq "SUBMIT.STS")  #SUBMIT.STS is only file in batch
            {
                Write-Log $logFile $batchPath 
                $reconFileName = Join-Path $rbeDir.ReconPath.Trim() ($batch.Name + ".RSL")

                if (!(Test-Path -Path $reconFileName)) 
                {
                    Write-Log $logFile "$($reconFileName) recon file does not exist in $($rbeDir.ReconPath)" 
                    return
                }

                #get batch count
                $reconFileData = Get-Content $reconFileName
                $recon = $reconFileData-Split [Environment]::NewLine
                $lineCount = $recon.Count - 1
                $dlnFirst = $recon[0].Substring( $recon[0].Length -16)
                $dlnSearch = $dlnFirst.Substring(0, 13)
                $reconLast = $recon[$lineCount]
                [int] $dlnCount = $reconLast.Substring(4, 3)
                Write-Log $logFile $reconLast
                Write-Log $logFile "Count: $($recon.Count - 1)"

                #verify in database
                #$queryCount  = Run-SQL ($rbeDir.SQLPath, $dlnSearch)
                $queryCount  = $dlnCount

                if ($queryCount -ne  $dlnCount)
                {
                    Write-Log $logFile  "$($reconFileName) recon batch file count does not match query count in $($rbeDir.SQLPath)" 
                    Write-Log $logFile  "recon batch file count: $($dlnCount)" 
                    Write-Log $logFile  "query count: $($queryCount)" 

                    return
                }

                #archive RBE folder and create ACK
                if (!(Test-Path $tempFolder))
                {
                    New-Item $tempFolder -ItemType Folder
                }

                Move-Item $batchPath -destination $tempFolder
                
                #create new ACK File in RBE folder
                $newAckFileName = $batch.Name + ".ACK"
                New-Item -path $rbeDir.RBEPath -ItemType File  -Name $newAckFileName 
                Write-Log $logFile "ACK crerated: $(Join-Path $rbeDir.RBEPath $newAckFileName)"
            }
        }

    }
}
    


if ($rbeProblemBatch -ne $null)
{
    $rbeProblemBatch | ConvertTo-Html  -Title $title -Property FullName | Out-File $outputFile
    $mailBody = Get-Content $outputFile  -Raw
    Send-MailMessage -SmtpServer "smtp.edd.ca.gov" -From $sender -To $recipients -Subject $title -Body $mailBody -BodyAsHtml
}
else
{
    Write-Host "Check 8 passed: no RBE Folders with SUBMIT.STS"
}
