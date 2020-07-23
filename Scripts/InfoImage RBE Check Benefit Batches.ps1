#Step 4 RBE Check Benefits
$title = "RBE Check Benefits"
cls
Write-Host $title
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = "..\Config"

$sender = gc  (join-path $settingsPath Sender.txt)
$recipients = gc  (join-path $settingsPath Recipients.txt)
$rbeDirs = gc  (join-path $settingsPath RBE-Dirs.txt) 

$rbeProblemBatch = $null
foreach ($rbeDir in $rbeDirs)
{
    $tempFolder = Split-Path -path $rbeDir -parent
    $tempFolder = Join-Path $rbeDir "RBETEMP"
    #$tempFolder = Join-Path $tempFolder (Split-Path -path $rbeDir -leaf)

    $rbeBatches =  Get-ChildItem -Path $rbeDir -Recurse -Directory -Force -ErrorAction SilentlyContinue
    foreach ($batch in $rbeBatches)
    {
        $batchPath = Join-Path $rbeDir $batch
        $batchCount =  Get-ChildItem -Path $batchPath -File -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty { $_.Count } 
        $batchFileName = Get-ChildItem -Path $batchPath -File -Force -ErrorAction SilentlyContinue -Filter "SUBMIT.STS" | Select-Object -ExpandProperty { $_.Name } 
        if ($batchCount -eq 1 )     #only one file in batch
        {
            if ($batchFileName -eq "SUBMIT.STS")  #SUBMIT.STS is only file in batch
            {
                $rbeProblemBatch += $batchPath
                Write-Host $batchPath
                if (Test-Path $tempFolder -eq $false)
                {
                    New-Item $tempFolder -ItemType Folder
                }
                Move-Item $batchPath -destination $tempFolder
                $newAckFile = Join-Path  $batch ".ack"
                Copy-Item -path $batchPath -ItemType File  -Destination $newAckFile -Filter "*.ACK"
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
