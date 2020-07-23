#Fix DDF for RBE Batch with multiple image types (PDF, TIF, JPG)
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
Write-Host $title
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$now = get-date

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

Function CheckDDFisOK ($fixedBatchPath)
{
    $jpegFiles = Get-ChildItem -Path $fixedBatchPath -Filter *.jpeg  
    $jpgFiles = Get-ChildItem -Path $fixedBatchPath -Filter *.jpg  
    $ddfFile = Get-ChildItem -Path $fixedBatchPath -Filter *.ddf
    $ddfContent = Get-Content -Path $ddfFile.FullName -Raw
    if (($jpgFiles -eq $null) -and ($ddfContent.IndexOf("jpg") -ge 0))
    {
        Copy-Item ddfFile.FullName 
        $ddfContent.Replace("jpg", "pdf")
        Set-Content -Path $ddfFile.FullName -Value $ddfContent -Force
    }

    if (($jpegFiles -eq $null) -and ($ddfContent.IndexOf("jpeg") -ge 0))
    {
        $ddfContent.Replace("jpeg", "pdf")
        Set-Content -Path $ddfFile.FullName -Value $ddfContent -Force
    }

}

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  sender.txt
$recipients = Get-Content recipients.txt
$recipientsAll = Get-Content recipients-all.txt

#read settings file
$rbeDirs = Get-Content -Path  "$($title).txt"
$ddfFixerExeDirectory = "D:\Program files\EDD\SSFFixerAutoRunExeFiles"
$ddfFixerExePath = Join-Path $ddfFixerExeDirectory "DMSFAX.DDFFixer.exe"
$ddfFixerInputPath = "D:\DDFFixer\Input"
$fixedBatch = $null
$unfixedBatch = $null

#loop through RBE folders specified in  settings file
foreach ($rbeDir in $rbeDirs)
{
    #look for error.sts file    
    $rbeServer = Split-Path $rbeDir -Parent
    $errorFiles = Get-ChildItem -Path $rbeDir -File -Filter "error.sts" -Recurse -ErrorAction SilentlyContinue

    if ($errorFiles -ne $null)
    {
         #loop through folders that contain error.sts file and see if they have multiple image types
        foreach ($errorFile in $errorFiles)
        {
            $batchPath = Convert-Path $errorFile.PSParentPath.Trim()
            $rbeLeaf = Split-Path $rbeDir -Leaf
            $rbeParent = Split-Path -path $rbeDir -parent
            $archiveFolder = Join-Path $rbeParent "DDF-Fix"   
            if (!(Test-Path -Path $archiveFolder)) 
            {
                New-Item $archiveFolder -ItemType Directory
            }
    
            $tifCount =  Get-ChildItem -Path $batchPath -File -Filter *.tif -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
            $pdfCount =  Get-ChildItem -Path $batchPath -File -Filter *.pdf -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
            $jpgCount =  Get-ChildItem -Path $batchPath -File -Filter *.jpg -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
            $jpegCount =  Get-ChildItem -Path $batchPath -File -Filter *.jpeg -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
            $jpegCount +=  $jpgCount

            #verify each pdf can be opened
            if (($tifCount -eq 0) -and ($jpegCount -eq 0) -and ($pdfCount -gt 0))
            {
                foreach ($pdfFile in $pdfFiles)
                {
                    $process = Start-Process $pdfFile.FullName 
                    Stop-Process -Id $process.Id -Force
                }
            }

            $batchName = Split-Path $batchPath -Leaf
            $archiveBatchPath = Join-Path $archiveFolder $batchName


            if ((Test-Path -Path $archiveBatchPath)) 
            {
                $message = "Fix-DDF was previously tried for $batchPath `r`n`t" 
                Write-Log $logFile $message
            }
         
            else
            {
                if (!(Test-Path "$($ddfFixerInputPath)\$($batchPath)"))
                {             
                    try
                    {
                        #move to input folder 
                        Move-Item $batchPath -destination $ddfFixerInputPath  -Force -ErrorAction Stop
                    

                        #convert all images to TIF and update DDF
                        $job = Start-Process -FilePath $ddfFixerExePath -WorkingDirectory  $ddfFixerExeDirectory -Wait -ErrorAction Stop
                        $fixedBatchPath = Join-Path $ddfFixerInputPath $batchName 
                        CheckDDFisOK $fixedBatchPath

                        #copy back to RBE folder
                        Copy-Item -path $fixedBatchPath -destination $rbeDir -Recurse -ErrorAction Stop
                        #move to archive folder 
                        Move-Item -path $fixedBatchPath -destination $archiveFolder -Force -ErrorAction Stop
                        $message = "DDF fixed: $(Join-Path $rbeDir $batchName)`t`r`n"
                        Write-Log $logFile $message
                        $fixedBatch += $message
                    }
                    catch 
                    {
                        $errorObject = $_
                        $message = "Error fixing DDF: $(Join-Path $rbeDir $batchName)`t`r`n$($errorObject)"
                        Write-Log $logFile $message
                        $fixedBatch += $message
                        Write-Log $logFile "DDF Fixer: $($job.ExitCode) $($job.Description)"
                    }

                }   
            }   
        }   #foreach ($errorFile in $errorFiles)
      }
     else   #no errors
    {
        Write-Log $logFile "$($rbeDir) passed error.sts check"
    }
   
    }   #rbe dirs
    
#if there were problem batches, send email
if ($unfixedBatch -ne $null)
{
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipientsAll -Subject "RBE Error Alert" -Body $unfixedBatch 
}

if ($fixedBatch -ne $null)
{
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $title -Body $fixedBatch 
}
else
{
    Write-Log $logFile "No RBE Folders with errors"
}
