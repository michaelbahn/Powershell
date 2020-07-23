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
$errorMessage = "errorMessageSent"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

#returns true if at least two parameters are greater than zero
function SendToFixer($tifCount,  $jpegCount, $pdfCount) 
{
    $tifs = ($tifCount -gt 0)
    $jpegs = ($jpegCount -gt 0)
    $pdfs = ($pdfCount-gt 0)

    if ($tifs -or $jpegs) 
    {
        return $true;
    }
    else
    {
        return $false;
    }
}

#need to clean up Fixer input folder
Function CleanUpFolder ($inputPath)
{
   $files = Get-ChildItem -Path $rbeDir -File  -Recurse -ErrorAction SilentlyContinue
   foreach ($file in $files)
   {
        Remove-Item $file -Force 
   }

   $folders = Get-ChildItem -Path $rbeDir -File  -Recurse -ErrorAction SilentlyContinue

    Get-Item $inputPath
    Get-Item  "$($inputPath)\*.*"  | 
}

#use archive folder to track batch fix retries
Function GetArchivePath ($batchName, $archiveFolder, $archiveBatchPathError)
{
    $archiveBatchPath = Join-Path $archiveFolder $batchName    #first try
    $archiveBatchPathRetry = "$($archiveFolder)\retry_$($batchName)"   #second try
    
    #set retry if archive folder already exists
    if (Test-Path -Path $archiveBatchPath)
    {
        $archiveBatchPath = $archiveBatchPathRetry

    }
    #return error if retry already exists
    if (Test-Path -Path $archiveBatchPathRetry)
    {
        $archiveBatchPath = $archiveBatchPathError
    }

    #return null if retries exceeded    
    if (Test-Path -Path $archiveBatchPathError)
    {
        $archiveBatchPath = $null
    }

    Write-Log $logFile "Archive Batch Path: $($archiveBatchPath)"
    return $archiveBatchPath
}

Function Remove-ErrorFiles ($fixedBatchPath)
{
    Write-Log $logFile "Removing error files from $($fixedBatchPath)"
    $errorSts = Join-Path $fixedBatchPath error.sts 
    Get-Item  $errorSts | Remove-Item -Force 
    $tmpFiles = "$($fixedBatchPath)\*.tmp"
    Get-Item  $tmpFiles | Remove-Item -Force
}

Function FixDDF ($fixedBatchPath)
{
    $jpegFiles = Get-ChildItem -Path $fixedBatchPath -Filter *.jpeg  
    $jpgFiles = Get-ChildItem -Path $fixedBatchPath -Filter *.jpg  
    $ddfFile = Get-ChildItem -Path $fixedBatchPath -Filter *.ddf
    $ddfContent = Get-Content -Path $ddfFile.FullName -Raw
    if (($jpgFiles -eq $null) -and ($ddfContent.IndexOf("jpg") -ge 0))
    {
        Write-Log $logFile "Editing $($ddfContent) to remove .jpg"
        $ddfContent -replace 'jpg', 'pdf' | Set-Content -Path $ddfFile.FullName
    }

    if (($jpegFiles -eq $null) -and ($ddfContent.IndexOf("jpeg") -ge 0))
    {
        Write-Log $logFile "Editing $($ddfContent) to remove .jpeg"
        $ddfContent -replace 'jpeg', 'pdf' | Set-Content -Path $ddfFile.FullName
    }
    
    Remove-ErrorFiles $fixedBatchPath
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

    if ($errorFiles -eq $null)
    {
        Write-Log $logFile "$($rbeDir) passed error.sts check"
    }
    else   #get RBE leaf folder for archive
    {
        $rbeLeaf = Split-Path $rbeDir -Leaf
        $rbeParent = Split-Path -path $rbeDir -parent
        $archiveFolder = Join-Path $rbeParent "DDF-Fix"   
        if (!(Test-Path -Path $archiveFolder)) 
        {
            New-Item $archiveFolder -ItemType Directory
        }
    }

    #loop through folders that contain error.sts file and see if they have multiple image types
    foreach ($errorFile in $errorFiles)
    {
        $batchPath = Convert-Path $errorFile.PSParentPath.Trim()
        $tifCount =  Get-ChildItem -Path $batchPath -File -Filter *.tif -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $pdfCount =  Get-ChildItem -Path $batchPath -File -Filter *.pdf -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $jpgCount =  Get-ChildItem -Path $batchPath -File -Filter *.jpg -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $jpegCount =  Get-ChildItem -Path $batchPath -File -Filter *.jpeg -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $jpegCount +=  $jpgCount

        $batchName = Split-Path $batchPath -Leaf
        $archiveBatchPathError = "$($archiveFolder)\error_$($batchName)"

        #determine if batch has been run through fix script before
        $archiveBatchPath = GetArchivePath $batchName $archiveFolder $archiveBatchPathError 
        
        If ($archiveBatchPath -eq $null)  #don't send email               
        {
            Write-Log $logFile "Error message already sent: $($batchName)"
        }
        elseif ($archiveBatchPath -ne $archiveBatchPathError)
        {
            $fixedBatchPath = Join-Path $ddfFixerInputPath $batchName 
            if (SendToFixer $tifCount $jpegCount $pdfCount)
            {
                try
				{
					#first clean up input folder if any old files there.
                    CleanUpFolder $ddfFixerInputPath 

                    #move batch to input folder 
					Move-Item $batchPath -destination $ddfFixerInputPath  -Force -ErrorAction Stop
				
					#convert all images to TIF and update DDF
					$job = Start-Process -FilePath $ddfFixerExePath -WorkingDirectory  $ddfFixerExeDirectory -Wait -ErrorAction Stop
                    if ($job)
                    {
                        Wait-Job -Job $job 
                    }
					
                    #copy back to RBE folder
					Copy-Item -path $fixedBatchPath -destination $rbeDir -Recurse -ErrorAction Stop
					#move to archive folder 
					Move-Item -path $fixedBatchPath -destination $archiveFolder -Force -ErrorAction Stop
					$message = "DDF fixed: $(Join-Path $rbeDir $batchName)`t`r`n"
					Write-Log $logFile $message
					$fixedBatch += $message
                    CleanUpFolder $ddfFixerInputPath 
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
    		else  #Fix DDF               
		    {
                Write-Log $logFile "Retry Fix DDF: $($batchName)"
				#copy back to archive folder
				Copy-Item -path $batchPath -destination $archiveBatchPath -Recurse 
                FixDDF $batchPath
				$message = "DDF fix retry: $(Join-Path $rbeDir $batchName)`t`r`n"
				Write-Log $logFile $message
				$fixedBatch += $message
		    }   
        }   
		else  
		{
            #indicate error message has been emailed so it's not repeated
            New-Item $archiveBatchPath -ItemType Directory     #so won't send another email

            $message = "Batch $batchName at $($batchPath) is in Error, needs systems administrator attention. `t`r`n "
            $unfixedBatch += $message
            Write-Log $logFile $message
		}	#else $archiveBatchPath 
    }   #errorfiles
}   #rbe dirs
    
#if there were problem batches, send email
if ($unfixedBatch -ne $null)
{
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipientsAll -Subject "RBE Error Alert" -Body $unfixedBatch 
}

if ($fixedBatch -ne $null)
{
    #Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $title -Body $fixedBatch 
    Write-Log $logFile "RBE fixing completing"
}
else
{
    Write-Log $logFile "No RBE Folders with errors"
}
