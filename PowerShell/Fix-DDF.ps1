#Fix DDF for RBE Batch with multiple image types (PDF, TIF, JPG)
cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
Write-Host $title
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$settingsPath = Resolve-Path ("..\Config")
$modulePath = "..\Scripts"
$logPath = "..\Logs"
$now = get-date

Import-Module (Join-Path $modulePath Utilities.psm1) -Force

$logFile= Initialize-Log $logPath $title
$sender = Get-Content  (join-path $settingsPath Sender.txt)
$recipients = Get-Content  (join-path $settingsPath emailRBE.txt)

#read settings file
$configFile = join-path $settingsPath "$($title).txt"
if (!(Test-Files $logFile, $configFile))    #check for settings file
{
    return
}
$rbeDirs = Get-Content -Path  ($configFile)   
$ddfFixerExeDirectory = "D:\Program files\EDD\SSFFixerAutoRunExeFiles"
$ddfFixerExePath = Join-Path $ddfFixerExeDirectory "DMSFAX.DDFFixer.exe"
$ddfFixerInputPath = "D:\DDFFixer\Input"
$fixedBatch = $null

#loop through RBE folders specified in  settings file
foreach ($rbeDir in $rbeDirs)
{
    #look for error.sts file    
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
            New-Item $archiveFolder -ItemType Folder
        }
    }

    #loop through folders that contain error.sts file and see if they have multiple image types
    foreach ($errorFile in $errorFiles)
    {
        $batchPath = $errorFile.PSParentPath.Trim()
        $tifCount =  Get-ChildItem -Path $batchPath -File -Filter *.tif -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $jpgCount =  Get-ChildItem -Path $batchPath -File -Filter *.jpg -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $pdfCount =  Get-ChildItem -Path $batchPath -File -Filter *.pdf -Force -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count

        if (AtLeastTwo $tifCount $jpgCount $pdfCount)
        {
                $rbeServer = Split-Path $rbeDir -Parent
                $batchName = Split-Path $batchPath -Leaf
                $archiveBatchPath = Join-Path $archiveFolder $batchName
                if ((Test-Path -Path $archiveBatchPath)) 
                {
                    $message = "Fix-DDF was previously tried for $batchName`r`n`t" 
                    Write-Log $logFile $message
                    break
                }

                #move to input folder 
                 Move-Item $batchPath -destination $ddfFixerInputPath  

                #convert all images to TIF and update DDF
                Start-Process -FilePath $ddfFixerExePath -WorkingDirectory  $ddfFixerExeDirectory -Wait
                $fixedBatchPath = Join-Path $ddfFixerInputPath $batchName 
                Copy-Item -path $fixedBatchPath -destination $rbeDir -Recurse
                Move-Item -path $fixedBatchPath -destination $archiveFolder -Force
                $message = "DDF fixed: $(Join-Path $rbeDir $batchName)`t`r`n"
                Write-Log $logFile $message
                $fixedBatch += $message
            }   #if

        }   #errorfiles
    }   #rbe dirs
    
#if there were problem batches, send email
if ($fixedBatch -ne $null)
{
    Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $title -Body $fixedBatch 
}
else
{
    Write-Log $logFile "No RBE Folders with errors"
}
