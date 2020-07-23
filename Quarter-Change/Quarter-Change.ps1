$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$modulePath = "..\Scripts"
$importImpressionsScript = join-path $modulePath "Allow-Remote-Only-NLA.ps1"

$logPath = "..\Logs"
#log file settings
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title

$today = get-date -Format  "MMddyyyy"
$oldDate = "01/1/2020"
$newDate = "04/1/2020"
$oldQuarter = "201"
$newQuarter = "202"


#get list of workstations to reboot
$formPaths = Get-Content  ("quarterChangeAll.txt")

foreach ($formPath in $formPaths) 
{   
    $runtimeISL = join-path $formPath "runtime.isl"
    $destination = "$($runtimeISL)_$($today)"     #change this to before the dot
    Copy-Item -Path $runtimeISL -Destination  $destination -Force
    Write-Log $logfile "$($runtimeISL) copied to $($destination)"

    $islContent = Get-Content -Path $runtimeISL -Raw
    $islContentLength = $islContent.Length
    $startIndex = $islContent.IndexOf("date DE6_current_quarter")
    if ($startIndex -ge 0)
    {
        $islContentNoChange = $islContent.Substring(0, $startIndex - 1)
        $islContentChange = $islContent.Substring($startIndex)
        $islContentChange = $islContentChange.Replace($oldDate, $newDate)
        $islContentChange = $islContentChange.Replace($oldQuarter, $newQuarter)
        $islContentNoChange + $islContentChange | Set-Content -Path $runtimeISL
        Write-Log $logfile "$($runtimeISL) changed: $($islContentChange)"
    }
    else
    {
        Write-Log $logfile "No Change: $($runtimeISL)"
    }

 }


         