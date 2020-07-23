cls
#Initialize settings
$now = get-date -format yyyy-MM-dd-HH-mm
$today = get-date -Format  d
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = Resolve-Path "..\Scripts"
$logPath = Resolve-Path "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force

#get list of servers
$servers = Get-Content "Server-List-All.txt"
$sender = "teamops@edd.ca.gov"
$recipients = Get-Content recipients.txt

$hotFixList  = @(0)
#loop thru each error check

foreach ($server in $servers) 
{   
    $session = New-PSSession $server
    $hotFixes = Invoke-Command -Session $session -ScriptBlock {Get-Hotfix} | Sort-Object -Descending -Property InstalledOn 
    $installedDate = $null

    foreach ($hotFix in $hotFixes) 
    {
        if (($hotFix.InstalledOn -eq $installedDate) -or ($installedDate -eq $null))
        {
            $hotFixList += $hotFix
            $installedDate = $hotFix.InstalledOn
        }        
        else
        {
            break
        }        

    }

    Remove-PSSession -Session $session
    Clear-Variable $session
}

$hotFixSorted = $hotFixList | Sort-Object InstalledOn, PSComputerName, HotFixID
$hotFixSorted | ConvertTo-Html  -Title "Last Windows Update Report" -Property InstalledOn, PSComputerName, HotFixID | Out-File "LastWindowsUpdateReport.htm" 

$MailBody = Get-Content "LastWindowsUpdateReport.htm"  -Raw
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject "Last Windows Update Report" -Body $MailBody -BodyAsHtml


