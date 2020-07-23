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
$serverXLS = Join-Path $logPath "QA-Server.xlsx"

Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$events = @()


#get list of servers
$servers = Get-Content "servers.txt"
$users = Get-Content "users.txt"

$date = (Get-Date).AddDays(-4)
$properties = @(	'TimeCreated',	@{n='Server';e={$server}}, @{n='UserID';e={$_.properties[5].value}} )
$events = @()
#loop thru each server
foreach ($server in $servers) 
{   
    
    foreach ($user in $users) 
    {   
        $winEvent = Get-WinEvent -ComputerName $server  -MaxEvents 1 -FilterHashtable @{ LogName='Security'; StartTime=$date; Id='4624' } | Select-Object $properties | Where-Object UserID -EQ $user
        if ($winEvent -ne $null)
        {
            $events += $winEvent
        }
    }
}

#send email with list of files deployed
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$events | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw

Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $title -Body $mailBody -BodyAsHtml  
