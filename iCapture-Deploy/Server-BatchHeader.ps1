$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$newFileList = @()

#log file settings
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile = Initialize-Log $logPath $title
$sender = "teaminf@edd.ca.gov"
$recipients = Get-Content  recipients.txt

#get list of servers
$servers = Get-Content  ("servers.txt")

foreach ($server in $servers) 
{   
    #Open HKLM registry hive on target workstation
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $server)
    #add value to HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Impression Technology\PRE_PROD50\General App\5.0\BatchHeader\ENTWHITE\document_description\VALUES
    $subKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\PRE_PROD50\General App\5.0\BatchHeader\ENTWHITE\document_description\VALUES", $true)
    $subKey.CreateSubKey("HIRING PACKAGES")
    #add value to HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Impression Technology\PRE_PROD50\General App\5.0\BatchHeader\ENTWHITE\wm_form_id\VALUES
    $subKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\PRE_PROD50\General App\5.0\BatchHeader\ENTWHITE\wm_form_id\VALUES", $true)
    $subKey.CreateSubKey("HIREPACK")
    #add value to HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Impression Technology\PRE_PROD50\General App\5.0\BatchHeader\ENTWHITE\workflow_route\VALUES
    $subKey = $reg.OpenSubKey("SOFTWARE\Wow6432Node\Impression Technology\PRE_PROD50\General App\5.0\BatchHeader\ENTWHITE\workflow_route\VALUES", $true)
    $subKey.CreateSubKey("WSBIDX")

#repeat for non-wow6432 node

}   #end for

#send email with list of files deployed
$tempFile = Join-Path $scriptPath "temp.html"
$htmlHeader = "<style>TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}</style>"
$newFileListSorted = $newFileList | Sort-Object FullName
$newFileListSorted  | Select-Object  Directory, Name, LastWriteTime | ConvertTo-Html  -Head $htmlHeader | Out-File $tempFile
$mailBody = Get-Content $tempFile -Raw
$subject = "iCapture Deployment Completed: $($newFileList.Count) files"    
Send-MailMessage -SmtpServer smtp.edd.ca.gov -From $sender -To $recipients -Subject $subject -Body $mailBody -BodyAsHtml  -Attachments $logFile       

 