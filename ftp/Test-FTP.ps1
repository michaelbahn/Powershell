cls
#initialize
$title = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath
$modulePath = "..\Scripts"
$logPath = "..\Logs"
Import-Module (Join-Path $modulePath Utilities.psm1) -Force
$logFile= Initialize-Log $logPath $title

# Load WinSCP .NET assembly
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

$ftpHost = "173.13.187.161"

# Setup session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    #Protocol = [WinSCP.Protocol]::Sftp
    Protocol = [WinSCP.Protocol]::ftp
    HostName = $ftpHost 
    UserName = "caedd"
    Password = "caedd"
    #SshHostKeyFingerprint = "ssh-rsa 2048 xxxxxxxxxxx...="
}
$session = New-Object WinSCP.Session

try
{
    # Connect
    $session.Open($sessionOptions)

    # Upload
    $session.PutFiles("C:\Users\mbahn-22\Documents\test.txt", "/").Check()
    Write-Log $logFile "File transferred to $(ftpHost)"
}
catch
{
    $errorMessage = "Error opening $(ftpHost): $($error[0])"
    Write-Log $logFile $errorMessage
    $session.Dispose()
}

finally
{
    # Disconnect, clean up
    $session.Dispose()
}