try{
    $status = ([wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities").DetermineIfRebootPending()

    $rebootPending = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
    $rebootRequired = Test-Path "HKLM:\SOFTWARE­\Microsoft­\Windows­\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    #$PendingFileRenameOperations = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA Ignore  
        
    if  (($status.RebootPending -ne $null) -or  ($rebootPending -eq $true) -or  ($rebootRequired -eq $true)) 
    {  
        return $true
    } 
    else
    {
        return $false
    } 

} catch{} 