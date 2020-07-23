

 $vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\Server-Reboot.txt'
 
Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down "
    continue
    
    }

         
                  else 
                     
                    {
              
                    $reboot = Shutdown /r /m \\$vm /t 30 /c "Shutting down to apply Windows Updates" /f     
            
                    if ($reboot -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Server Rebooting "

                     }
                     
                               
                    }     
         
                    }
         
   

         