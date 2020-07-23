
#################################### Start Infoimage Domain Servers ###################################################################
##$vms = Get-Content 'D:\Scripts\Domain-dev.txt'


$vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\Domain-Prod.txt'

Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down "
    continue
    
    }

         
                  else 
                     
                    {
             
                    $result = schtasks /run /s $vm /tn "StartIIF" 
            
                    if ($result -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Starting Services "

                     }
                     
                               
                    }     
         
                    }
         
         start-sleep -seconds 60

         


#################################### Start Infoimage Image-01 Servers ###################################################################
##$vms = Get-Content 'D:\Scripts\Image01-DEV.txt'

$vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\Image01-Prod.txt'

Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down "
    continue
    
    }

         
                  else 
                     
                    {
             
                    $result = schtasks /run /s $vm /tn "StartIIF" 
            
                    if ($result -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Starting Services "

                     }
                     
                               
                    }     
         
                    }

start-sleep -seconds 60

#################################### Start Infoimage Image-02 Servers ###################################################################

##$vms = Get-Content 'D:\Scripts\Image02-DEV.txt'

$vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\Image02-Prod.txt'

Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down "
    continue
    
    }

         
                  else 
                     
                    {
             
                    $result = schtasks /run /s $vm /tn "StartIIF" 
            
                    if ($result -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Starting Services "

                     }
                     
                               
                    }     
         
                    }


start-sleep -seconds 60
#################################### Start Infoimage Image-03 Servers ###################################################################
##$vms = Get-Content 'D:\Scripts\Image03-DEV.txt'

$vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\Image03-Prod.txt'


Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down " -ForegroundColor Red
    continue
    
    }

         
                  else 
                     
                    {
             
                    $result = schtasks /run /s $vm /tn "StartIIF" 
            
                    if ($result -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Starting Services " -ForegroundColor Green

                     }
                     
                               
                    }     
         
                    }
start-sleep -seconds 60

#################################### Start Infoimage Image-04 Servers ###################################################################
##$vms = Get-Content 'D:\Scripts\Image04-DEV.txt'
 
$vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\Image04-Prod.txt'
 
Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down " -ForegroundColor Red
    
    
    continue
    
    }

         
                  else 
                     
                    {
             
                    $result = schtasks /run /s $vm /tn "StartIIF" 
            
                    if ($result -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Starting Services "

                     }
                     
                               
                    }     
         
                    }
         

start-sleep -seconds 60

#################################### Start Infoimage RBE Servers ###################################################################
##$vms = Get-Content 'D:\Scripts\RBE-DEV.txt'

$vms = Get-Content 'D:\Powershell-Production-scripts\Scripts\RBE-Prod.txt'
  
Foreach ($vm in $vms) {

$PingRequet = Test-Connection -ComputerName  $vm -Count 2  -Quiet

if ($PingRequet -eq $false)
 { 
    Write-Host $vm   " Server Down " -ForegroundColor Red
    
    
    continue
    
    }

         
                  else 
                     
                    {
             
                    $result = schtasks /run /s $vm /tn "StartRBE" 
            
                    if ($result -like "success*" ) 
                 
                    {

                       Write-Host $vm  " Starting Services "

                     }
                     
                               
                    }     
         
                    }
         
