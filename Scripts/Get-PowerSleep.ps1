$plan = Get-WmiObject -Class win32_powerplan -Namespace root\cimv2\power -Filter “isActive=true”   
$regex = [regex]“{(.*?)}$” 
$planGuid = $regex.Match($plan.instanceID.Tostring()).groups[1].value 
$powercfgs = powercfg -query $planGuid SUB_DISK
foreach ($powercfg in $powercfgs)
{
    $index = $powercfg.Indexof("Current AC Power Setting Index:")
    if ($index -ge 0)
    {
        $powerIndexHex = $powercfg.Substring(36)
        $powerIndex = [int] $powerIndexHex
    }
 }

 switch ($powerIndex)
 {
    0	{return "Do nothing"; break;}
    1	{return "Sleep"; break;}
    2	{return "Hibernate"; break;}
    3	{return "Shut down"; break;}
    4	{return "Turn off the display"; break;}
    default	{return "Unknown"; break;}
 }

Index Number	Action

