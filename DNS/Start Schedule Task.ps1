

cls
$vms = Get-Content 'd:\Script\Dev\Severs.txt'

Foreach ($vm in $vms)



#{
 #SCHTASKS /Create /f /s $vm  /RU EDD_DOMAIN\l8dmsinfosvcpp /RP Password1 /TN Start-Services /tr 'D:\Scripts\Start Services.ps1' /SC Once /st 17:00
 #Write-Host $vm.name
#}



                     #############Unmark One Line at a Time and Run Each Line Once############
								 
####################################This line Schedules the  job############################################

{
 SCHTASKS /Create /f /s $vm  /RU EDD_DOMAIN\l8dmsinfosvcpp /RP Password1 /TN Start-Services /tr 'C:\Utility\Start-Local Services.ps1' /SC Once /st 17:00
 Write-Host $vm.name
}

#/RU EDD_DOMAIN\L8DMSSCRTPP /RP EDDadmin!1
#####################################This line Starts the Schedules job##############################################

#{
# SCHTASKS /run /s $vm /TN Start-Services
#}


##########################This line Schedules the Make Directory Batch file job###################################

#{
#SCHTASKS /Create /f /s $vm  /RU EDD_DOMAIN\l8dmsinfosvcpp /RP Password1 /TN makedir /tr 'C:\MD.bat' /SC Once /st 17:00
#}


#####################################This line Starts the Schedules job##############################################
#{
# SCHTASKS /run /s $vm /TN makedir 
#}
# /RU EDD_DOMAIN\l8dmsinfosvcpp










##SCHTASKS /create /f /s $vm /RU EDD_DOMAIN\l8dmsinfosvcpp /TN TCPmon2 /XML c:\TCPmon\TCPmon.xml