# 
#
#

$ErrorActionPreference = "silentlycontinue"
$FileName = "\\dgvsopspd01\wwwroot\CstmAppsSrvcs.htm" 
$configs = import-csv "\\dgvsopspd03\scripts\CstmApps\CstmAppsSrvcs.csv"
$Logfilename = "\\DGVSOPSPD01\Logs\CstmAppsSrvcs.txt"
$n = 1 
# First lets create a file, where we will later save the HTML status  
New-Item -ItemType file $FileName -Force 

$date = Get-Date
$date = Get-Date | Out-File $Logfilename -Append -Force

# Function to write the HTML Header to the file  
# Add parameters for message and data needed to email
Function SendMail 
 	{
 	param($from, $to, $Subject, $Body)
	$date = Get-Date
	# Create from/to addresses  
	$from = New-Object System.Net.Mail.MailAddress "CustAppsMonitor@edd.ca.gov"
	$to =   New-Object System.Net.Mail.MailAddress "teaminf@edd.ca.gov"
	# Create Message  
	$message = new-object  System.Net.Mail.MailMessage $from, $to  
	$message.Subject =  $Subject 
	
	$message.Body = $Body
	
	  
	# Set SMTP Server and create SMTP Client  
	$server = "151.143.2.2"  

	$client = new-object system.net.mail.smtpclient $server  
	  
	# Send the message  
	"Sending an e-mail message to {0} by using SMTP host {1} port {2}." -f $to.ToString(), $client.Host, $client.Port  
	try {  
	   $client.Send($message)  
	   "Message to: {0}, from: {1} has beens successfully sent" -f $from, $to  
	}  
	catch {  
	  "Exception caught in CreateTestMessage: {0}" -f $Error.ToString()  
	}
	$message.subject | Out-File $logfilename -Append -Force
	$message.body | Out-File $logfilename -Append -Force
	}
Function writeHtmlHeader  
{  
param($fileName)  
$date = Get-Date  
Add-Content $fileName "<html>" 
Add-Content $fileName "<head>" 
Add-Content $fileName "<META HTTP-EQUIV=refresh CONTENT=15>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $fileName '<title>Custom Application Services Monitor - Goethe</title>' 
add-content $fileName '<STYLE TYPE="text/css">' 
add-content $fileName  "<!--" 
add-content $fileName  "td {" 
add-content $fileName  "font-family: Tahoma;" 
add-content $fileName  "font-size: 11px;" 
add-content $fileName  "border-top: 1px solid #999999;" 
add-content $fileName  "border-right: 1px solid #999999;" 
add-content $fileName  "border-bottom: 1px solid #999999;" 
add-content $fileName  "border-left: 1px solid #999999;" 
add-content $fileName  "padding-top: 0px;" 
add-content $fileName  "padding-right: 0px;" 
add-content $fileName  "padding-bottom: 0px;" 
add-content $fileName  "padding-left: 0px;" 
add-content $fileName  "}" 
add-content $fileName  "body {" 
add-content $fileName  "margin-left: 5px;" 
add-content $fileName  "margin-top: 5px;" 
add-content $fileName  "margin-right: 0px;" 
add-content $fileName  "margin-bottom: 10px;" 
add-content $fileName  "" 
add-content $fileName  "table {" 
add-content $fileName  "border: thin solid #000000;" 
add-content $fileName  "}" 
add-content $fileName  "-->" 
add-content $fileName  "</style>" 
Add-Content $fileName "</head>" 
Add-Content $fileName "<body>" 
    
add-content $fileName  "<table width='100%'>" 
add-content $fileName  "<tr bgcolor='#D0E6FF'>" 
Add-Content $FileName  "<a href='http://dgvsopspd01/default.aspx'>Back</a></td>"
add-content $fileName  "<td colspan='7' height='25' align='center'>" 
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Custom Application Services - Goethe - $date</strong></font>" 
add-content $fileName  "</td>" 
add-content $fileName  "</tr>" 
add-content $fileName  "</table>" 
   
 }  
    
 # Function to write the HTML Header to the file  
 Function writeTableHeader  
 {  
 param($fileName)  
    
 Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
 Add-Content $fileName "<td width='25%' align='center'><strong>Server Name</strong></td>" 
 Add-Content $fileName "<td width='25%' align='center'><strong>Service</strong></td>" 
 Add-Content $fileName "<td width='25%' align='center'><strong>Status</strong></td>" 
 Add-Content $fileName "</tr>" 
 }  
    
 Function writeHtmlFooter  
 {  
 param($fileName) 
 Add-Content $FileName "<table width='100%'><tbody>" 
 Add-Content $FileName "<tr bgcolor='#D0E6FF'>" 
 Add-Content $FileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>End Report</strong></font></td>" 
 Add-Content $FileName "</tr>"    
 Add-Content $fileName "</body>" 
 Add-Content $fileName "</html>" 
 }  
    
    
 writeHtmlHeader $FileName 
 
 
  Add-Content $FileName "<table width='100%'><tbody>" 
  Add-Content $FileName "<tr bgcolor='#D0E6FF'>" 
  Add-Content $FileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>  </strong></font></td>" 
  Add-Content $FileName "</tr>" 
 
  writeTableHeader $FileName 


# This is the key part
##############################################################

#$j = ((Get-Content "\\dgvsopspd02\scripts\CstmApps\CstmAppsSrvcs.csv").count)
$services = Import-Csv  "\\dgvsopspd02\scripts\CstmApps\CstmAppsSrvcs.csv"
$i = 0

foreach($service in $services)
#while($i -lt ($j))
{
	$name = $service.ServiceName
	$server = $service.MachineName
	$svc = get-Service -ComputerName $server -Name "$name"
	
	if ($svc.Status -notlike "Running") 
	{
	
	 	$server = $server.ToUpper()
 		$service = $svc.DisplayName
 		$Status = $svc.Status
		
 
 		Add-Content $fileName "<tr bgcolor=#FF0000>" 
 		Add-Content $fileName "<td width='25%' align='center'><strong><font color='#FFFF00'>$server</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><strong><font color='#FFFF00'>$service</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><strong><font color='#FFFF00'>$Status</td>" 
 		Add-Content $fileName "</tr>" 

	}
	else
	{
		$server = $server.ToUpper()
 		$service = $svc.DisplayName
 		$Status = $svc.Status
 
 		Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
		Add-Content $fileName "<td width='25%'' align='center'>$server</td>" 
 		Add-Content $fileName "<td width='25%'' align='center'>$service</td>" 
 		Add-Content $fileName "<td width='25%'' align='center'>$Status</td>" 
 		Add-Content $fileName "</tr>" 
		
	
	}
	Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
 	Add-Content $fileName "<td width='25%' align='center'></td>" 
 	Add-Content $fileName "<td width='25%' align='center'></td>" 
 	Add-Content $fileName "<td width='25%' align='center'></td>" 
 	Add-Content $fileName "</tr>" 
	$i++
} 
  
  
writeHtmlFooter $FileName 

exit

