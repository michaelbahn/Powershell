#
#This script runs every 30 minutes (at x:15 and x:45) to check SIDES processing 
#directory for Zip Wrap. If there are files left after procesing hours 
#(from 06:00 to 22:00) the script will turn the HTML RED and ALERT 
#teaminf@edd.ca.gov via email.
#AUTHOR:	Grant Fraser
#Date: 		07/10/2017
#File: 		\\dgvsopspd03\scripts\CstmApps\SIDESCountsHTML.ps1
#Version 	1.0
#

$SidesDir =        "\\dgvmappuipd01\h$\Shares\DMS\DE1101cz\iBatch_Zipwrap\Input\"
$nonftpfile = ""

$fileName =  "\\dgvsopspd01\wwwroot\SIDES.htm"


New-Item -itemtype file -Path $fileName -Force
$date = Get-Date

#================================================================================ 
# Add parameters for message and data needed to email
Function SendMail 
 	{
 	param()
	$date = Get-Date
	# Create from/to addresses  
	$from = New-Object System.Net.Mail.MailAddress "SIDES@edd.ca.gov"  
	$to =   New-Object System.Net.Mail.MailAddress "teaminf@edd.ca.gov"
#	"grant.fraser@edd.ca.gov"
#	"teaminf@edd.ca.gov"
		$MachineName = $svc.MachineName 
		$ServiceName =	$svc.Name 
		$Status = $svc.Status
			$svc.Status
	# Create Message  
	$message = new-object  System.Net.Mail.MailMessage $from, $to  
	$message.Subject = "SIDES ZipWrap Application is down" 
	
	$message.Body = "The zipWrap application in CA-SIDES is not processing. Please refer to http://dgvsopspd01/SIDES.htm" 
		  
	# Set SMTP Server and create SMTP Client  
	$server = "151.143.2.2"  

	$client = new-object system.net.mail.smtpclient $server  
	  
	# Send the message  
	Write-Output "Sending an e-mail message to {0} by using SMTP host {1} port {2}. $client.Host, $client.Port"  
	try {  
	   $client.Send($message)  
	   Write-Output "Message to: {0}, from: {1} has beens successfully sent $from, $to ."  
	}  
	catch {  
	  Write-Output "Exception caught in CreateTestMessage: {0}" -f $Error.ToString()  
	}
	$message.subject 
	#| Out-File $logfilename -Append -Force
	$message.body 
	#| Out-File $logfilename -Append -Force
	}
#================================================================================ 	


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
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>SIDES Processing File Counts</strong></font>   $date" 
add-content $fileName  "</td>" 
add-content $fileName  "</tr>" 
add-content $fileName  "</table>" 
   
 }  
    
 # Function to write the HTML Header to the file  
 Function writeTableHeader  
 {  
 param($fileName)  
    
 Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
 Add-Content $fileName "<td width='25%' align='center'><strong>Directory</strong></td>" 
 Add-Content $fileName "<td width='25%' align='center'><strong>Count</strong></td>" 
 Add-Content $fileName "<td width='25%' align='center'><strong>Link to Directory</strong></td>" 
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
    
    
  

 
 
 
############################################
writehtmlheader $fileName
Add-Content $FileName "<table width='100%'><tbody>" 
Add-Content $FileName "<tr bgcolor='#D0E6FF'>" 
Add-Content $FileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>  </strong></font></td>" 
Add-Content $FileName "</tr>" 
############################################  main() #######################################################
writeTableHeader $fileName  

#Get count of files in directory $sidesdir
$SidesCount = (Get-ChildItem -Path $SidesDir).count

if($SidesCount -le 0)
{
	Write-Output "SIDES Dir count is " $SidesCount
	
	Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
 		Add-Content $fileName "<td width='25%' align='center'>$SidesFile</td>"
 		Add-Content $fileName "<td width='25%' align='center'>$SidesCount</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><a href=file:$SidesDir>Directory</a></td>"
		#$oldftpcount</td>" 																								<a href=file:'$SidesFile'>Directory</a></td>
 		Add-Content $fileName "</tr>" 
		
}
elseif($SidesCount -gt 0)
{
	Write-Output "SIDES Dir more than zero count" $SidesCount 
	
	Add-Content $fileName "<tr bgcolor=#FF0000>" 
 	Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'>$SidesFile</td>" 
 	Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'>$SidesCount</td>" 
 	Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'><a href='file:$SidesDir'>Directory</a></td>" 
 	Add-Content $fileName "</tr>" 
	Sendmail
}
 writeHtmlFooter  $FileName
